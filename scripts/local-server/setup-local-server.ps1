<#
One-time setup for the OpenClerk offline package: binds a self-signed
localhost certificate to a single loopback port, installs the add-in content
+ server script, registers a scheduled task to run the server (hidden) at
logon, and installs the manifest into Word's WEF folder.

The certificate bind + URL ACL steps touch machine-wide HTTP/TLS
configuration and require an elevated (admin) PowerShell session; this
script will prompt for elevation via UAC for just that portion if it isn't
already running elevated (via -ElevatedCertStep, an internal re-entry point
-- don't pass it directly). Nothing else here requires admin rights.

Run this once per machine. Re-run any time to rotate the secret or move the
install location.
#>
param(
    [int]$Port = 44399,
    [string]$InstallDir = "",
    [switch]$DryRun,
    [switch]$ElevatedCertStep
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($InstallDir)) {
    # $env:LOCALAPPDATA is Windows-only; this script is Windows-only in real use, but the
    # -DryRun path is also exercised by CI smoke tests running cross-platform pwsh, so fall
    # back to a harmless path instead of failing to even parse arguments there.
    $InstallDir = if ($env:LOCALAPPDATA) {
        Join-Path $env:LOCALAPPDATA 'OpenClerk\LocalServer'
    } else {
        Join-Path ([System.IO.Path]::GetTempPath()) 'OpenClerk\LocalServer'
    }
}

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Install-LocalhostCertBinding([int]$Port) {
    $cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -eq 'CN=OpenClerkLocalServer' } | Select-Object -First 1
    if (-not $cert) {
        $cert = New-SelfSignedCertificate -DnsName 'localhost' -Subject 'CN=OpenClerkLocalServer' `
            -CertStoreLocation Cert:\LocalMachine\My -NotAfter (Get-Date).AddYears(5) -KeyUsage DigitalSignature, KeyEncipherment
    }

    $rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root', 'LocalMachine')
    $rootStore.Open('ReadWrite')
    if (-not ($rootStore.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint })) {
        $rootStore.Add($cert)
    }
    $rootStore.Close()

    netsh http delete sslcert ipport=127.0.0.1:$Port 2>$null | Out-Null
    netsh http delete urlacl url="https://127.0.0.1:$Port/" 2>$null | Out-Null

    $appId = '{6f6a6e7e-6f9e-4e8f-9a6f-77c1a2c0b001}'
    netsh http add urlacl url="https://127.0.0.1:$Port/" user="$env:USERDOMAIN\$env:USERNAME" | Out-Null
    netsh http add sslcert ipport=127.0.0.1:$Port certhash=$($cert.Thumbprint) appid="$appId" | Out-Null
}

# --- Internal re-entry point: runs only the elevated cert/URL-ACL step, then exits. ---
if ($ElevatedCertStep) {
    Install-LocalhostCertBinding -Port $Port
    exit 0
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageRoot = Split-Path -Parent $scriptDir
$appSourceDir = Join-Path $packageRoot 'app'
$manifestTemplatePath = Join-Path $packageRoot 'manifest.xml'

if (-not (Test-Path $appSourceDir)) {
    Write-Error "Add-in content not found: $appSourceDir"
    exit 1
}
if (-not (Test-Path $manifestTemplatePath)) {
    Write-Error "Manifest template not found: $manifestTemplatePath"
    exit 1
}

$secret = [guid]::NewGuid().ToString('N')
$appDir = Join-Path $InstallDir 'app'
$serverScriptDest = Join-Path $InstallDir 'serve-openclerk.ps1'
$secretFilePath = Join-Path $InstallDir 'secret.key'

Write-Host "Install dir:  $InstallDir"
Write-Host "Port:         $Port"

if ($DryRun) {
    Write-Host 'Dry run enabled. No files were copied, no certificate/URL ACL/scheduled task changes were made.'
    exit 0
}

# --- Elevated portion: certificate + URL ACL binding for the port -----------
if (Test-IsAdmin) {
    Install-LocalhostCertBinding -Port $Port
} else {
    Write-Host "Requesting administrator rights to bind a localhost certificate to port $Port..."
    $elevatedArgs = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$($MyInvocation.MyCommand.Path)`"",
        '-Port', $Port, '-ElevatedCertStep'
    )
    Start-Process powershell -Verb RunAs -ArgumentList $elevatedArgs -Wait
}
Write-Host 'Certificate bound and trusted for 127.0.0.1.'

# --- Non-elevated portion: content, manifest, scheduled task ----------------
New-Item -ItemType Directory -Path $appDir -Force | Out-Null
Copy-Item -Path (Join-Path $appSourceDir '*') -Destination $appDir -Recurse -Force
Copy-Item -Path (Join-Path $scriptDir 'serve-openclerk.ps1') -Destination $serverScriptDest -Force

$manifestContent = Get-Content $manifestTemplatePath -Raw
$manifestContent = $manifestContent.Replace('{{PORT}}', $Port).Replace('{{SECRET}}', $secret)
$manifestOutPath = Join-Path $InstallDir 'manifest.xml'
Set-Content -Path $manifestOutPath -Value $manifestContent -NoNewline

# Write the secret to its own file instead of passing it as a Scheduled Task argument -- task
# definitions are readable by any process running as the same user (e.g. `schtasks /query /fo
# LIST /v` or WMI), which would otherwise leak the secret to anything running under this
# account. Restrict the file's ACL to the installing user only, stripping inherited/Everyone
# entries, so even other processes on this account can't casually read it off disk.
Set-Content -Path $secretFilePath -Value $secret -NoNewline
$secretAcl = Get-Acl $secretFilePath
$secretAcl.SetAccessRuleProtection($true, $false)
foreach ($rule in @($secretAcl.Access)) {
    $secretAcl.RemoveAccessRule($rule) | Out-Null
}
$ownerRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    "$env:USERDOMAIN\$env:USERNAME", 'Read,Write', 'Allow'
)
$secretAcl.AddAccessRule($ownerRule)
Set-Acl -Path $secretFilePath -AclObject $secretAcl

$taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument `
    "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$serverScriptDest`" -Port $Port -SecretFile `"$secretFilePath`" -ContentRoot `"$appDir`""
$taskTrigger = New-ScheduledTaskTrigger -AtLogOn
# Task Scheduler kills any task after 72 hours by default (ExecutionTimeLimit) and does not
# restart it on failure unless told to -- both would silently leave OpenClerk's local server
# dead (72h is well within a normal multi-day logged-in session) with no recovery until the
# next login. RestartCount/-Interval covers crashes; ExecutionTimeLimit Zero means "no limit".
$taskSettings = New-ScheduledTaskSettingsSet -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) `
    -ExecutionTimeLimit ([TimeSpan]::Zero) -DontStopOnIdleEnd
Register-ScheduledTask -TaskName 'OpenClerkLocalServer' -Action $taskAction -Trigger $taskTrigger -Settings $taskSettings `
    -Description 'Serves the OpenClerk add-in locally on 127.0.0.1 for offline use.' -Force | Out-Null

Stop-ScheduledTask -TaskName 'OpenClerkLocalServer' -ErrorAction SilentlyContinue
Start-ScheduledTask -TaskName 'OpenClerkLocalServer'

$wefTarget = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\WEF'
New-Item -ItemType Directory -Path $wefTarget -Force | Out-Null
Copy-Item -Path $manifestOutPath -Destination (Join-Path $wefTarget 'openclerk-manifest-local.xml') -Force

Write-Host ''
Write-Host 'OpenClerk local server is set up and running.'
Write-Host 'It will also start automatically the next time you log in.'
Write-Host 'Restart Word to see the OpenClerk button on the Home ribbon.'
