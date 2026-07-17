<#
One-time setup for the OpenClerk offline package: creates a self-signed,
per-user localhost certificate, installs the add-in content + server script,
registers a scheduled task to run the server (hidden) at logon, and installs
the manifest into Word's WEF folder.

Runs entirely as the current user -- no admin/UAC prompt. serve-openclerk.ps1
terminates TLS in-process via TcpListener + SslStream rather than
System.Net.HttpListener, so none of the machine-wide `netsh http`
sslcert/urlacl bindings that used to require elevation are needed. The
certificate's private key is created non-exportable in Cert:\CurrentUser\My
and never leaves that per-user store; the public certificate is separately
trusted via Cert:\CurrentUser\Root (also per-user, not machine-wide).

Run this once per machine/user. Re-run any time to rotate the secret, renew
the certificate, or move the install location. Pass -Uninstall to remove
everything this script created.
#>
param(
    [int]$Port = 44399,
    [string]$InstallDir = "",
    [switch]$DryRun,
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'

$certSubject = 'CN=OpenClerkLocalServer'
$taskName = 'OpenClerkLocalServer'
$wefTarget = if ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\WEF' } else { $null }
$wefManifestPath = if ($wefTarget) { Join-Path $wefTarget 'openclerk-manifest-local.xml' } else { $null }

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

# --- Uninstall ----------------------------------------------------------
if ($Uninstall) {
    if ($DryRun) {
        Write-Host "Dry run enabled. Would stop/remove scheduled task '$taskName', remove certs with subject '$certSubject' from CurrentUser\My and CurrentUser\Root, remove '$InstallDir', and remove '$wefManifestPath'."
        exit 0
    }

    Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue | ForEach-Object {
        Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "Removed scheduled task '$taskName'."
    }

    foreach ($storeLocation in 'My', 'Root') {
        Get-ChildItem "Cert:\CurrentUser\$storeLocation" -ErrorAction SilentlyContinue |
            Where-Object { $_.Subject -eq $certSubject } |
            ForEach-Object {
                Remove-Item -LiteralPath "Cert:\CurrentUser\$storeLocation\$($_.Thumbprint)" -Force
                Write-Host "Removed cert $($_.Thumbprint) from CurrentUser\$storeLocation."
            }
    }

    if (Test-Path $InstallDir) {
        Remove-Item -Path $InstallDir -Recurse -Force
        Write-Host "Removed install directory: $InstallDir"
    }
    if ($wefManifestPath -and (Test-Path $wefManifestPath)) {
        Remove-Item -Path $wefManifestPath -Force
        Write-Host "Removed WEF manifest: $wefManifestPath"
    }

    Write-Host 'Uninstall complete.'
    exit 0
}

# --- Per-user certificate (non-exportable key, no admin required) -------
function Install-LocalhostCert {
    # Reuse an existing cert with reasonable remaining life instead of minting a new one on every
    # run, so re-running this script to rotate the secret or move the install dir doesn't churn
    # trust anchors or force the taskpane to re-warm a new TLS session unnecessarily.
    $cert = Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue |
        Where-Object { $_.Subject -eq $certSubject -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
        Select-Object -First 1

    if (-not $cert) {
        $cert = New-SelfSignedCertificate -DnsName 'localhost' -Subject $certSubject `
            -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(2) `
            -KeyUsage DigitalSignature, KeyEncipherment -KeyExportPolicy NonExportable `
            -TextExtension @('2.5.29.37={text}1.3.6.1.5.5.7.3.1') # EKU: Server Authentication only
    }

    $rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root', 'CurrentUser')
    $rootStore.Open('ReadWrite')
    if (-not ($rootStore.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint })) {
        # Import only the public certificate -- the non-exportable private key stays in My and is
        # never touched by this step.
        $rootStore.Add([System.Security.Cryptography.X509Certificates.X509Certificate2]::new($cert.RawData))
    }
    $rootStore.Close()

    return $cert
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

# Cryptographically random 256-bit secret (base64url, no padding) -- stronger than a v4 GUID's
# 122 bits and matches the entropy the query-string/cookie auth check deserves.
$secretBytes = [byte[]]::new(32)
[System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($secretBytes)
$secret = [Convert]::ToBase64String($secretBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')

$appDir = Join-Path $InstallDir 'app'
$serverScriptDest = Join-Path $InstallDir 'serve-openclerk.ps1'
$secretFilePath = Join-Path $InstallDir 'secret.key'

Write-Host "Install dir:  $InstallDir"
Write-Host "Port:         $Port"

if ($DryRun) {
    Write-Host 'Dry run enabled. No files were copied, no certificate/scheduled task changes were made.'
    exit 0
}

$cert = Install-LocalhostCert
Write-Host "Certificate ready: $($cert.Thumbprint) (CurrentUser\My, non-exportable; trusted via CurrentUser\Root)."

# --- Content, manifest, scheduled task -----------------------------------
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
    "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$serverScriptDest`" -Port $Port -SecretFile `"$secretFilePath`" -ContentRoot `"$appDir`" -CertThumbprint `"$($cert.Thumbprint)`""
$taskTrigger = New-ScheduledTaskTrigger -AtLogOn
# RunLevel Limited is the scheduled-task equivalent of "standard user rights" -- explicit here so
# it's clear (and enforced) that nothing about running this task ever needs elevation.
$taskPrincipal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
# Task Scheduler kills any task after 72 hours by default (ExecutionTimeLimit) and does not
# restart it on failure unless told to -- both would silently leave OpenClerk's local server
# dead (72h is well within a normal multi-day logged-in session) with no recovery until the
# next login. RestartCount/-Interval covers crashes; ExecutionTimeLimit Zero means "no limit".
$taskSettings = New-ScheduledTaskSettingsSet -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) `
    -ExecutionTimeLimit ([TimeSpan]::Zero) -DontStopOnIdleEnd
Register-ScheduledTask -TaskName $taskName -Action $taskAction -Trigger $taskTrigger -Principal $taskPrincipal -Settings $taskSettings `
    -Description 'Serves the OpenClerk add-in locally on 127.0.0.1 for offline use.' -Force | Out-Null

Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
Start-ScheduledTask -TaskName $taskName

New-Item -ItemType Directory -Path $wefTarget -Force | Out-Null
Copy-Item -Path $manifestOutPath -Destination $wefManifestPath -Force

Write-Host ''
Write-Host 'OpenClerk local server is set up and running (no admin rights were used).'
Write-Host 'It will also start automatically the next time you log in.'
Write-Host 'Restart Word to see the OpenClerk button on the Home ribbon.'
Write-Host "To remove everything: powershell -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`" -Uninstall"
