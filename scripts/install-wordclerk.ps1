param(
    [string]$ManifestPath = "manifest.xml",
    [string]$Target = "",
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir

if ([string]::IsNullOrWhiteSpace($Target)) {
    if ($env:LOCALAPPDATA) {
        $Target = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\WEF'
    } else {
        $Target = Join-Path $repoRoot '.installer-test\wef'
    }
}

$manifestFullPath = if ([System.IO.Path]::IsPathRooted($ManifestPath)) {
    $ManifestPath
} else {
    Join-Path $repoRoot $ManifestPath
}

$nodeScript = Join-Path $scriptDir 'install-wordclerk.js'

$dryRunArg = if ($DryRun) { '--dry-run' } else { '' }

Write-Host "Installing WordClerk manifest..."
node $nodeScript --manifest $manifestFullPath --target $Target $dryRunArg

if ($LASTEXITCODE -ne 0) {
    throw "Installer failed with exit code $LASTEXITCODE"
}

Write-Host "Done."
