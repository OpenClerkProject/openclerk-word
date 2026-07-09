param(
    [string]$ManifestPath = "manifest.xml",
    [string]$Target = "",
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

# $scriptDir is the installer/ subfolder inside the extracted release package.
# The manifest lives one level up, at the package root.
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageRoot = Split-Path -Parent $scriptDir

if ([string]::IsNullOrWhiteSpace($Target)) {
    if ($env:LOCALAPPDATA) {
        $Target = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\WEF'
    } else {
        $Target = Join-Path $packageRoot '.installer-test\wef'
    }
}

$manifestFullPath = if ([System.IO.Path]::IsPathRooted($ManifestPath)) {
    $ManifestPath
} else {
    Join-Path $packageRoot $ManifestPath
}

if (-not (Test-Path $manifestFullPath)) {
    Write-Error "Manifest not found: $manifestFullPath"
    exit 1
}

$targetManifest = Join-Path $Target 'openclerk-manifest.xml'

Write-Host "Source manifest: $manifestFullPath"
Write-Host "Install target:  $targetManifest"

if ($DryRun) {
    Write-Host 'Dry run enabled. No files were copied.'
} else {
    New-Item -ItemType Directory -Path $Target -Force | Out-Null
    Copy-Item -Path $manifestFullPath -Destination $targetManifest -Force
    Write-Host 'Manifest installed successfully.'
}

Write-Host 'Done.'
