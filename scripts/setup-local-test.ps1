$ErrorActionPreference = 'Stop'

Write-Host 'Setting up OpenClerk local test environment...'

if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
    Write-Error 'npm is not installed or not available in PATH. Install Node.js and try again.'
    exit 1
}

Write-Host 'Installing npm dependencies...'
npm install

Write-Host 'Building the add-in...'
npm run build

Write-Host 'Packaging the add-in...'
npm run package

Write-Host 'Local test environment is ready.'
Write-Host 'Generated package: openclerk-addin.zip'
Write-Host ''
Write-Host 'Next steps:'
Write-Host '  1. Run `npm run start` to sideload the manifest into Word Desktop.'
Write-Host '  2. In Word, go to Insert -> My Add-ins -> Upload My Add-in -> Add from file and select manifest.xml if you want to load it manually.'
