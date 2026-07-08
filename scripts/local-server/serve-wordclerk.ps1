<#
Minimal static-file HTTPS server for the WordClerk offline package.

Serves $ContentRoot on https://127.0.0.1:$Port/ only -- it never binds to any
other interface or port, so nothing on the network (or even other ports on
this machine) can reach it.

Access control: since loopback sockets aren't restricted to a single
application on Windows, requests must prove they know the secret (read from
-SecretFile at startup, never passed as a command-line argument -- Scheduled
Task definitions are readable by any process running as the same user via
`schtasks /query /fo LIST /v`, which would otherwise leak it). The first
request must include it as a query string (?k=<secret>), which sets an
HttpOnly session cookie; every subsequent request (including the page's own
JS/CSS/image fetches) is authorized via that cookie instead. Requests with
neither a valid ?k= nor a valid cookie get 403.

Requires the certificate + URL ACL for $Port to already be bound via
setup-local-server.ps1 (that part needs admin once; this script does not).
#>
param(
    [Parameter(Mandatory = $true)][int]$Port,
    [Parameter(Mandatory = $true)][string]$SecretFile,
    [Parameter(Mandatory = $true)][string]$ContentRoot
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $ContentRoot)) {
    Write-Error "Content root not found: $ContentRoot"
    exit 1
}
if (-not (Test-Path $SecretFile)) {
    Write-Error "Secret file not found: $SecretFile"
    exit 1
}
$Secret = (Get-Content $SecretFile -Raw).Trim()
$ContentRoot = (Resolve-Path $ContentRoot).Path
$contentRootWithSep = $ContentRoot.TrimEnd('\') + '\'
$cookieName = 'wc_auth'

$mimeTypes = @{
    '.html' = 'text/html; charset=utf-8'
    '.js'   = 'application/javascript; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'
    '.xml'  = 'application/xml; charset=utf-8'
    '.png'  = 'image/png'
    '.jpg'  = 'image/jpeg'
    '.svg'  = 'image/svg+xml'
    '.ico'  = 'image/x-icon'
    '.map'  = 'application/json; charset=utf-8'
}

function Test-Authorized($request) {
    $queryKey = $request.QueryString['k']
    if ($queryKey -and $queryKey -eq $Secret) {
        return $true
    }
    $cookie = $request.Cookies[$cookieName]
    if ($cookie -and $cookie.Value -eq $Secret) {
        return $true
    }
    return $false
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("https://127.0.0.1:$Port/")
$listener.Start()
Write-Host "WordClerk local server listening on https://127.0.0.1:$Port/ (content root: $ContentRoot)"

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        try {
            if (-not (Test-Authorized $request)) {
                $response.StatusCode = 403
                $response.Close()
                continue
            }

            if ($request.QueryString['k'] -eq $Secret) {
                $authCookie = New-Object System.Net.Cookie($cookieName, $Secret, '/')
                $authCookie.HttpOnly = $true
                $authCookie.Secure = $true
                $response.Cookies.Add($authCookie)
            }

            $relativePath = [Uri]::UnescapeDataString($request.Url.AbsolutePath).TrimStart('/')
            if ([string]::IsNullOrWhiteSpace($relativePath)) {
                $relativePath = 'taskpane.html'
            }

            $filePath = Join-Path $ContentRoot $relativePath
            $resolvedFilePath = [System.IO.Path]::GetFullPath($filePath)

            # Path-traversal guard: resolved path must stay inside ContentRoot.
            if (-not $resolvedFilePath.StartsWith($contentRootWithSep, [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path $resolvedFilePath -PathType Leaf)) {
                $response.StatusCode = 404
                $response.Close()
                continue
            }

            $ext = [System.IO.Path]::GetExtension($resolvedFilePath).ToLowerInvariant()
            $contentType = $mimeTypes[$ext]
            if (-not $contentType) {
                $contentType = 'application/octet-stream'
            }
            $response.ContentType = $contentType

            $bytes = [System.IO.File]::ReadAllBytes($resolvedFilePath)
            $response.ContentLength64 = $bytes.Length
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
            $response.Close()
        } catch {
            try {
                $response.StatusCode = 500
                $response.Close()
            } catch {}
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
