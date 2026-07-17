<#
Minimal static-file HTTPS server for the OpenClerk offline package.

Serves $ContentRoot on https://127.0.0.1:$Port/ only -- it binds a
TcpListener to the loopback address and terminates TLS in-process with an
SslStream, so nothing on the network (or even other interfaces on this
machine) can reach it. Terminating TLS in-process (rather than via
System.Net.HttpListener / HTTP.sys) is deliberate: it means the server
needs no machine-wide `netsh http` sslcert/urlacl binding, and its
certificate's private key can live in the per-user store
(Cert:\CurrentUser\My) as a non-exportable key rather than in the machine
store where HTTP.sys would require it. See setup-local-server.ps1.

Access control: since loopback sockets aren't restricted to a single
application on Windows, requests must prove they know the secret (read from
-SecretFile at startup, never passed as a command-line argument -- Scheduled
Task definitions are readable by any process running as the same user via
`schtasks /query /fo LIST /v`, which would otherwise leak it). The first
request must include it as a query string (?k=<secret>), which sets an
HttpOnly session cookie; every subsequent request (including the page's own
JS/CSS/image fetches) is authorized via that cookie instead. Requests with
neither a valid ?k= nor a valid cookie get 403.

The TLS certificate is looked up by -CertThumbprint in Cert:\CurrentUser\My;
setup-local-server.ps1 creates it there (non-exportable) and trusts it via
the per-user Root store. No admin rights are needed by either script.
#>
param(
    [Parameter(Mandatory = $true)][int]$Port,
    [Parameter(Mandatory = $true)][string]$SecretFile,
    [Parameter(Mandatory = $true)][string]$ContentRoot,
    [Parameter(Mandatory = $true)][string]$CertThumbprint
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

$serverCert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $CertThumbprint } | Select-Object -First 1
if (-not $serverCert) {
    Write-Error "Server certificate with thumbprint $CertThumbprint not found in Cert:\CurrentUser\My. Re-run setup-local-server.ps1."
    exit 1
}

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

# Same policy as the <meta> CSP tags baked into taskpane.html/commands.html, plus
# frame-ancestors 'none' (only enforceable via header, not <meta>) to block this content from
# being embedded in an iframe on another origin.
$contentSecurityPolicy = "default-src 'none'; script-src 'self' https://appsforoffice.microsoft.com; " +
    "style-src 'self' https://res-1.cdn.office.net; font-src https://res-1.cdn.office.net; " +
    "img-src 'self' data:; connect-src 'none'; object-src 'none'; frame-src 'none'; " +
    "frame-ancestors 'none'; base-uri 'none';"

# Cap on the request header block we're willing to buffer before giving up -- a static GET server
# never needs more, and it bounds how much an unauthenticated caller can make us hold in memory.
$maxHeaderBytes = 16384

function Test-Authorized($queryKey, $cookieValue) {
    if ($queryKey -and $queryKey -eq $Secret) {
        return $true
    }
    if ($cookieValue -and $cookieValue -eq $Secret) {
        return $true
    }
    return $false
}

# Reads the request line + header block (everything up to the CRLFCRLF terminator) from the TLS
# stream, one small chunk at a time, and returns it as an ASCII string. A GET request has no body,
# so the headers are all we need. Returns $null if the client sends more than $maxHeaderBytes
# without terminating (malformed / abusive) or closes early.
function Read-RequestHeaders($stream) {
    $buffer = New-Object System.IO.MemoryStream
    $chunk = New-Object 'byte[]' 1024
    while ($buffer.Length -le $maxHeaderBytes) {
        $read = $stream.Read($chunk, 0, $chunk.Length)
        if ($read -le 0) {
            return $null
        }
        $buffer.Write($chunk, 0, $read)
        $bytes = $buffer.ToArray()
        # Look for the end-of-headers marker (\r\n\r\n).
        for ($i = 3; $i -lt $bytes.Length; $i++) {
            if ($bytes[$i] -eq 10 -and $bytes[$i - 1] -eq 13 -and $bytes[$i - 2] -eq 10 -and $bytes[$i - 3] -eq 13) {
                return [System.Text.Encoding]::ASCII.GetString($bytes, 0, $i + 1)
            }
        }
    }
    return $null
}

# Parses "?k=<secret>" out of a raw query string (URL-decoded), or $null if absent.
function Get-QueryKey($queryString) {
    if (-not $queryString) { return $null }
    foreach ($pair in $queryString.Split('&')) {
        $eq = $pair.IndexOf('=')
        if ($eq -ge 0 -and $pair.Substring(0, $eq) -eq 'k') {
            return [Uri]::UnescapeDataString($pair.Substring($eq + 1))
        }
    }
    return $null
}

# Pulls the wc_auth cookie value out of a Cookie header line, or $null if absent.
function Get-CookieValue($cookieHeader) {
    if (-not $cookieHeader) { return $null }
    foreach ($pair in $cookieHeader.Split(';')) {
        $trimmed = $pair.Trim()
        $eq = $trimmed.IndexOf('=')
        if ($eq -ge 0 -and $trimmed.Substring(0, $eq) -eq $cookieName) {
            return $trimmed.Substring($eq + 1)
        }
    }
    return $null
}

function Write-Response($stream, [int]$status, [string]$statusText, [hashtable]$headers, [byte[]]$body) {
    if (-not $body) { $body = New-Object 'byte[]' 0 }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("HTTP/1.1 $status $statusText`r`n")
    [void]$sb.Append("Content-Length: $($body.Length)`r`n")
    [void]$sb.Append("Connection: close`r`n")
    if ($headers) {
        foreach ($name in $headers.Keys) {
            [void]$sb.Append("$name`: $($headers[$name])`r`n")
        }
    }
    [void]$sb.Append("`r`n")
    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($sb.ToString())
    $stream.Write($headerBytes, 0, $headerBytes.Length)
    if ($body.Length -gt 0) {
        $stream.Write($body, 0, $body.Length)
    }
    $stream.Flush()
}

$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
$listener.Start()
Write-Host "OpenClerk local server listening on https://127.0.0.1:$Port/ (content root: $ContentRoot)"

try {
    while ($true) {
        $client = $listener.AcceptTcpClient()
        $sslStream = $null
        try {
            $netStream = $client.GetStream()
            $sslStream = New-Object System.Net.Security.SslStream($netStream, $false)
            # Fixed TLS 1.2: this server runs under Windows PowerShell 5.1 / .NET Framework (the
            # scheduled task launches powershell.exe), whose SslStream doesn't support TLS 1.3.
            $sslStream.AuthenticateAsServer($serverCert, $false, [System.Security.Authentication.SslProtocols]::Tls12, $false)
            $sslStream.ReadTimeout = 10000
            $sslStream.WriteTimeout = 10000

            $headerText = Read-RequestHeaders $sslStream
            if (-not $headerText) {
                Write-Response $sslStream 400 'Bad Request' $null $null
                continue
            }

            $lines = $headerText -split "`r`n"
            $requestLine = $lines[0]
            $parts = $requestLine.Split(' ')
            if ($parts.Length -lt 2) {
                Write-Response $sslStream 400 'Bad Request' $null $null
                continue
            }
            $method = $parts[0]
            $target = $parts[1]

            # Only GET is served; anything else is rejected before touching the filesystem.
            if ($method -ne 'GET') {
                Write-Response $sslStream 405 'Method Not Allowed' @{ 'Allow' = 'GET' } $null
                continue
            }

            $queryString = ''
            $rawPath = $target
            $q = $target.IndexOf('?')
            if ($q -ge 0) {
                $rawPath = $target.Substring(0, $q)
                $queryString = $target.Substring($q + 1)
            }

            $cookieHeader = $null
            foreach ($line in $lines[1..($lines.Length - 1)]) {
                $colon = $line.IndexOf(':')
                if ($colon -gt 0 -and $line.Substring(0, $colon).Trim().ToLowerInvariant() -eq 'cookie') {
                    $cookieHeader = $line.Substring($colon + 1).Trim()
                    break
                }
            }

            $queryKey = Get-QueryKey $queryString
            $cookieValue = Get-CookieValue $cookieHeader

            if (-not (Test-Authorized $queryKey $cookieValue)) {
                Write-Response $sslStream 403 'Forbidden' $null $null
                continue
            }

            $relativePath = [Uri]::UnescapeDataString($rawPath).TrimStart('/')
            if ([string]::IsNullOrWhiteSpace($relativePath)) {
                $relativePath = 'taskpane.html'
            }

            $filePath = Join-Path $ContentRoot $relativePath
            $resolvedFilePath = [System.IO.Path]::GetFullPath($filePath)

            # Path-traversal guard: resolved path must stay inside ContentRoot.
            if (-not $resolvedFilePath.StartsWith($contentRootWithSep, [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path $resolvedFilePath -PathType Leaf)) {
                Write-Response $sslStream 404 'Not Found' $null $null
                continue
            }

            $ext = [System.IO.Path]::GetExtension($resolvedFilePath).ToLowerInvariant()
            $contentType = $mimeTypes[$ext]
            if (-not $contentType) {
                $contentType = 'application/octet-stream'
            }

            $responseHeaders = @{ 'Content-Type' = $contentType }
            # On the first (query-string-authorized) request, hand back an HttpOnly+Secure cookie so
            # the page's own subresource fetches don't need to carry the secret in their URLs.
            if ($queryKey -eq $Secret) {
                $responseHeaders['Set-Cookie'] = "$cookieName=$Secret; Path=/; HttpOnly; Secure; SameSite=Strict"
            }
            if ($ext -eq '.html') {
                $responseHeaders['Content-Security-Policy'] = $contentSecurityPolicy
            }

            $bodyBytes = [System.IO.File]::ReadAllBytes($resolvedFilePath)
            Write-Response $sslStream 200 'OK' $responseHeaders $bodyBytes
        } catch {
            try {
                if ($sslStream) {
                    Write-Response $sslStream 500 'Internal Server Error' $null $null
                }
            } catch {}
        } finally {
            if ($sslStream) { $sslStream.Dispose() }
            $client.Close()
        }
    }
} finally {
    $listener.Stop()
}
