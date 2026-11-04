# Xuất CA gốc mà máy này đang thấy khi ra HTTPS, sang .local/machine-ca.pem (không vào git, không vào image).
# Máy có phần mềm quét HTTPS (ở máy phát triển là Norton) thì cần file này để build Docker mà không phải tắt xác minh TLS.
# Máy không bị quét thì file rỗng là đủ. Đây là chứng chỉ công khai, không phải bí mật.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$dir = Join-Path $root '.local'
New-Item -ItemType Directory -Force $dir | Out-Null
$out = Join-Path $dir 'machine-ca.pem'

$tcp = New-Object Net.Sockets.TcpClient('pypi.org', 443)
$script:issuer = $null
$cb = [Net.Security.RemoteCertificateValidationCallback] { param($s, $cert, $chain, $errors) $script:issuer = $cert.Issuer; $true }
$tls = New-Object Net.Security.SslStream($tcp.GetStream(), $false, $cb)
$tls.AuthenticateAsClient('pypi.org')
$tls.Dispose(); $tcp.Dispose()

$found = Get-ChildItem Cert:\CurrentUser\Root, Cert:\LocalMachine\Root | Where-Object { $_.Subject -eq $script:issuer } | Select-Object -First 1
if ($null -eq $found -or $script:issuer -notmatch 'Antivirus|Shield|Scanning|Proxy|Inspection') {
    Set-Content -Path $out -Value '' -NoNewline -Encoding ascii
    Write-Output "Không thấy CA quét HTTPS (nhà cấp: $script:issuer). Đã ghi file rỗng."
    exit 0
}
$b64 = [Convert]::ToBase64String($found.RawData, 'InsertLineBreaks')
Set-Content -Path $out -Value "-----BEGIN CERTIFICATE-----`n$b64`n-----END CERTIFICATE-----`n" -Encoding ascii
Write-Output "Đã xuất: $($found.Subject)"
Write-Output "SHA-1: $($found.Thumbprint)"
