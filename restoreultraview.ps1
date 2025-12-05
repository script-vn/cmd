
# 1) Link công khai + download=1 để buộc tải file
$Url = "https://sasinvn-my.sharepoint.com/:u:/g/personal/nam_tran_sasin_vn/IQCHeBUhSwLgTZiU8kHkYKgkARGmIijmE9MPMclmRhxbJ7k?e=IUiSCt&download=1"

# 2) File tạm
$TempFile = Join-Path $env:TEMP "ConnectionOutHistory.ini"

# 3) Bật TLS 1.2/1.3 (cho chắc chắn khi tải HTTPS)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

# 4) Tải file
try {
    Invoke-WebRequest -Uri $Url -OutFile $TempFile -MaximumRedirection 10 -UseBasicParsing -ErrorAction Stop
#    Write-Host "Đã tải file về: $TempFile"
} catch {
#    Write-Host "Tải file lỗi: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# 5) Kiểm tra kích thước file tải về (tránh ghi đè bằng file rỗng)
$size = (Get-Item $TempFile).Length
if ($size -lt 10) {
#    Write-Host "File tải về có vẻ không hợp lệ (size=$size bytes). Dừng để an toàn." -ForegroundColor Yellow
    exit 1
}

# 6) Đường dẫn đích UltraViewer trong AppData\Roaming (user hiện tại)
$TargetDir  = Join-Path $env:APPDATA "UltraViewer"
$TargetFile = Join-Path $TargetDir "ConnectionOutHistory.ini"

# 7) Tạo thư mục nếu chưa có
if (-not (Test-Path $TargetDir)) {
    New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
}

# 8) Backup file cũ (nếu có)
if (Test-Path $TargetFile) {
    $bk = Join-Path $TargetDir ("ConnectionOutHistory.ini.bak." + (Get-Date).ToString("yyyyMMdd-HHmmss"))
    Copy-Item $TargetFile $bk -Force
#    Write-Host "Đã tạo backup: $bk"
}

# 9) Thay thế file mới
Copy-Item $TempFile $TargetFile -Force
#Write-Host "Đã thay thế: $TargetFile"

# 10) (Tuỳ chọn) Ép UTF-8 nếu muốn chuẩn Unicode (bật nếu thấy bể font)
# $content = Get-Content $TargetFile
# Set-Content -Path $TargetFile -Value $content -Encoding UTF8

Write-Host "Done. Ultraview => File => Restart"
