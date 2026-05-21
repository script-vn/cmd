# =============================================
# SETUP TOOL + CHROME + EXTENSION
# =============================================

# ✅ Check Admin
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)

if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "❌ Run as Administrator!" -ForegroundColor Red
    Pause
    exit
}

Write-Host "=== 🚀 SETUP TOOL & CHROME ===" -ForegroundColor Cyan

# -------------------------------
# 1. Install 7-Zip
# -------------------------------
Write-Host "`n[1] Installing 7-Zip..."

$toolPath = "D:\Tool"
$file = Join-Path $toolPath "7zip.exe"
$url = "https://www.7-zip.org/a/7z2301-x64.exe"

# Tạo folder nếu chưa có
if (!(Test-Path $toolPath)) {
    New-Item -ItemType Directory -Path $toolPath | Out-Null
}

try {
    Invoke-WebRequest -Uri $url -OutFile $file -UseBasicParsing
    Start-Process -FilePath $file -ArgumentList "/S" -Wait
    Write-Host "✅ 7-Zip Installed"
} catch {
    Write-Host "❌ 7-Zip error: $_" -ForegroundColor Red
}

# -------------------------------
# 2. Open Chrome with Extension Tabs
# -------------------------------
Write-Host "`n[2] Opening Chrome extensions..."

# Check cả 2 đường dẫn Chrome
$chromePath1 = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$chromePath2 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

$chromePath = $null

if (Test-Path $chromePath1) {
    $chromePath = $chromePath1
} elseif (Test-Path $chromePath2) {
    $chromePath = $chromePath2
}

if ($chromePath) {

    $urls = @(
        "https://chromewebstore.google.com/detail/office-editing-for-docs-s/gbkeegbaiigmenfmjfclcdgdpimamgkj",
        "https://chromewebstore.google.com/detail/google-scholar-pdf-reader/dahenjhkoodjbpjheillcadbppiidmhp",
        "https://chromewebstore.google.com/detail/google-docs-offline/ghbmnnjooekpmoecnnnilnnbdlolhkhi"
    )

    # ✅ mở 1 Chrome nhiều tab
    Start-Process $chromePath ($urls -join " ")

    Write-Host "✅ Chrome opened → Click 'Add to Chrome'" -ForegroundColor Green
} else {
    Write-Host "⚠ Chrome not found!" -ForegroundColor Yellow
}

# -------------------------------
# 3. Open Default Apps Settings
# -------------------------------

Write-Host "=== 🌐 AUTO SET DEFAULT APP (CHROME) ===" -ForegroundColor Cyan

# -------------------------------
# 1. Tạo file mẫu
# -------------------------------
$tempPath = "$env:TEMP\DefaultTest"
if (!(Test-Path $tempPath)) {
    New-Item -ItemType Directory -Path $tempPath | Out-Null
}

# tạo file mẫu
Set-Content "$tempPath\test.pdf" "test"
Set-Content "$tempPath\test.docx" "test"
Set-Content "$tempPath\test.xlsx" "test"
Set-Content "$tempPath\test.pptx" "test"

# -------------------------------
# 2. Function mở file + yêu cầu chọn Chrome
# -------------------------------
function Set-Default($file, $ext) {

    Write-Host "`n👉 Đang set mặc định cho .$ext" -ForegroundColor Yellow
    Write-Host "➡ Chọn: Google Chrome + tick 'Always use this app'" -ForegroundColor Green

    Start-Process $file

    Pause
}

# -------------------------------
# 3. Lần lượt từng loại file
# -------------------------------
Set-Default "$tempPath\test.pdf" "pdf"
Set-Default "$tempPath\test.docx" "docx"
Set-Default "$tempPath\test.xlsx" "xlsx"
Set-Default "$tempPath\test.pptx" "pptx"

# -------------------------------
# DONE
# -------------------------------
Write-Host "`n✅ DONE - Chrome đã được gán mặc định!" -ForegroundColor Cyan

Pause

