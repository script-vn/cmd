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
        "https://chrome.google.com/webstore/detail/office-editing-for-docs/aohghmighlieiainnegkcijnfilokake",
        "https://chromewebstore.google.com/detail/google-scholar-pdf-reader/dahenjhkoodjbpjheillcadbppiidmhp"
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
Write-Host "`n[3] Set Chrome as default..." -ForegroundColor Yellow

Start-Process "ms-settings:defaultapps"

Write-Host "`n👉 Gợi ý cần làm (bắt buộc thủ công):"
Write-Host "   • Chọn Chrome làm Default Browser"
Write-Host "   • Set Chrome cho:"
Write-Host "       .pdf"
Write-Host "       .docx"
Write-Host "       .xlsx"
Write-Host "       .pptx"

# -------------------------------
# DONE
# -------------------------------
Write-Host "`n✅ DONE ✅" -ForegroundColor Green
Pause
