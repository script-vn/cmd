# ================================
# INSTALL 7-ZIP + OPEN PROGRAMS & FEATURES
# ================================

# ✅ Check Admin
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)

if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "❌ Run as Administrator!" -ForegroundColor Red
    exit
}

Write-Host "=== 📦 SETUP 7-ZIP ===" -ForegroundColor Cyan

# -------------------------------
# 1. Tạo thư mục Tool
# -------------------------------
$toolPath = "D:\Tool"
if (!(Test-Path $toolPath)) {
    New-Item -ItemType Directory -Path $toolPath | Out-Null
    Write-Host "✅ Created D:\Tool"
}

# -------------------------------
# 2. Download 7-Zip
# -------------------------------
$file = "$toolPath\7zip.exe"
$url = "https://www.7-zip.org/a/7z2301-x64.exe"

Write-Host "`n⬇ Downloading 7-Zip..."

try {
    Invoke-WebRequest -Uri $url -OutFile $file
    Write-Host "✅ Download done"
} catch {
    Write-Host "❌ Download eror: $_"
    exit
}

# -------------------------------
# 3. Install silent
# -------------------------------
Write-Host "`n📦 Installing 7-Zip..."

try {
    Start-Process -FilePath $file -ArgumentList "/S" -Wait
    Write-Host "✅ Installed 7-Zip"
} catch {
    Write-Host "❌ Install error: $_"
}

# -------------------------------
# 4. Mở Programs & Features
# -------------------------------
Write-Host "`n📂 Opening Programs & Features..." -ForegroundColor Yellow
Start-Process "appwiz.cpl"

Write-Host "`n👉 Uninstall now!" -ForegroundColor Green

Pause
