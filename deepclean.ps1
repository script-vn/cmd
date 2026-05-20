# ================================
# DEEP CLEAN + INSTALL 7ZIP + CHECK APP
# ================================

# ✅ Check Admin
if (-not ([Security.Principal.WindowsPrincipal] `
[Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole]::Administrator)) {

    Write-Host "❌ Vui lòng chạy bằng Administrator!" -ForegroundColor Red
    exit
}

Write-Host "=== 🔥 BAT DAU DEEP CLEAN 🔥 ===" -ForegroundColor Cyan

# -------------------------------
# 1. REMOVE WINDOWS KEY
# -------------------------------
Write-Host "`n[1] Gỡ key Windows..." -ForegroundColor Yellow

try {
    cscript //nologo $env:SystemRoot\System32\slmgr.vbs /upk
    cscript //nologo $env:SystemRoot\System32\slmgr.vbs /cpky
    Write-Host "✅ Đã gỡ key Windows"
} catch {
    Write-Host "❌ Lỗi: $_"
}

# -------------------------------
# 2. STOP SERVICES
# -------------------------------
Write-Host "`n[2] Stop services..." -ForegroundColor Yellow

Get-Service ClickToRunSvc -ErrorAction SilentlyContinue | Stop-Service -Force

# -------------------------------
# 3. UNINSTALL APP
# -------------------------------
function Uninstall-App($keyword) {

    $paths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $paths) {
        Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object {
            $_.DisplayName -like "*$keyword*"
        } | ForEach-Object {

            $name = $_.DisplayName
            $uninstall = $_.UninstallString

            Write-Host "➡ Gỡ: $name"

            if ($uninstall) {
                if ($uninstall -match "msiexec") {
                    $cmd = $uninstall -replace "/I", "/X"
                    Start-Process cmd.exe -ArgumentList "/c $cmd /quiet /norestart" -Wait
                } else {
                    Start-Process cmd.exe -ArgumentList "/c `"$uninstall`" /quiet /norestart" -Wait
                }
            }
        }
    }
}

Write-Host "`n[3] Gỡ Office + WinRAR..."
Uninstall-App "Microsoft Office"
Uninstall-App "WinRAR"

# -------------------------------
# ✅ MỞ PROGRAMS AND FEATURES
# -------------------------------
Write-Host "`n[4] Mở danh sách ứng dụng để kiểm tra..." -ForegroundColor Cyan
Start-Process "appwiz.cpl"

# -------------------------------
# 5. DELETE FOLDERS
# -------------------------------
Write-Host "`n[5] Xóa folder..." -ForegroundColor Yellow

$folders = @(
    "C:\Program Files\Microsoft Office",
    "C:\Program Files (x86)\Microsoft Office",
    "$env:ProgramData\Microsoft\Office",
    "$env:APPDATA\Microsoft\Office",
    "$env:LOCALAPPDATA\Microsoft\Office",
    "C:\Program Files\WinRAR",
    "$env:APPDATA\WinRAR"
)

foreach ($f in $folders) {
    if (Test-Path $f) {
        Remove-Item $f -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# -------------------------------
# 6. INSTALL 7-ZIP
# -------------------------------
Write-Host "`n[6] Cài 7-Zip..." -ForegroundColor Yellow

$toolPath = "D:\Tool"
$installer = "$toolPath\7zip.exe"

if (!(Test-Path $toolPath)) {
    New-Item -ItemType Directory -Path $toolPath | Out-Null
}

$url = "https://www.7-zip.org/a/7z2301-x64.exe"

try {
    Invoke-WebRequest -Uri $url -OutFile $installer
    Write-Host "✅ Tải xong 7-Zip"

    Start-Process -FilePath $installer -ArgumentList "/S" -Wait
    Write-Host "✅ Đã cài 7-Zip"

} catch {
    Write-Host "❌ Lỗi tải/cài: $_"
}

# -------------------------------
# 7. CLEAN TEMP
# -------------------------------
Write-Host "`n[7] Cleanup temp..." -ForegroundColor Yellow
Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue

# -------------------------------
# DONE
# -------------------------------
Write-Host "`n🔥 HOAN TAT!" -ForegroundColor Green
Write-Host "👉 Bạn có thể kiểm tra lại trong Programs and Features" -ForegroundColor Cyan

Pause
