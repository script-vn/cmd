# ================================
# DEEP CLEAN + INSTALL 7ZIP + CHECK APP
# ================================

# ✅ Check Admin
if (-not ([Security.Principal.WindowsPrincipal] `
[Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole]::Administrator)) {

    Write-Host "❌ Run with Administrator!" -ForegroundColor Red
    exit
}

Write-Host "=== 🔥 BAT DAU DEEP CLEAN 🔥 ===" -ForegroundColor Cyan

# -------------------------------
# 1. REMOVE WINDOWS KEY
# -------------------------------
Write-Host "`n[1] UnInstall key Windows..." -ForegroundColor Yellow

try {
    cscript //nologo $env:SystemRoot\System32\slmgr.vbs /upk
    cscript //nologo $env:SystemRoot\System32\slmgr.vbs /cpky
    Write-Host "✅ Clear key Windows"
} catch {
    Write-Host "❌ Error: $_"
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

Write-Host "`n[3] UnInstall Office + WinRAR..."
Uninstall-App "Microsoft Office"
Uninstall-App "WinRAR"

# -------------------------------
# ✅ MỞ PROGRAMS AND FEATURES
# -------------------------------
Write-Host "`n[4] Open App check..." -ForegroundColor Cyan
Start-Process "appwiz.cpl"

# -------------------------------
# 5. DELETE FOLDERS
# -------------------------------
Write-Host "`n[5] Clear folder..." -ForegroundColor Yellow

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
Write-Host "`n[6] SetUp 7-Zip..." -ForegroundColor Yellow

$toolPath = "D:\Tool"
$installer = "$toolPath\7zip.exe"

if (!(Test-Path $toolPath)) {
    New-Item -ItemType Directory -Path $toolPath | Out-Null
}

$url = "https://www.7-zip.org/a/7z2301-x64.exe"

try {
    Invoke-WebRequest -Uri $url -OutFile $installer
    Write-Host "✅ Download xong 7-Zip"

    Start-Process -FilePath $installer -ArgumentList "/S" -Wait
    Write-Host "✅ Done 7-Zip"

} catch {
    Write-Host "❌ Error Download/Setup: $_"
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
Write-Host "👉 |check on Programs and Features" -ForegroundColor Cyan

Pause
