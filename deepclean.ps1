# ================================
# DEEP CLEAN + OFFICE FULL CLEAN + INSTALL 7ZIP
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

$services = @("ClickToRunSvc","OfficeSvc")
foreach ($svc in $services) {
    Get-Service $svc -ErrorAction SilentlyContinue | Stop-Service -Force
}

# Kill Office process
$procs = @("winword","excel","powerpnt","outlook","onenote","msaccess","lync")
foreach ($p in $procs) {
    Get-Process $p -ErrorAction SilentlyContinue | Stop-Process -Force
}

# -------------------------------
# 3. ✅ FULL CLEAN OFFICE (AUTO DETECT)
# -------------------------------
Write-Host "`n[3] FULL CLEAN Microsoft Office..." -ForegroundColor Yellow

$officeFound = $false

# --- MSI uninstall ---
$paths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($path in $paths) {
    Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object {
        $_.DisplayName -match "Microsoft Office"
    } | ForEach-Object {

        $name = $_.DisplayName
        $guid = $_.PSChildName

        Write-Host "➡ Found: $name"

        if ($guid -match "^{.*}$") {
            Start-Process "msiexec.exe" -ArgumentList "/x $guid /quiet /norestart" -Wait
            Write-Host "✅ Removed MSI: $name"
            $officeFound = $true
        } elseif ($_.UninstallString) {
            Start-Process "cmd.exe" -ArgumentList "/c $($_.UninstallString)" -Wait
            Write-Host "✅ Removed CMD: $name"
            $officeFound = $true
        }
    }
}

# --- Click-to-Run ---
$ctr = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
if (Test-Path $ctr) {
    Write-Host "➡ Removing Click-to-Run Office..."
    Start-Process $ctr -ArgumentList "scenario=install scenariosubtype=uninstall" -Wait
    Write-Host "✅ Removed CTR Office"
    $officeFound = $true
}

if (-not $officeFound) {
    Write-Host "⚠ Not found Office"
}

# -------------------------------
# 4. ✅ UNINSTALL WINRAR
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

            Write-Host "➡ UnInstall: $($_.DisplayName)"

            if ($_.UninstallString) {
                Start-Process cmd.exe -ArgumentList "/c $($_.UninstallString) /quiet" -Wait
            }
        }
    }
}

Write-Host "`n[4] UnInstall WinRAR..."
Uninstall-App "WinRAR"

# -------------------------------
# ✅ MỞ PROGRAMS AND FEATURES
# -------------------------------
Write-Host "`n[5] Open App check..." -ForegroundColor Cyan
Start-Process "appwiz.cpl"

# -------------------------------
# 6. DELETE FOLDERS
# -------------------------------
Write-Host "`n[6] Clear folder..." -ForegroundColor Yellow

$folders = @(
    "C:\Program Files\Microsoft Office",
    "C:\Program Files (x86)\Microsoft Office",
    "C:\Program Files\Common Files\Microsoft Shared\ClickToRun",
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
# 7. REMOVE REGISTRY + MSI CACHE
# -------------------------------
Write-Host "`n[7] Clean registry..." -ForegroundColor Yellow

$regs = @(
    "HKCU:\Software\Microsoft\Office",
    "HKLM:\Software\Microsoft\Office",
    "HKLM:\Software\WOW6432Node\Microsoft\Office"
)

foreach ($r in $regs) {
    if (Test-Path $r) {
        Remove-Item $r -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# MSI cache
Get-ChildItem "HKLM:\Software\Classes\Installer\Products" | ForEach-Object {
    $p = $_.PSPath
    $name = (Get-ItemProperty $p -ErrorAction SilentlyContinue).ProductName

    if ($name -like "*Office*") {
        Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "🧨 Removed MSI cache: $name"
    }
}

# -------------------------------
# 8. INSTALL 7-ZIP
# -------------------------------
Write-Host "`n[8] SetUp 7-Zip..." -ForegroundColor Yellow

$toolPath = "D:\Tool"
$installer = "$toolPath\7zip.exe"

if (!(Test-Path $toolPath)) {
    New-Item -ItemType Directory -Path $toolPath | Out-Null
}

$url = "https://www.7-zip.org/a/7z2301-x64.exe"

try {
    Invoke-WebRequest -Uri $url -OutFile $installer
    Start-Process -FilePath $installer -ArgumentList "/S" -Wait
    Write-Host "✅ Done 7-Zip"
} catch {
    Write-Host "❌ Error Download/Setup: $_"
}

# -------------------------------
# 9. CLEAN TEMP
# -------------------------------
Write-Host "`n[9] Cleanup temp..." -ForegroundColor Yellow
Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue

# -------------------------------
# DONE
# -------------------------------
Write-Host "`n🔥 HOAN TAT!" -ForegroundColor Green
Write-Host "👉 Check lai trong Programs and Features" -ForegroundColor Cyan

Pause
