# Download Windows Exporter
$Installer = "$env:TEMP\windows_exporter.msi"

Write-Host "Downloading windows_exporter..." -ForegroundColor Yellow

Invoke-WebRequest `
-Uri "https://github.com/prometheus-community/windows_exporter/releases/download/v0.31.8/windows_exporter-0.31.8-amd64.msi" `
-OutFile $Installer

# Install silently
Write-Host "Installing windows_exporter..." -ForegroundColor Yellow

Start-Process msiexec.exe -ArgumentList "/i `"$Installer`" /quiet" -Wait

# Create Firewall Rule
if (-not (Get-NetFirewallRule -DisplayName "windows_exporter" -ErrorAction SilentlyContinue)) {

    New-NetFirewallRule `
    -DisplayName "windows_exporter" `
    -Direction Inbound `
    -Protocol TCP `
    -LocalPort 9182 `
    -Action Allow

    Write-Host "Firewall rule created." -ForegroundColor Green
}
else {
    Write-Host "Firewall rule already exists." -ForegroundColor Green
}

# Wait for service
Start-Sleep -Seconds 5

# Check service
Write-Host ""
Write-Host "=== Service Status ===" -ForegroundColor Cyan

Get-Service windows_exporter

# Check port
Write-Host ""
Write-Host "=== Port Check ===" -ForegroundColor Cyan

Test-NetConnection localhost -Port 9182

# Open Metrics page
Write-Host ""
Write-Host "Opening metrics page..." -ForegroundColor Green

Start-Process "http://localhost:9182/metrics"

Write-Host ""
Write-Host "Installation completed." -ForegroundColor Green
