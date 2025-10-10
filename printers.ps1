# Open-DevicesAndPrinters.ps1
# This script opens the "Devices and Printers" window on Windows 10.
# in CMD can run : control printers or : explorer shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}
Start-Process control.exe printers
Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
