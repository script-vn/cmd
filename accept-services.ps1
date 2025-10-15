$sid = Read-Host "input SID User:"
cd C:\Users\Admin
do {
    $servicename = Read-Host "Input Service Name or Press Enter to quick:"
    
    if (![string]::IsNullOrWhiteSpace($servicename)) {
      #  $sid = (whoami /user | Select-String -Pattern "S-1.*").Matches.Value
        $sdString = "D:(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;RPWPCR;;;{0})" -f $sid
        
        Write-Host "Accepting Services : $servicename"
        sc.exe sdset $servicename $sdString
        Write-Host "Acceptep Start-Stop Services `"$servicename`".`n"
    }
} while (![string]::IsNullOrWhiteSpace($servicename))

Write-Host "Bye!."
