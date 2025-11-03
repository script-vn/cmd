
$loggedInUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
$username = $loggedInUser.Split('\')[-1]
$sid = (Get-WmiObject Win32_UserAccount | Where-Object { $_.Name -eq $username }).SID
Write-Output "User: $username"
Write-Output "SID: $sid"
$services = Get-Service | Where-Object { $_.Name -match "rAgent" -or $_.Name -match "FARCARDS" }
if ($services.Count -gt 0) {
    foreach ($svc in $services) {

         $sdString = "D:(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;RPWPCR;;;{0})" -f $sid
	Write-Output "`nName: $($svc.Name)"
        sc.exe sdset $svc.Name $sdString	
    }
}
