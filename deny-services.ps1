# $sid = Read-Host "input SID User:"
$sidList = @(wmic useraccount get name,sid | Select-String "S-1")
if ($sidList.Count -ge 5) {
    $line = $sidList[4].Line
    $sid = ($line -split '\s+')[1]
#    Write-Output "SID thứ 5 là: $sid"
} else {
#    Write-Output "Không có đủ 5 SID trong danh sách."
}

cd C:\Users\Admin
do {
    $servicename = Read-Host "Input Service Name or Press Enter to quick:"
    
    if (![string]::IsNullOrWhiteSpace($servicename)) {
        $sdString = "D:(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;RPWPCR;;;{0})" -f $sid
        $sdString = "D:(D;;RPWPCR;;;{0})(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)" -f $sid
        Write-Host "Deny Services : $servicename"
        sc.exe sdset $servicename $sdString
        Write-Host "Deny Start-Stop Services `"$servicename`".`n"
    }
} while (![string]::IsNullOrWhiteSpace($servicename))

Write-Host "Bye!."
