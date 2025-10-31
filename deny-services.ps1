# $sid = Read-Host "input SID User:"

$sidList = @(wmic useraccount get name,sid | Select-String "S-1")
$sidListFiltered = $sidList | Where-Object { $_.Line -match "1002" }

if ($sidListFiltered.Count -gt 0) {
    foreach ($entry in $sidListFiltered) {
        $line = $entry.Line
        $parts = $line -split '\s+'
        $name = $parts[0]
        $sid = $parts[1]
        Write-Output "HUY Quyen USER: $name - co SID: $sid"
    }
} else {
    Write-Output "Chon USER Can HUY Quyen:"
    foreach ($entry in $sidList) {
        $line = $entry.Line
        $parts = $line -split '\s+'
        $name = $parts[0]
        $sid = $parts[1]
        Write-Output "USER: $name - SID: $sid"
    }

    $userSID = Read-Host "NHAP SID da COPY"
    Write-Output "SID can HUY quyen: $userSID"
}


cd C:\Users\Admin
do {
    $servicename = Read-Host "NHAP Service Name or NHAN Enter de THOAT:"
    
    if (![string]::IsNullOrWhiteSpace($servicename)) {
        $sdString = "D:(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;RPWPCR;;;{0})" -f $sid
        $sdString = "D:(D;;RPWPCR;;;{0})(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)" -f $sid
        Write-Host "HUY Quyen Services : $servicename"
        sc.exe sdset $servicename $sdString
        Write-Host "HUY Quyen Start-Stop Services `"$servicename`".`n"
    }
} while (![string]::IsNullOrWhiteSpace($servicename))

Write-Host "Bye!."
