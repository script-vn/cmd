$sidList = @(wmic useraccount get name,sid | Select-String "S-1")
Write-Output "Chon USER cap quyen:`n"

foreach ($entry in $sidList) {
    $line = $entry.Line
    $parts = $line -split '\s+'
    if ($parts.Count -ge 2) {
        $name = $parts[0]
        $sid = $parts[1]
        Write-Output "USER: $name`nSID: $sid"
    }
}


$sidInput = Read-Host "`nInput SID of USER: "
$userFound = $false
$nameFound = ""

foreach ($entry in $sidList) {
    $line = $entry.Line
    $parts = $line -split '\s+'
    $name = $parts[0]
    $sid = $parts[1]

    if ($sid -eq $sidInput) {
        Write-Output "Cap Quyen Cho: $name"
        $userFound = $true
	$nameFound = $name
        break
    }
}

if (-not $userFound) {
    Write-Output "`nSID not found!!`n"
	return
}

cd C:\Users\Admin
do {
    $servicename = Read-Host "Nhap Services-Name: "
    if (![string]::IsNullOrWhiteSpace($servicename)) {
        try {
            $service = Get-Service -Name $servicename -ErrorAction Stop
        } catch {
            Write-Output "`nService '$servicename' not found, Again:`n"
            continue  # again
        }
        $sdString = "D:(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;RPWPCR;;;{0})" -f $sid
        try {
            $result = sc.exe sdset $servicename $sdString 2>&1
            if ($result -match "Access is denied" -or $result -match "Khong Du Quyen") {
                Write-Output "`nCan chay voi **Administrator**."
                return
            } else {
                Write-Output "`nAccepted service '$servicename'."
            }
        } catch {
            Write-Output "`nError sdset: $_"
            return
        }
    }
} while (![string]::IsNullOrWhiteSpace($servicename))

Write-Host "`nBye!."
