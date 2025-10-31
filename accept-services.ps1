# $sid = Read-Host "input SID User:"

# Da chinh sua


$sidList = @(wmic useraccount get name,sid | Select-String "S-1")
Write-Output "Chon USER can cap quyen:`n"

foreach ($entry in $sidList) {
    $line = $entry.Line
    $parts = $line -split '\s+'
    $name = $parts[0]
    $sid = $parts[1]
    Write-Output "USER: $name : SID: $sid"
}

$sidInput = Read-Host "`nSID input:`n"
$userFound = $false

foreach ($entry in $sidList) {
    $line = $entry.Line
    $parts = $line -split '\s+'
    $name = $parts[0]
    $sid = $parts[1]

    if ($sid -eq $sidInput) {
        Write-Output "`nUser : $name"
        $userFound = $true
        break
    }
}

if (-not $userFound) {
    Write-Output "`nSID not found!!`n"
}

cd C:\Users\Admin
do {
    $servicename = Read-Host "`nNhap Services-Name: "
    
    if (![string]::IsNullOrWhiteSpace($servicename)) {
      #  $sid = (whoami /user | Select-String -Pattern "S-1.*").Matches.Value
        $sdString = "D:(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;RPWPCR;;;{0})" -f $sid
        
        Write-Host "`nCap Quyen Services : $servicename"
        sc.exe sdset $servicename $sdString
    }
} while (![string]::IsNullOrWhiteSpace($servicename))

Write-Host "`nBye!."
