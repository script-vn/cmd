# Lấy danh sách SID từ useraccount
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

$sidInput = Read-Host "`nInput SID of USER "
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

# Tìm các service có tên khớp
$services = Get-Service | Where-Object { $_.Name -match "rAgent" -or $_.Name -match "FARCARDS" }

if ($services.Count -gt 0) {
    Write-Output "`nList:`n"
    foreach ($svc in $services) {
        Write-Output "Name: $($svc.Name) | Status: $($svc.Status)"

        $sdString = "D:(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;RPWPCR;;;{0})" -f $sid
        try {
            $result = sc.exe sdset $svc.Name $sdString 2>&1
            if ($result -match "Access is denied" -or $result -match "Khong Du Quyen") {
                Write-Output "Run **Administrator**."
                return
            } else {
                Write-Output "Accepted service '$($svc.DisplayName)'."
            }
        } catch {
            Write-Output "`nError sdset: $_"
            return
        }

        try {
            $svc.Refresh()
            $canStart = $svc.Status -eq 'Stopped'
            $canStop = $svc.CanStop

            Write-Output "'$($svc.DisplayName)' trạng thái: $($svc.Status)"

            if ($canStart) {
                Write-Output "Ready Start."
            } elseif ($canStop) {
                Write-Output "Ready Stop/Restart."
            } else {
                Write-Output "Ready Start."
            }
        } catch {
            Write-Output "Can't Check Services. $_"
        }
    }
} else {
    Write-Output "`nnot found services. Input: .`n"
    do {
        $servicename = Read-Host "Nhap Services-Name "
        if (![string]::IsNullOrWhiteSpace($servicename)) {
            try {
                $service = Get-Service -Name $servicename -ErrorAction Stop
                $displayName = $service.DisplayName
                $sdString = "D:(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;RPWPCR;;;{0})" -f $sid
                $result = sc.exe sdset $servicename $sdString 2>&1
                if ($result -match "Access is denied" -or $result -match "Khong Du Quyen") {
                    Write-Output "Run **Administrator**."
                    return
                } else {
                    Write-Output "Accepted service '$displayName'."
                }

                $service.Refresh()
                $canStart = $service.Status -eq 'Stopped'
                $canStop = $service.CanStop

                Write-Output "'$displayName' checking is: $($service.Status)"

                if ($canStart) {
                    Write-Output "Ready Start."
                } elseif ($canStop) {
                    Write-Output "Ready Stop/Restart."
                } else {
                    Write-Output "Ready Start."
                }
            } catch {
                Write-Output "Service '$servicename' Can't Check Services: $_"
            }
        }
    } while (![string]::IsNullOrWhiteSpace($servicename))
}

Write-Host "`nEnd script."
