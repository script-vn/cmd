
do {
    $servicename = Read-Host "Nhập tên dịch vụ cần chỉnh sửa (hoặc nhấn Enter để thoát)"
    
    if (![string]::IsNullOrWhiteSpace($servicename)) {
        $sid = (whoami /user | Select-String -Pattern "S-1.*").Matches.Value
        $sdString = "D:(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;RPWPCR;;;{0})" -f $sid
        
        Write-Host "Đang áp dụng quyền cho dịch vụ: $servicename"
        sc.exe sdset $servicename $sdString
        Write-Host "Hoàn tất chỉnh sửa quyền cho dịch vụ `"$servicename`".`n"
    }
} while (![string]::IsNullOrWhiteSpace($servicename))

Write-Host "Đã thoát script."
