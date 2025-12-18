# ================== WinForms GUI:==================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Bắt buộc: bật Visual Styles và đảm bảo dùng Application.Run
[void][System.Windows.Forms.Application]::EnableVisualStyles()

#=== Kiểm tra quyền Admin (dùng cho các thao tác đặc quyền) ===
function Test-IsAdmin {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

#=== Khởi tạo Form & Controls ===
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tool Support POS Order SASIN - Dev by NamIT"
$form.Size = New-Object System.Drawing.Size(520,400)
$form.StartPosition = "CenterScreen"

# Label tạo user
$labelNewUser = New-Object System.Windows.Forms.Label
$labelNewUser.Text = "Add new User:"
$labelNewUser.Location = New-Object System.Drawing.Point(20,10)
$labelNewUser.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($labelNewUser)

# Label User name
$labelUserName = New-Object System.Windows.Forms.Label
$labelUserName.Text = "User name:"
$labelUserName.Location = New-Object System.Drawing.Point(120,10)
$labelUserName.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($labelUserName)

# TextBox username
$textUserName = New-Object System.Windows.Forms.TextBox
$textUserName.Location = New-Object System.Drawing.Point(120,30)
$textUserName.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($textUserName)

# Label Full name
$labelFullName = New-Object System.Windows.Forms.Label
$labelFullName.Text = "Full name:"
$labelFullName.Location = New-Object System.Drawing.Point(230,10)
$labelFullName.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($labelFullName)

# TextBox full name
$textFullName = New-Object System.Windows.Forms.TextBox
$textFullName.Location = New-Object System.Drawing.Point(230,30)
$textFullName.Size = New-Object System.Drawing.Size(150,20)
$form.Controls.Add($textFullName)

# Nút tạo user
$buttonCreateUser = New-Object System.Windows.Forms.Button
$buttonCreateUser.Text = "Create user"
$buttonCreateUser.Location = New-Object System.Drawing.Point(390,29)
$buttonCreateUser.Size = New-Object System.Drawing.Size(80,23)
$form.Controls.Add($buttonCreateUser)

# Label chọn user
$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Text = "Select user:"
$labelUser.Location = New-Object System.Drawing.Point(20,70)
$labelUser.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($labelUser)

# ComboBox user
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(120,70)
$comboBox.Size = New-Object System.Drawing.Size(200,20)
$comboBox.DropDownStyle = "DropDownList"
$form.Controls.Add($comboBox)

# Nút chặn Zalo
$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Text = "Run Deny Zalo"
$buttonExecute.Location = New-Object System.Drawing.Point(320,69)
$buttonExecute.Size = New-Object System.Drawing.Size(150,23)
$form.Controls.Add($buttonExecute)

# TextBox ghi log
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(15,110)
$logBox.Size = New-Object System.Drawing.Size(300,160)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$form.Controls.Add($logBox)
$logBox.Text = "đã mở log... fdsfsdfs. fdsfdsf .sdfdsfsdf .sdfsdfdf fdsfdsf fdsfdsf"

$noteBox = New-Object System.Windows.Forms.TextBox
$noteBox.Location = New-Object System.Drawing.Point(320,110)
$noteBox.Size = New-Object System.Drawing.Size(170,160)
$noteBox.Multiline = $true
$noteBox.ReadOnly = $true
$form.Controls.Add($noteBox)
$noteBox.Text = "====== SETUP-POS-NSO =====`r`n***Admin:`r`n+ Pass Admin + User(ZaloPC)`r`n+ Driver Printer - Update windows`r`n+ Change User account Admin`r`n*** USER`r`n+ Never Sleep`r`n+ Set IP LAN`r`n+ Active Services-Dcorp`r`n+ Set Audio note run system, ultraview(Pass + Auto run)"

# Nút Resources
$buttonResources = New-Object System.Windows.Forms.Button
$buttonResources.Text = "Resources"
$buttonResources.Location = New-Object System.Drawing.Point(20, 280)
$buttonResources.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($buttonResources)

# Nút Set IP LAN và DNS
$buttonSetIP = New-Object System.Windows.Forms.Button
$buttonSetIP.Text = "Set IP LAN & DNS"
$buttonSetIP.Location = New-Object System.Drawing.Point(150,280)
$buttonSetIP.Size = New-Object System.Drawing.Size(180,30)
$form.Controls.Add($buttonSetIP)

# Nút Driver Printer
$buttonPrinterDriver = New-Object System.Windows.Forms.Button
$buttonPrinterDriver.Text = "Driver Printer"
$buttonPrinterDriver.Location = New-Object System.Drawing.Point(350,280)
$buttonPrinterDriver.Size = New-Object System.Drawing.Size(140,30)
$form.Controls.Add($buttonPrinterDriver)


# Nút Active Services Dcorp
$buttonSetPerm = New-Object System.Windows.Forms.Button
$buttonSetPerm.Text = "Active Services Dcorp"
$buttonSetPerm.Location = New-Object System.Drawing.Point(20, 320)   # bạn có thể chỉnh vị trí tùy ý
$buttonSetPerm.Size = New-Object System.Drawing.Size(180, 30)
$form.Controls.Add($buttonSetPerm)


# Nút Open Printers
$buttonOpenPrinters = New-Object System.Windows.Forms.Button
$buttonOpenPrinters.Text = "Open Printers"
$buttonOpenPrinters.Location = New-Object System.Drawing.Point(220, 320)
$buttonOpenPrinters.Size = New-Object System.Drawing.Size(140, 30)
$form.Controls.Add($buttonOpenPrinters)


#=== Logging ===
function Write-Log($message) {
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $logBox.AppendText("[$timestamp] $message`r`n")
}

#=== Tiền điều kiện hệ thống (khuyến nghị): ProfSvc & Default Profile ===
function Ensure-ProfSvcAndDefaultProfileOK {
    try {
        $svc = Get-Service -Name "ProfSvc" -ErrorAction Stop
        if ($svc.StartType -ne "Automatic") { Set-Service -Name "ProfSvc" -StartupType Automatic }
        if ($svc.Status -ne "Running") { Start-Service -Name "ProfSvc" }
        Write-Log "ProfSvc OK (Automatic & Running)."
    } catch {
        Write-Log "ProfSvc khong kha dung: $($_.Exception.Message)"
    }

    $def = "C:\Users\Default"
    if (-not (Test-Path $def)) { Write-Log "Thieu C:\Users\Default" }
    else {
        $nt = Join-Path $def "NTUSER.DAT"
        if (-not (Test-Path $nt)) { Write-Log "Thieu NTUSER.DAT trong Default Profile." }
    }
}

#=== Helpers: Lấy profile path từ Registry & Tạo profile an toàn ===
function Get-ProfilePath {
    param([string]$UserName)
    try {
        $sid = (Get-LocalUser -Name $UserName).Sid.Value
        $regBase = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
        return (Get-ItemProperty -Path $regBase -ErrorAction SilentlyContinue).ProfileImagePath
    } catch { return $null }
}

function Ensure-UserProfileCreated {
    param([string]$UserName)

    # Tạm cấp quyền admin tránh deny logon locally (nếu có)
    try { Add-LocalGroupMember -Group "Administrators" -Member $UserName -ErrorAction SilentlyContinue } catch {}

    # Kích hoạt tạo profile và ĐỢI nó thoát tự nhiên
    $cred = Get-Credential -UserName $UserName -Message "Nhap mat khau de tao profile cho user $UserName"
    $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c whoami" -Credential $cred -PassThru
    try { Wait-Process -Id $p.Id -Timeout 30 } catch {}

    # Lấy path từ Registry
    $profilePath = Get-ProfilePath -UserName $UserName
    if (-not $profilePath) { throw "Khong tim duoc ProfileImagePath trong Registry cho user $UserName." }

    $ntUserDatPath = Join-Path $profilePath "NTUSER.DAT"
    if (-not (Test-Path $ntUserDatPath)) { throw "NTUSER.DAT chua ton tai sau khi tao profile." }

    # Thu hồi quyền admin tạm thời
    try { Remove-LocalGroupMember -Group "Administrators" -Member $UserName -ErrorAction SilentlyContinue } catch {}

    return $ntUserDatPath
}

function Set-DenyZalo {
    param([string]$UserName)

    $ntUserDatPath = Ensure-UserProfileCreated -UserName $UserName
    $hiveName = "HKU\TempHive"

    try {
        Write-Log "Dang load hive NTUSER.DAT..."
        & reg.exe load $hiveName $ntUserDatPath | Out-Null

        & reg.exe add "$hiveName\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v DisallowRun /t REG_DWORD /d 1 /f | Out-Null
        & reg.exe add "$hiveName\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" /v 1 /t REG_SZ /d Zalo.exe /f | Out-Null
        & reg.exe add "$hiveName\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" /v 2 /t REG_SZ /d ZaloUpdate.exe /f | Out-Null

        Write-Log "Da ghi registry chan Zalo cho user '$UserName'."
    }
    finally {
        Write-Log "Dang unload hive..."
        & reg.exe unload $hiveName | Out-Null
    }
}

#=== Load danh sách user ===
function LoadUsers {
    param([string]$selectUser = $null)
    $comboBox.Items.Clear()
    $users = Get-LocalUser | Where-Object { $_.Enabled -eq $true -and $_.Name -ne "Admin" } | Select-Object -ExpandProperty Name
    if ($users) { $comboBox.Items.AddRange($users) }
    Write-Log "Da tai danh sach user."
    if ($selectUser -and ($users -contains $selectUser)) {
        $comboBox.SelectedItem = $selectUser
        Write-Log "Da chon user mac dinh: $selectUser"
    }
}

#=== Sự kiện nút Create user (giữ nguyên hành vi cũ, nhưng an toàn hơn) ===
$buttonCreateUser.Add_Click({
    $newUser = $textUserName.Text.Trim()
    $fullName = $textFullName.Text.Trim()
    if (-not $newUser) {
        [System.Windows.Forms.MessageBox]::Show("Vui long nhap User name.","Thong bao",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if (-not (Test-IsAdmin)) {
        [System.Windows.Forms.MessageBox]::Show("Can Admin rights de tao user.","Thong bao",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    try {
        Write-Log "Dang tao user '$newUser'..."
        $password = ConvertTo-SecureString "123" -AsPlainText -Force
        New-LocalUser -Name $newUser -FullName $fullName -Password $password -ErrorAction Stop
        Set-LocalUser -Name $newUser -PasswordNeverExpires $true
        & net.exe user $newUser /passwordchg:no | Out-Null
        Add-LocalGroupMember -Group "Users" -Member $newUser -ErrorAction SilentlyContinue
        LoadUsers $newUser
        $textUserName.Text = ""
        $textFullName.Text = ""
        Write-Log "Da tao user '$newUser' thanh cong."
    } catch {
        Write-Log "Loi khi tao user: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Loi khi tao user: $($_.Exception.Message)","Loi",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

#=== Sự kiện nút Run Deny Zalo ===
$buttonExecute.Add_Click({
    $userName = $comboBox.SelectedItem
    if (-not $userName) {
        [System.Windows.Forms.MessageBox]::Show("Vui long chon user.","Thong bao",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if (-not (Test-IsAdmin)) {
        [System.Windows.Forms.MessageBox]::Show("Can Admin rights de thuc thi.","Thong bao",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    try {
        Write-Log "Dang thuc thi Deny Zalo cho '$userName'..."
        Set-DenyZalo -UserName $userName
        [System.Windows.Forms.MessageBox]::Show("Da chan Zalo cho user '$userName' thanh cong.","Thanh cong",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    } catch {
        Write-Log "Loi: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Loi: $($_.Exception.Message)","Loi",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})


# Xử lý khi nhấn nút Resources
$buttonResources.Add_Click({
    try {
        Write-Log "Dang mo trang Resources..."
        $url = "https://sasinvn-my.sharepoint.com/:f:/g/personal/nam_tran_sasin_vn/EgtgabtBUqtGoh-XXw4Xdj0BR8hhr6sfyR3Uok5T6LaFPw"

        # Thử mở bằng Chrome nếu có, nếu không thì dùng browser mặc định
        $chromePaths = @(
            "$Env:ProgramFiles\Google\Chrome\Application\chrome.exe",
            "$Env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe"
        )

        $chromeExe = $null
        foreach ($p in $chromePaths) {
            if (Test-Path $p) { $chromeExe = $p; break }
        }

        if (-not $chromeExe) {
            # Nếu chrome không nằm ở vị trí chuẩn, thử gọi theo tên (PATH)
            $chromeExe = "chrome.exe"
        }

        try {
            Start-Process $chromeExe $url -ErrorAction Stop
            Write-Log "Da mo Chrome toi Resources."
        } catch {
            Write-Log "Khong mo duoc bang Chrome, dang mo bang trinh duyet mac dinh..."
            Start-Process $url
            Write-Log "Da mo trang Resources bang trinh duyet mac dinh."
        }
    } catch {
        Write-Log "Loi khi mo Resources: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Loi khi mo Resources: $($_.Exception.Message)", "Loi", "OK", "Error")
    }
})


#=== Sự kiện nút Set IP LAN & DNS ===
$buttonSetIP.Add_Click({
    if (-not (Test-IsAdmin)) {
        [System.Windows.Forms.MessageBox]::Show("Can Admin rights de dat IP/DNS.","Thong bao",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    try {
        Write-Log "Dang cau hinh IP LAN va DNS..."
        $interface = Get-NetAdapter | Where-Object { $_.Status -eq "Up" } | Select-Object -First 1
        if (-not $interface) { throw "Khong tim thay card mang dang hoat dong." }
        $interfaceName = $interface.Name
        Write-Log "Da chon card mang: $interfaceName"
        New-NetIPAddress -InterfaceAlias $interfaceName -IPAddress "192.168.1.11" -PrefixLength 24 -DefaultGateway "192.168.1.1" -ErrorAction Stop
        Write-Log "Da dat IP: 192.168.1.11, Gateway: 192.168.1.1"
        Set-DnsClientServerAddress -InterfaceAlias $interfaceName -ServerAddresses ("8.8.8.8","8.8.4.4") -ErrorAction Stop
        Write-Log "Da dat DNS: 8.8.8.8, 8.8.4.4"
        [System.Windows.Forms.MessageBox]::Show("Da cau hinh IP va DNS thanh cong.","Thanh cong",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    } catch {
        Write-Log "Loi khi cau hinh IP/DNS: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Loi: $($_.Exception.Message)","Loi",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

#=== Sự kiện nút Driver Printer ===
$buttonPrinterDriver.Add_Click({
    try {
        Write-Log "Dang tai xuong file driver may in..."
        $url = "https://sasinvn-my.sharepoint.com/:u:/g/personal/nam_tran_sasin_vn/EaSWlvxwl7RIljE-aEhCU3ABAKTlJG2jTIFv6hJxR9-5xA?download=1"
        $destination = "D:\DriverPrinter.printerExport"
        Invoke-WebRequest -Uri $url -OutFile $destination
        Write-Log "Da tai xuong file driver tai $destination"
        $printBrmPath = "C:\Windows\System32\PrintBrmUi.exe"
        if (Test-Path $printBrmPath) {
            Start-Process $printBrmPath
            Write-Log "Da mo PrintBrmUi.exe de import driver."
        } else {
            Write-Log "Khong tim thay PrintBrmUi.exe tai $printBrmPath"
            [System.Windows.Forms.MessageBox]::Show("Khong tim thay PrintBrmUi.exe","Loi",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    } catch {
        Write-Log "Loi khi tai hoac mo driver: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Loi: $($_.Exception.Message)","Loi",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})


# Handler nút Set Permissions Dcorp
$buttonSetPerm.Add_Click({
    try {
        Write-Log "Dang thiet lap quyen Dcorp cho cac dich vu (rAgent, FARCARDS)..."

        # Lấy user hiện đăng nhập và SID (ưu tiên .NET, fallback WMI nếu cần)
        try {
            $username = $env:USERNAME
            $sid = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
            Write-Log "User dang login: $username"
            Write-Log "SID: $sid"
        } catch {
            Write-Log "Khong lay duoc SID bang .NET, dang thu bang WMI..."
            $loggedInUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
            $username = $loggedInUser.Split('\')[-1]
            $sid = (Get-WmiObject Win32_UserAccount | Where-Object { $_.Name -eq $username }).SID
            Write-Log "User dang login (WMI): $username"
            Write-Log "SID (WMI): $sid"
        }

        if (-not $sid) {
            throw "Khong xac dinh duoc SID cua user $username."
        }

        # Tìm các dịch vụ cần đặt quyền
        $services = Get-Service | Where-Object { $_.Name -match "rAgent" -or $_.Name -match "FARCARDS" }
        if (-not $services -or $services.Count -eq 0) {
            Write-Log "Khong tim thay dich vu nao co ten khop 'rAgent' hoac 'FARCARDS'."
            return
        }

        # Xây chuỗi Security Descriptor (DACL) cấp quyền cho SID người dùng
        # (A;;RPWPCR;;;{SID}) => Allow: Read Property, Write Property, Change Config, v.v. trên service
        foreach ($svc in $services) {
            Write-Log ("`nDich vu: {0}" -f $svc.Name)
            $sdString = ("D:(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;RPWPCR;;;{0})" -f $sid)
            Write-Log "DACL moi: $sdString"

            try {
                # Dùng sc.exe sdset
                $proc = Start-Process -FilePath "sc.exe" -ArgumentList @("sdset",$svc.Name,$sdString) -PassThru -WindowStyle Hidden -Wait
                if ($proc.ExitCode -eq 0) {
                    Write-Log "=> Dat quyen thanh cong cho dich vu '$($svc.Name)'."
                } else {
                    Write-Log "=> Loi: sc.exe tra ve ExitCode $($proc.ExitCode) cho dich vu '$($svc.Name)'."
                }
            } catch {
                Write-Log "=> Loi khi thiet lap quyen cho '$($svc.Name)': $($_.Exception.Message)"
            }
        }

        Write-Log "Hoan tat thiet lap quyen Dcorp."
    } catch {
        Write-Log "Loi tong the: $($_.Exception.Message)"
    }
})


# Xử lý khi nhấn nút Open Printers
$buttonOpenPrinters.Add_Click({
    try {
        Write-Log "Dang mo cua so Devices and Printers..."
        Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
        Write-Log "Da mo Devices and Printers."
    } catch {
        Write-Log "Loi khi mo Devices and Printers: $($_.Exception.Message)"
    }
})


#=== Sự kiện Form Shown: nạp user & kiểm tra dịch vụ ===
$form.Add_Shown({
    if (-not (Test-IsAdmin)) {
        Write-Log "Canh bao: nen chay tool voi quyen Admin de tranh loi khi tao/sua user."
    }
    Ensure-ProfSvcAndDefaultProfileOK
    LoadUsers
})

#=== Chạy ứng dụng WinForms (bảo đảm form hiển thị) ===
[System.Windows.Forms.Application]::Run($form)
