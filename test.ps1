Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Quản lý User & Chặn Zalo"
$form.Size = New-Object System.Drawing.Size(520,380)
$form.StartPosition = "CenterScreen"

# Label tạo user
$labelNewUser = New-Object System.Windows.Forms.Label
$labelNewUser.Text = "Tạo user mới:"
$labelNewUser.Location = New-Object System.Drawing.Point(20,20)
$labelNewUser.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($labelNewUser)

# Label User name
$labelUserName = New-Object System.Windows.Forms.Label
$labelUserName.Text = "User name:"
$labelUserName.Location = New-Object System.Drawing.Point(120,20)
$labelUserName.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($labelUserName)

# TextBox username
$textUserName = New-Object System.Windows.Forms.TextBox
$textUserName.Location = New-Object System.Drawing.Point(120,40)
$textUserName.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($textUserName)

# Label Full name
$labelFullName = New-Object System.Windows.Forms.Label
$labelFullName.Text = "Full name:"
$labelFullName.Location = New-Object System.Drawing.Point(230,20)
$labelFullName.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($labelFullName)

# TextBox full name
$textFullName = New-Object System.Windows.Forms.TextBox
$textFullName.Location = New-Object System.Drawing.Point(230,40)
$textFullName.Size = New-Object System.Drawing.Size(150,20)
$form.Controls.Add($textFullName)

# Nút tạo user
$buttonCreateUser = New-Object System.Windows.Forms.Button
$buttonCreateUser.Text = "Tạo user"
$buttonCreateUser.Location = New-Object System.Drawing.Point(390,40)
$buttonCreateUser.Size = New-Object System.Drawing.Size(80,25)
$form.Controls.Add($buttonCreateUser)

# Label chọn user
$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Text = "Chọn user:"
$labelUser.Location = New-Object System.Drawing.Point(20,80)
$labelUser.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($labelUser)

# ComboBox user
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(120,80)
$comboBox.Size = New-Object System.Drawing.Size(200,20)
$comboBox.DropDownStyle = "DropDownList"
$form.Controls.Add($comboBox)

# Nút chặn Zalo
$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Text = "Thực hiện chặn Zalo"
$buttonExecute.Location = New-Object System.Drawing.Point(150,110)
$buttonExecute.Size = New-Object System.Drawing.Size(180,30)
$form.Controls.Add($buttonExecute)

# TextBox ghi log
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(20,150)
$logBox.Size = New-Object System.Drawing.Size(470,160)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$form.Controls.Add($logBox)

# Nút Set IP LAN và DNS
$buttonSetIP = New-Object System.Windows.Forms.Button
$buttonSetIP.Text = "Set IP LAN & DNS"
$buttonSetIP.Location = New-Object System.Drawing.Point(150,320)
$buttonSetIP.Size = New-Object System.Drawing.Size(180,30)
$form.Controls.Add($buttonSetIP)

# Hàm ghi log
function Write-Log($message) {
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $logBox.AppendText("[$timestamp] $message`r`n")
}

# Load danh sách user và chọn mặc định
function LoadUsers($selectUser = $null) {
    $comboBox.Items.Clear()
    $users = Get-LocalUser | Where-Object { $_.Enabled -eq $true -and $_.Name -ne "Admin" } | Select-Object -ExpandProperty Name
    $comboBox.Items.AddRange($users)
    Write-Log "Đã tải danh sách user."
    if ($selectUser -and $users -contains $selectUser) {
        $comboBox.SelectedItem = $selectUser
        Write-Log "Đã chọn user mặc định: $selectUser"
    }
}

LoadUsers

# Tạo user mới
$buttonCreateUser.Add_Click({
    $newUser = $textUserName.Text.Trim()
    $fullName = $textFullName.Text.Trim()

    if (-not $newUser) {
        [System.Windows.Forms.MessageBox]::Show("Vui lòng nhập User name.", "Thông báo", "OK", "Warning")
        return
    }

    try {
        Write-Log "Đang tạo user '$newUser'..."
        $password = ConvertTo-SecureString "123" -AsPlainText -Force
        New-LocalUser -Name $newUser -FullName $fullName -Password $password
        Set-LocalUser -Name $newUser -PasswordNeverExpires $true
        net user $newUser /passwordchg:no
        Add-LocalGroupMember -Group "Users" -Member $newUser
        LoadUsers $newUser
        $textUserName.Text = ""
        $textFullName.Text = ""
        Write-Log "Đã tạo user '$newUser' thành công và xóa thông tin nhập."
    } catch {
        Write-Log "Lỗi khi tạo user: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Lỗi khi tạo user: $($_.Exception.Message)", "Lỗi", "OK", "Error")
    }
})

# Thực hiện chặn Zalo
$buttonExecute.Add_Click({
    $userName = $comboBox.SelectedItem
    if (-not $userName) {
        [System.Windows.Forms.MessageBox]::Show("Vui lòng chọn user.", "Thông báo", "OK", "Warning")
        return
    }

    try {
        Write-Log "Đang cấp quyền admin cho user '$userName'..."
        Add-LocalGroupMember -Group "Administrators" -Member $userName
        Write-Log "Đã cấp quyền admin."

        $userProfilePath = "C:\Users\$userName"
        $ntUserDatPath = "$userProfilePath\NTUSER.DAT"
        $hiveName = "HKU\TempHive"

        if (-not (Test-Path $ntUserDatPath)) {
            Write-Log "NTUSER.DAT chưa tồn tại. Đang tạo profile bằng notepad..."
            $cred = Get-Credential -UserName $userName -Message "Nhập mật khẩu để tạo profile cho user $userName"
            $np = Start-Process -FilePath "notepad.exe" -Credential $cred -PassThru
            Stop-Process -Id $np.Id -Force
            Write-Log "Đã mở và tắt notepad ngay lập tức để tạo profile."
        }

        if (Test-Path $ntUserDatPath) {
            Write-Log "Đã tìm thấy NTUSER.DAT. Đang ghi registry chặn Zalo..."
            reg load $hiveName $ntUserDatPath | Out-Null

            reg add "$hiveName\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v DisallowRun /t REG_DWORD /d 1 /f | Out-Null
            reg add "$hiveName\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" /v 1 /t REG_SZ /d Zalo.exe /f | Out-Null
            reg add "$hiveName\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" /v 2 /t REG_SZ /d ZaloUpdate.exe /f | Out-Null

            reg unload $hiveName | Out-Null
            Write-Log "Đã ghi registry chặn Zalo cho user '$userName'."

            Write-Log "Đang thu hồi quyền admin của user '$userName'..."
            Remove-LocalGroupMember -Group "Administrators" -Member $userName
            Write-Log "Đã thu hồi quyền admin."

            Start-Process "lusrmgr.msc"
            Write-Log "Đã mở lusrmgr.msc để chỉnh sửa thủ công nếu cần."

            [System.Windows.Forms.MessageBox]::Show("Đã chặn Zalo cho user '$userName' thành công.", "Thành công", "OK", "Information")
        } else {
            throw "Không thể tạo NTUSER.DAT cho user $userName."
        }
    } catch {
        Write-Log "Lỗi: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Lỗi: $($_.Exception.Message)", "Lỗi", "OK", "Error")
    }
})

# Xử lý khi nhấn nút Set IP LAN & DNS
$buttonSetIP.Add_Click({
    try {
        Write-Log "Đang cấu hình IP LAN và DNS..."

        $interface = Get-NetAdapter | Where-Object { $_.Status -eq "Up" } | Select-Object -First 1
        if (-not $interface) {
            throw "Không tìm thấy card mạng đang hoạt động."
        }

        $interfaceName = $interface.Name
        Write-Log "Đã chọn card mạng: $interfaceName"

        # Set IP tĩnh
        New-NetIPAddress -InterfaceAlias $interfaceName -IPAddress "192.168.1.11" -PrefixLength 24 -DefaultGateway "192.168.1.1" -ErrorAction Stop
        Write-Log "Đã đặt IP: 192.168.1.11, Gateway: 192.168.1.1"

        # Set DNS
        Set-DnsClientServerAddress -InterfaceAlias $interfaceName -ServerAddresses ("8.8.8.8","8.8.4.4") -ErrorAction Stop
        Write-Log "Đã đặt DNS: 8.8.8.8, 8.8.4.4"

        [System.Windows.Forms.MessageBox]::Show("Đã cấu hình IP và DNS thành công.", "Thành công", "OK", "Information")
    } catch {
        Write-Log "Lỗi khi cấu hình IP/DNS: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Lỗi: $($_.Exception.Message)", "Lỗi", "OK", "Error")
    }
})

[void]$form.ShowDialog()
