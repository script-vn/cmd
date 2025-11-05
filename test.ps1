Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Tool Add User & Deny Zalo"
$form.Size = New-Object System.Drawing.Size(520,380)
$form.StartPosition = "CenterScreen"

# Label tạo user
$labelNewUser = New-Object System.Windows.Forms.Label
$labelNewUser.Text = "Add new User:"
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
$buttonCreateUser.Text = "Create user"
$buttonCreateUser.Location = New-Object System.Drawing.Point(390,40)
$buttonCreateUser.Size = New-Object System.Drawing.Size(80,25)
$form.Controls.Add($buttonCreateUser)

# Label chọn user
$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Text = "Select user:"
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
$buttonExecute.Text = "Run Deny Zalo"
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


# Nút Driver Printer
$buttonPrinterDriver = New-Object System.Windows.Forms.Button
$buttonPrinterDriver.Text = "Driver Printer"
$buttonPrinterDriver.Location = New-Object System.Drawing.Point(350,320)
$buttonPrinterDriver.Size = New-Object System.Drawing.Size(140,30)
$form.Controls.Add($buttonPrinterDriver)


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
    Write-Log "Da tai danh sach user."
    if ($selectUser -and $users -contains $selectUser) {
        $comboBox.SelectedItem = $selectUser
        Write-Log "Da chon user mac dinh: $selectUser"
    }
}

LoadUsers

# Tạo user mới
$buttonCreateUser.Add_Click({
    $newUser = $textUserName.Text.Trim()
    $fullName = $textFullName.Text.Trim()

    if (-not $newUser) {
        [System.Windows.Forms.MessageBox]::Show("Vui long nhap User name.", "Thong bao", "OK", "Warning")
        return
    }

    try {
        Write-Log "Dang tao user '$newUser'..."
        $password = ConvertTo-SecureString "123" -AsPlainText -Force
        New-LocalUser -Name $newUser -FullName $fullName -Password $password
        Set-LocalUser -Name $newUser -PasswordNeverExpires $true
        net user $newUser /passwordchg:no
        Add-LocalGroupMember -Group "Users" -Member $newUser
        LoadUsers $newUser
        $textUserName.Text = ""
        $textFullName.Text = ""
        Write-Log "Da tao user '$newUser' thanh cong."
    } catch {
        Write-Log "Loi khi tao user: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Loi khi tao user: $($_.Exception.Message)", "Loi", "OK", "Error")
    }
})

# Thực hiện chặn Zalo
$buttonExecute.Add_Click({
    $userName = $comboBox.SelectedItem
    if (-not $userName) {
        [System.Windows.Forms.MessageBox]::Show("Vui long chon user.", "Thong bao", "OK", "Warning")
        return
    }

    try {
        Write-Log "Dang cap quyen admin cho user '$userName'..."
        Add-LocalGroupMember -Group "Administrators" -Member $userName
        Write-Log "Da cap quyen admin."

        $userProfilePath = "C:\Users\$userName"
        $ntUserDatPath = "$userProfilePath\NTUSER.DAT"
        $hiveName = "HKU\TempHive"

        if (-not (Test-Path $ntUserDatPath)) {
            Write-Log "NTUSER.DAT chua ton tai. Dang tao profile bang notepad..."
            $cred = Get-Credential -UserName $userName -Message "Nhap mat khau de profile cho user $userName"
            $np = Start-Process -FilePath "notepad.exe" -Credential $cred -PassThru
            Stop-Process -Id $np.Id -Force
            Write-Log "Da mo va tat notepad de tao profile."
        }

        if (Test-Path $ntUserDatPath) {
            Write-Log "Da tim thay NTUSER.DAT. Dang ghi registry chan Zalo..."
            reg load $hiveName $ntUserDatPath | Out-Null

            reg add "$hiveName\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v DisallowRun /t REG_DWORD /d 1 /f | Out-Null
            reg add "$hiveName\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" /v 1 /t REG_SZ /d Zalo.exe /f | Out-Null
            reg add "$hiveName\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" /v 2 /t REG_SZ /d ZaloUpdate.exe /f | Out-Null

            reg unload $hiveName | Out-Null
            Write-Log "Da ghi registry chan Zalo cho user '$userName'."

            Write-Log "Dang thu hoi quyen admin cua user '$userName'..."
            Remove-LocalGroupMember -Group "Administrators" -Member $userName
            Write-Log "Da thu hoi quyen admin."

            Start-Process "lusrmgr.msc"
            Write-Log "Da mo lusrmgr.msc de chinh sua thu cong neu can."

            [System.Windows.Forms.MessageBox]::Show("Da chan Zalo cho user '$userName' thanh cong.", "Thanh cong", "OK", "Information")
        } else {
            throw "Khong the tao NTUSER.DAT cho user $userName."
        }
    } catch {
        Write-Log "Loi: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Loi: $($_.Exception.Message)", "Loi", "OK", "Error")
    }
})

# Xử lý khi nhấn nút Set IP LAN & DNS
$buttonSetIP.Add_Click({
    try {
        Write-Log "Dang cau hinh IP LAN va DNS..."

        $interface = Get-NetAdapter | Where-Object { $_.Status -eq "Up" } | Select-Object -First 1
        if (-not $interface) {
            throw "Khong tim thay card mang dang hoat dong."
        }

        $interfaceName = $interface.Name
        Write-Log "Da chon card mang: $interfaceName"

        # Set IP tĩnh
        New-NetIPAddress -InterfaceAlias $interfaceName -IPAddress "192.168.1.11" -PrefixLength 24 -DefaultGateway "192.168.1.1" -ErrorAction Stop
        Write-Log "Da dat IP: 192.168.1.11, Gateway: 192.168.1.1"

        # Set DNS
        Set-DnsClientServerAddress -InterfaceAlias $interfaceName -ServerAddresses ("8.8.8.8","8.8.4.4") -ErrorAction Stop
        Write-Log "Da dat DNS: 8.8.8.8, 8.8.4.4"

        [System.Windows.Forms.MessageBox]::Show("Da cau hinh IP và DNS thanh cong.", "Thanh cong", "OK", "Information")
    } catch {
        Write-Log "Loi khi cau hinh IP/DNS: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Loi: $($_.Exception.Message)", "Loi", "OK", "Error")
    }
})


# Xử lý khi nhấn nút Driver Printer
$buttonPrinterDriver.Add_Click({
    try {
        Write-Log "Dang tai xuong file driver may in..."
        $url = "https://sasinvn-my.sharepoint.com/:u:/g/personal/nam_tran_sasin_vn/EaSWlvxwl7RIljE-aEhCU3ABAKTlJG2jTIFv6hJxR9-5xA?download=1"
        $destination = "D:\DriverPrinter.printerExport"
        Invoke-WebRequest -Uri $url -OutFile $destination
        Write-Log "Da tai xuong file driver tai $destination"

        # Mở phần mềm PrintBrmUi.exe
        $printBrmPath = "C:\Windows\System32\PrintBrmUi.exe"
        if (Test-Path $printBrmPath) {
            Start-Process $printBrmPath
            Write-Log "Da mo PrintBrmUi.exe de import driver."
        } else {
            Write-Log "Khong tim thay PrintBrmUi.exe tai $printBrmPath"
            [System.Windows.Forms.MessageBox]::Show("Khong tim thay PrintBrmUi.exe", "Loi", "OK", "Error")
        }
    } catch {
        Write-Log "Loi khi tai hoac mo driver: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Loi: $($_.Exception.Message)", "Loi", "OK", "Error")
    }
})


[void]$form.ShowDialog()
