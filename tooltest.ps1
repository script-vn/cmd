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
$form.Text = "Tool Support IT - Dev by NamIT"
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
$buttonResources.Location = New-Object System.Drawing.Point(15, 280)
$buttonResources.Size = New-Object System.Drawing.Size(146, 30)
$form.Controls.Add($buttonResources)

# Nút Set IP LAN và DNS
$buttonSetIP = New-Object System.Windows.Forms.Button
$buttonSetIP.Text = "Set IP LAN & DNS"
$buttonSetIP.Location = New-Object System.Drawing.Point(181,280)
$buttonSetIP.Size = New-Object System.Drawing.Size(146,30)
$form.Controls.Add($buttonSetIP)

# Nút Driver Printer
$buttonPrinterDriver = New-Object System.Windows.Forms.Button
$buttonPrinterDriver.Text = "Setup Driver Full Zy303"
$buttonPrinterDriver.Location = New-Object System.Drawing.Point(347,280)
$buttonPrinterDriver.Size = New-Object System.Drawing.Size(146,30)
$form.Controls.Add($buttonPrinterDriver)


# Nút Active Services Dcorp
$buttonSetPerm = New-Object System.Windows.Forms.Button
$buttonSetPerm.Text = "Active Services Dcorp"
$buttonSetPerm.Location = New-Object System.Drawing.Point(15, 320)
$buttonSetPerm.Size = New-Object System.Drawing.Size(146, 30)
$form.Controls.Add($buttonSetPerm)


# Nút Open Printers
$buttonOpenPrinters = New-Object System.Windows.Forms.Button
$buttonOpenPrinters.Text = "Open Printers"
$buttonOpenPrinters.Location = New-Object System.Drawing.Point(181, 320)
$buttonOpenPrinters.Size = New-Object System.Drawing.Size(146, 30)
$form.Controls.Add($buttonOpenPrinters)


# Nút Change UAC Settings
$buttonUAC = New-Object System.Windows.Forms.Button
$buttonUAC.Text = "Change UAC Settings"
$buttonUAC.Location = New-Object System.Drawing.Point(347, 320)
$buttonUAC.Size = New-Object System.Drawing.Size(146, 30)
$form.Controls.Add($buttonUAC)


# Nút Edit Power Plan
$buttonEditPower = New-Object System.Windows.Forms.Button
$buttonEditPower.Text = "Edit Power Plan"
$buttonEditPower.Location = New-Object System.Drawing.Point(15, 360)
$buttonEditPower.Size = New-Object System.Drawing.Size(146, 30)
$form.Controls.Add($buttonEditPower)


# Nút App volume and device preferences
$buttonAppVolume = New-Object System.Windows.Forms.Button
$buttonAppVolume.Text = "App Volume & Device"
$buttonAppVolume.Location = New-Object System.Drawing.Point(181, 360)
$buttonAppVolume.Size = New-Object System.Drawing.Size(146, 30)
$form.Controls.Add($buttonAppVolume)


# Nút SFC /SCANNOW
$buttonSFC = New-Object System.Windows.Forms.Button
$buttonSFC.Text = "Run SFC /scannow"
$buttonSFC.Location = New-Object System.Drawing.Point(347, 360)
$buttonSFC.Size = New-Object System.Drawing.Size(146, 30)
$form.Controls.Add($buttonSFC)


# Nút Setup May in Brother HL2100D
$buttonBrother = New-Object System.Windows.Forms.Button
$buttonBrother.Text = "Setup Brother HL2100D"
$buttonBrother.Location = New-Object System.Drawing.Point(15, 400)
$buttonBrother.Size = New-Object System.Drawing.Size(146, 30)
$form.Controls.Add($buttonBrother)


# Nút Check USB Printer Zy303
$buttonUSB80C = New-Object System.Windows.Forms.Button
$buttonUSB80C.Text = "USB Printer Zy303"
$buttonUSB80C.Location = New-Object System.Drawing.Point(181, 400)
$buttonUSB80C.Size = New-Object System.Drawing.Size(146, 30)
$form.Controls.Add($buttonUSB80C)


# Nút Setup All Printer Zy303
$buttonZY303 = New-Object System.Windows.Forms.Button
$buttonZY303.Text = "Set All Zy303"
$buttonZY303.Location = New-Object System.Drawing.Point(347, 400)
$buttonZY303.Size = New-Object System.Drawing.Size(146, 30)
$form.Controls.Add($buttonZY303)


<#
.SYNOPSIS
  Kiểm tra & sửa lỗi khiến user mới đăng nhập xong bị "Signing out" trên Windows 10.
.DESCRIPTION
  - Đảm bảo dịch vụ User Profile Service (ProfSvc) ở trạng thái Automatic & Running.
  - Sửa các khóa Winlogon (Shell, Userinit) về giá trị mặc định an toàn.
  - Dọn & sửa các SID .bak trong ProfileList theo hướng dẫn chuẩn.
  - Đặt lại quyền NTFS cơ bản cho C:\Users và C:\Users\Default (kế thừa bật).
  - Kiểm tra dung lượng trống, trạng thái Default Profile (NTUSER.DAT).
  - Xuất Local Security Policy để cảnh báo nếu có "Deny log on locally".
  - Backup registry trước khi chỉnh sửa.
  - Ghi log đầy đủ vào %TEMP%\FixUserProfile_YYYYMMDD_HHMMSS.log

.NOTES
  Chạy với quyền Administrator. Áp dụng cho Windows 10.
#>
Function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR")][string]$Level = "INFO"
    )
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}
# endregion
# region Backup Registry
Function Backup-RegistryKey {
    param([string]$RegPath, [string]$FileName)
    try {
        $full = Join-Path $backupDir $FileName
        # Dùng reg.exe để export
        & reg.exe export $RegPath $full /y | Out-Null
#      Write-Log -Message ("Done backup: {0} -> {1}" -f $RegPath, $full)
    } catch {
       Write-Log -Message ("Backup Fail {0}: {1}" -f $RegPath, $_) -Level "WARN"
    }
}
# endregion
# region Dịch vụ User Profile Service
Function Ensure-UserProfileService {
    try {
        $svc = Get-Service -Name "ProfSvc" -ErrorAction Stop
        if ($svc.StartType -ne "Automatic") {
            Write-Log -Message "Set ProfSvc StartupType -> Automatic"
            Set-Service -Name "ProfSvc" -StartupType Automatic
        }
        if ($svc.Status -ne "Running") {
            Write-Log -Message "Start up ProfSvc"
            Start-Service -Name "ProfSvc"
        }
        $svc = Get-Service -Name "ProfSvc"
#      Write-Log -Message ("ProfSvc: {0}, StartupType: {1}" -f $svc.Status, $svc.StartType)
    } catch {
       Write-Log -Message ("Notfound/Not control ProfSvc: {0}" -f $_) -Level "ERROR"
        throw
    }
}
# endregion
# region Sửa khóa Winlogon (Userinit, Shell)
Function Fix-WinlogonKeys {
    $key = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    try {
        # Giá trị mặc định an toàn
        $userinitDefault = "$env:SystemRoot\system32\userinit.exe,"
        $shellDefault = "explorer.exe"

        $currUserinit = (Get-ItemProperty -Path $key -Name "Userinit" -ErrorAction SilentlyContinue).Userinit
        $currShell = (Get-ItemProperty -Path $key -Name "Shell" -ErrorAction SilentlyContinue).Shell

        if ($currUserinit -ne $userinitDefault) {
#         Write-Log -Message ("Fix Winlogon\Userinit -> {0} (trước: '{1}')" -f $userinitDefault, $currUserinit)
            Set-ItemProperty -Path $key -Name "Userinit" -Value $userinitDefault
        } else {
#          Write-Log -Message ("Winlogon\Userinit OK: {0}" -f $currUserinit)
        }

        if ($currShell -ne $shellDefault) {
#         Write-Log -Message ("Fix Winlogon\Shell -> {0} (trước: '{1}')" -f $shellDefault, $currShell)
            Set-ItemProperty -Path $key -Name "Shell" -Value $shellDefault
        } else {
#          Write-Log -Message ("Winlogon\Shell OK: {0}" -f $currShell)
        }
    } catch {
        Write-Log -Message ("Error set Winlogon: {0}" -f $_) -Level "ERROR"
        throw
    }
}
# endregion
# region Kiểm tra dung lượng trống ổ hệ thống
Function Check-SystemDriveFreeSpace {
    $sysDriveName = ($env:SystemDrive).TrimEnd(':')
    $sysDrive = Get-PSDrive -Name $sysDriveName
    $freeGB = [math]::Round($sysDrive.Free/1GB,2)
#   Write-Log -Message ("Free Space {0}: {1} GB" -f $sysDrive.Name, $freeGB)
    if ($freeGB -lt 2) {
#       Write-Log -Message "Warring < 2GB do Profile not create." -Level "WARN"
    }
}
# endregion
# region Đặt lại quyền NTFS cho C:\Users và C:\Users\Default
Function Reset-UsersAcl {
    param([string]$Path)
    if (-not (Test-Path $Path)) { Write-Log -Message ("Khong Ton Tai: {0}" -f $Path) -Level "WARN"; return }
    try {
        Write-Log -Message ("Bat Ke Thua ACL cho: {0}" -f $Path)
        & icacls $Path /inheritance:e | Out-Null

        Write-Log -Message ("Cap Quyen Co ban cho: {0}" -f $Path)
        & icacls $Path /grant "SYSTEM:(F)" "BUILTIN\Administrators:(F)" "BUILTIN\Users:(RX)" /T /C | Out-Null

#      Write-Log -Message ("Set ACL Default xong cho: {0}" -f $Path)
    } catch {
        Write-Log -Message ("Error set ACL cho {0}: {1}" -f $Path, $_) -Level "ERROR"
    }
}
# endregion
# region Kiểm tra Default Profile (NTUSER.DAT)
Function Test-DefaultProfileHealth {
    $def = "C:\Users\Default"
    if (-not (Test-Path $def)) {
#       Write-Log -Message ("Thieu thu muc Default Profile: {0}" -f $def) -Level "ERROR"
#       Write-Log -Message "Huong da: Khoi phuc Default tu nguon cai dat/may kha cung version. Khong tu sao chep tu systemprofile." -Level "WARN"
        return
    }
    $ntuser = Join-Path $def "NTUSER.DAT"
    if (-not (Test-Path $ntuser)) {
#      Write-Log -Message "Thieu NTUSER.DAT trong Default Profile -> nguyen nhan pho bien gay 'Signing out'." -Level "ERROR"
#       Write-Log -Message "Huong dan: Sao chep NTUSER.DAT hop le tu may Windows 10 cung build hoac media cai dat." -Level "WARN"
    } else {
        $sizeMB = [math]::Round((Get-Item $ntuser).Length/1MB,2)
#       Write-Log -Message ("NTUSER.DAT ton tai, kich thuoc: {0} MB" -f $sizeMB)
        if ($sizeMB -lt 0.2) {
#           Write-Log -Message "Canh bao: NTUSER.DAT kich thuoc bat thuong (<0.2MB) -> co the hong." -Level "WARN"
        }
    }
}
# endregion
# region Sửa ProfileList .bak & dọn entry mồ côi
Function Fix-ProfileListBak {
    $base = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    $keys = Get-ChildItem -Path $base -ErrorAction Stop

    foreach ($k in $keys) {
        $sid = Split-Path $k.PSChildName -Leaf
        $isBak = $sid.EndsWith(".bak")
        $plainSid = $sid -replace "\.bak$",""

        if ($isBak) {
            $nonBak = Join-Path $base $plainSid
            if (Test-Path $nonBak) {
#            Write-Log -Message ("Phat hien cap SID: {0} và {1} -> tien hanh sua theo chuan" -f $sid, $plainSid)
                try {
                    Set-ItemProperty -Path $k.PSPath -Name "RefCount" -Value 0 -ErrorAction SilentlyContinue
                    Set-ItemProperty -Path $k.PSPath -Name "State" -Value 0 -ErrorAction SilentlyContinue

                    # Đổi tên key: nonBak -> .tmp, bak -> nonBak, xóa .tmp
                    $tmp = $plainSid + ".tmp"
                    & reg.exe rename ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" + $plainSid) $tmp /f | Out-Null
                    & reg.exe rename ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" + $sid) $plainSid /f | Out-Null
                    Remove-Item -Path (Join-Path $base $tmp) -Recurse -Force -ErrorAction SilentlyContinue
#                  Write-Log -Message ("Da xu ly SID .bak -> {0}" -f $plainSid)
                } catch {
#                   Write-Log -Message ("Loi doi ten SID .bak: {0}" -f $_) -Level "ERROR"
                }
            } else {
                try {
                    & reg.exe rename ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" + $sid) $plainSid /f | Out-Null
#                 Write-Log -Message ("Doi ten {0} -> {1} (Khong co ban nonBak)" -f $sid, $plainSid)
                } catch {
#                  Write-Log -Message ("Loi rename {0}: {1}" -f $sid, $_) -Level "ERROR"
                }
            }
        }

        # Dọn các profile mồ côi (đường dẫn không tồn tại) trừ các SID hệ thống
        $pip = (Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue).ProfileImagePath
        if ($pip) {
            $specialSids = @("S-1-5-18","S-1-5-19","S-1-5-20")
            $sidPlain = $sid -replace "\.bak$",""
            $skip = $specialSids | Where-Object { $sidPlain.StartsWith($_) }
            if (-not $skip) {
                if (-not (Test-Path $pip)) {
                    Write-Log -Message ("Profile SID {0} tro den '{1}' khong ton tai -> xoa de tranh xung dot" -f $sidPlain, $pip)
                    try {
                        Remove-Item -Path $k.PSPath -Recurse -Force
#                    Write-Log -Message ("Da xoa SID {0} mo coi" -f $sidPlain)
                    } catch {
#                      Write-Log -Message ("Khong the xoa SID {0}: {1}" -f $sidPlain, $_) -Level "WARN"
                    }
                }
            }
        }
    }
}
# endregion
# region Xuất & cảnh báo 'Deny log on locally' (nếu có)
Function Check-DenyLogonLocally {
    try {
        $cfg = Join-Path $backupDir "secpol_$timestamp.inf"
        secedit /export /cfg $cfg | Out-Null
        if (Test-Path $cfg) {
            $lines = Get-Content $cfg
            $deny = $lines | Where-Object { $_ -match "^SeDenyInteractiveLogonRight\s*=" }
            if ($deny) {
#              Write-Log -Message ("Phat hien cau hinh 'Deny log on locally': {0}" -f ($deny -join "; ")) -Level "WARN"
#               Write-Log -Message "Neu chua 'Users' hoac tai khoan muc tieu, can go chinh sach nay trong Local Security Policy (secpol.msc)." -Level "WARN"
            } else {
#              Write-Log -Message "Khong thay cau hinh 'Deny log on locally'."
            }
        }
    } catch {
        Write-Log -Message ("Khong the xuat Local Security Policy: {0}" -f $_) -Level "WARN"
    }
}
# endregion
# region Khuyến nghị SFC/DISM (tùy chọn)
Function Run-SystemIntegrityCheck {
    param([switch]$RunNow)
    if ($RunNow) {
        try {
#           Write-Log -Message "Chay DISM /RestoreHealth (co the mat thoi gian)..."
            & DISM.exe /Online /Cleanup-Image /RestoreHealth | Tee-Object -FilePath (Join-Path $backupDir "DISM_RestoreHealth.log")
#          Write-Log -Message "Chay SFC /SCANNOW..."
            & sfc.exe /scannow | Tee-Object -FilePath (Join-Path $backupDir "SFC_scannow.log")
        } catch {
           Write-Log -Message ("DISM/SFC loi: {0}" -f $_) -Level "WARN"
        }
    } else {
 #       Write-Log -Message "Khuyen nghi chay DISM & SFC neu loi van con: DISM /Online /Cleanup-Image /RestoreHealth ; sfc /scannow" -Level "INFO"
    }
}




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


#=== Sự kiện nút Create user ===
$buttonCreateUser.Add_Click({

    # Bỏ try ngoài hoặc thêm catch; ở đây mình bỏ để đơn giản.
    $newUser  = $textUserName.Text.Trim()
    $fullName = $textFullName.Text.Trim()

    if (-not $newUser) {
        [System.Windows.Forms.MessageBox]::Show(
            "Vui long nhap User name.",
            "Thong bao",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if (-not (Test-IsAdmin)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Can Admin rights de tao user.",
            "Thong bao",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    try {
        Write-Log -Message "Dang tao user '$newUser'..."
        $password = ConvertTo-SecureString "123" -AsPlainText -Force

        New-LocalUser -Name $newUser -FullName $fullName -Password $password -ErrorAction Stop
        Set-LocalUser -Name $newUser -PasswordNeverExpires $true

        # Sửa &amp; -> &
        & net.exe user $newUser /passwordchg:no | Out-Null

        Add-LocalGroupMember -Group "Users" -Member $newUser -ErrorAction SilentlyContinue

        LoadUsers $newUser
        $textUserName.Text = ""
        $textFullName.Text = ""
        Write-Log -Message "Da tao user '$newUser' thanh cong."
    }
    catch {
        Write-Log -Message "Loi khi tao user: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Loi khi tao user: $($_.Exception.Message)",
            "Loi",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }

    #region Chuẩn bị log & kiểm tra quyền
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $LogFile   = Join-Path $env:TEMP "FixUserProfile_$timestamp.log"
    $ErrorActionPreference = "Stop"

    Write-Log -Message "=== Bat dau FixUserProfile ==="

    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
        ).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
        Write-Log -Message "Script can chay voi quyen Administrator. Thoat..." -Level "ERROR"
        throw "Vui long mo PowerShell 'Run as Administrator'."
    }

    # Xác nhận phiên bản Windows 10
    $os = Get-CimInstance Win32_OperatingSystem
    Write-Log -Message ("He dieu hanh: {0} {1}" -f $os.Caption, $os.Version)

    $backupDir = Join-Path $env:TEMP "FixUserProfile_Backup_$timestamp"
    New-Item -Path $backupDir -ItemType Directory -Force | Out-Null

    Backup-RegistryKey -RegPath "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"   -FileName "Winlogon.reg"
    Backup-RegistryKey -RegPath "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" -FileName "ProfileList.reg"

    Ensure-UserProfileService
    Fix-WinlogonKeys
    Check-SystemDriveFreeSpace
  # Reset-UsersAcl -Path "C:\Users"
  # Reset-UsersAcl -Path "C:\Users\Default"
    Test-DefaultProfileHealth
    Fix-ProfileListBak
    Check-DenyLogonLocally
    Run-SystemIntegrityCheck
    #endregion

    Write-Log -Message ("=== Hoan tat. Log luu tai: {0} ===" -f $LogFile)
    Write-Host "`n-> Vui long khoi dong lai may va thu dang nhap tai khoan moi."
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

# Handler: mở cửa sổ đổi mức UAC
$buttonUAC.Add_Click({
    try {
        Write-Log "Dang mo cua so 'Change User Account Control settings'..."

        $uacExe = Join-Path $env:SystemRoot "System32\UserAccountControlSettings.exe"
        if (Test-Path $uacExe) {
            Start-Process $uacExe -ErrorAction Stop
            Write-Log "Da mo UAC bang UserAccountControlSettings.exe."
        } else {
            # Fallback qua Control Panel canonical name
            Write-Log "Khong thay UserAccountControlSettings.exe, dang mo qua Control Panel..."
            Start-Process "control.exe" -ArgumentList "/name Microsoft.UserAccountControlSettings" -ErrorAction Stop
            Write-Log "Da mo UAC bang Control Panel."
        }
    } catch {
        Write-Log "Loi khi mo UAC: $($_.Exception.Message)"
    }
})


# Handler: mở trang Edit Power Plan / Power Options
$buttonEditPower.Add_Click({
    try {
        Write-Log "Dang mo trang chinh sua Power Plan (Power Options)..."

        # Cách 1: mở Power Options theo canonical name (ổn định trên nhiều bản Windows)
        try {
            Start-Process "control.exe" -ArgumentList "/name Microsoft.PowerOptions" -ErrorAction Stop
            Write-Log "Da mo Power Options qua Control Panel (Microsoft.PowerOptions)."
        } catch {
            Write-Log "Khong mo duoc qua canonical name, thu mo bang powercfg.cpl..."
            # Cách 2: mở trực tiếp applet powercfg.cpl (fallback)
            Start-Process "control.exe" -ArgumentList "powercfg.cpl" -ErrorAction Stop
            Write-Log "Da mo Power Options bang powercfg.cpl."
        }

        # (Tuỳ chọn nâng cao) Nếu muốn cố gắng mở trang 'Edit plan settings' của plan đang dùng:
        # Lưu ý: Trang Edit plan settings phụ thuộc plan cụ thể và giao diện Control Panel,
        # nên cách an toàn là mở Power Options để người dùng chọn "Change plan settings".
        # Một số bản Windows hỗ trợ tham số phụ cho powercfg.cpl nhưng không nhất quán.
    } catch {
        Write-Log "Loi khi mo Edit Power Plan: $($_.Exception.Message)"
    }
})


# Handler: mở trang App volume and device preferences
$buttonAppVolume.Add_Click({
    try {
        Write-Log "Dang mo 'App volume and device preferences'..."
        Start-Process "ms-settings:apps-volume" -ErrorAction Stop
        Write-Log "Da mo trang App volume & device preferences."
    } catch {
        Write-Log "Loi khi mo App volume & device: $($_.Exception.Message)"
    }
})


# Handler: chạy System File Checker
$buttonSFC.Add_Click({
    try {
        Write-Log "Dang chay 'sfc /scannow' (yeu cau quyen Administrator)..."
        # Mở cmd nâng quyền và chạy sfc
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c sfc /scannow" -Verb RunAs -WindowStyle Normal
        Write-Log "Da goi SFC. Vui long doi tien trinh kiem tra/ sua chua hoan tat trong cua so Command Prompt."
    } catch {
        Write-Log "Loi khi chay SFC: $($_.Exception.Message)"
    }
})


# Handler: tải driver Brother về ổ D và (tuỳ chọn) chạy installer
$buttonBrother.Add_Click({
    try {
        # Link tải (trang hướng dẫn/tải của Brother - có thể redirect)
        $url      = "https://sasinvn-my.sharepoint.com/:u:/g/personal/nam_tran_sasin_vn/IQDaZilunXpwTp2BNRbY6GfRAYE1e3fNHrM74a_YXaaUO58?download=1"

        # Thư mục & tên file đích
        $destDir  = "D:\"
        $destFile = Join-Path $destDir "Brother_Printer_Driver.exe"
        $minSize  = 1024 * 100  # 100 KB - ngưỡng nhận diện file hợp lệ (tránh HTML)

        if (-not (Test-Path $destDir)) {
            New-Item -Path $destDir -ItemType Directory -Force | Out-Null
        }

        # Hàm phụ: cài đặt nếu file hợp lệ (hỗ trợ .exe/.msi; .zip sẽ giải nén)
        function Install-IfValid($path, $threshold) {
            if ((Test-Path $path) -and ((Get-Item $path).Length -ge $threshold)) {
                try { Unblock-File -Path $path } catch { }
                $ext = [System.IO.Path]::GetExtension($path).ToLower()
                if ($ext -eq ".zip") {
                    Write-Log "File tai ve la ZIP - dang giai nen..."
                    $extractDir = Join-Path (Split-Path $path -Parent) "BrotherDriver_Extract"
                    try { New-Item -Path $extractDir -ItemType Directory -Force | Out-Null } catch { }
                    try {
                        Expand-Archive -Path $path -DestinationPath $extractDir -Force
                        Write-Log "Da giai nen vao: $extractDir"
                        # tìm file setup .exe/.msi
                        $setup = Get-ChildItem -Path $extractDir -Recurse -Include *.exe,*.msi | Select-Object -First 1
                        if ($setup) {
                            Write-Log ("Dang chay cai dat: {0}" -f $setup.FullName)
                            if ($setup.Extension -ieq ".msi") {
                                Start-Process "msiexec.exe" -ArgumentList "/i `"$($setup.FullName)`" /qn" -Verb RunAs
                            } else {
                                Start-Process $setup.FullName -Verb RunAs
                            }
                            return $true
                        } else {
                            Write-Log "Khong tim thay file cai dat (.exe/.msi) trong goi ZIP."
                            return $false
                        }
                    } catch {
                        Write-Log "Loi khi giai nen ZIP: $($_.Exception.Message)"
                        return $false
                    }
                } else {
                    Write-Log ("Dang chay cai dat: {0}" -f $path)
                    if ($ext -eq ".msi") {
                        Start-Process "msiexec.exe" -ArgumentList "/i `"$path`" /qn" -Verb RunAs
                    } else {
                        Start-Process $path -Verb RunAs  # .exe
                    }
                    return $true
                }
            } else {
                Write-Log "File khong hop le hoac khong ton tai."
                return $false
            }
        }

        # 1) Nếu đã có file hợp lệ -> cài luôn, bỏ qua tải
        if (Install-IfValid -path $destFile -threshold $minSize) {
            return
        }

        # 2) Chưa có/không hợp lệ -> tiến hành tải
        Write-Log "File chua co/khong hop le. Dang tai Brother driver tu link online..."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        try {
            Invoke-WebRequest -Uri $url -OutFile $destFile -UseBasicParsing -ErrorAction Stop
            Write-Log ("Da tai driver ve: {0}" -f $destFile)

            # 3) Tải xong -> xác minh & cài
            if (-not (Install-IfValid -path $destFile -threshold $minSize)) {
                Write-Log "File tai ve khong hop le (co the la HTML hoac file qua nho)."
            }

        } catch {
            # 4) Tải lỗi (thường do cần đăng nhập/redirect) -> mở trình duyệt
            Write-Log "Khong tai duoc truc tiep (co the can dang nhap/redirect). Dang mo trinh duyet..."
            try {
                $chromePaths = @(
                    "$Env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                    "$Env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe"
                )
                $chromeExe = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
                if (-not $chromeExe) { $chromeExe = "chrome.exe" }

                Start-Process $chromeExe $url
                Write-Log "Da mo Chrome toi trang tai Brother. Vui long dang nhap/tai thu cong."
            } catch {
                Write-Log ("Loi khi mo Chrome: {0}. Thu mo bang trinh duyet mac dinh..." -f $_.Exception.Message)
                Start-Process $url
            }
        }

    } catch {
        Write-Log ("Loi khi xu ly Brother driver: {0}" -f $_.Exception.Message)
    }
})



# Handler: tải file USBZy303 từ link SharePoint về D:\ và (tuỳ chọn) chạy cài đặt
$buttonUSB80C.Add_Click({
    try {
        $url      = "https://sasinvn-my.sharepoint.com/:u:/g/personal/nam_tran_sasin_vn/IQB-e4-oDqmCQqEJDammgcCSAUSob7nF3HqU8PGYfFwq1IA?download=1"
        $destDir  = "D:\"
        $destFile = Join-Path $destDir "USBZy303_Setup.exe"
        $minSize  = 1024 * 100  # 100 KB, ngưỡng để nhận diện file "thực" (tránh HTML)

        if (-not (Test-Path $destDir)) {
            New-Item -Path $destDir -ItemType Directory -Force | Out-Null
        }

        # Hàm phụ: cài đặt nếu file hợp lệ
        function Install-IfValid($path, $threshold) {
            if ((Test-Path $path) -and ((Get-Item $path).Length -ge $threshold)) {
                try { Unblock-File -Path $path } catch { }
                Write-Log ("Dang chay cai dat: {0}" -f $path)
                Start-Process $path -Verb RunAs  # UAC prompt
                return $true
            } else {
                Write-Log "File khong hop le hoac khong ton tai."
                return $false
            }
        }

        # 1) Nếu đã có file hợp lệ -> cài luôn, bỏ qua tải
        if (Install-IfValid -path $destFile -threshold $minSize) {
            return
        }

        # 2) Chưa có/không hợp lệ -> tiến hành tải
        Write-Log "File chua co/khong hop le. Dang tai USBZy303 tu link online..."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        try {
            Invoke-WebRequest -Uri $url -OutFile $destFile -UseBasicParsing -ErrorAction Stop
            Write-Log ("Da tai USBZy303 ve: {0}" -f $destFile)

            # 3) Tải xong -> xác minh & cài
            if (-not (Install-IfValid -path $destFile -threshold $minSize)) {
                Write-Log "File tai ve khong hop le (co the la HTML hoac file qua nho)."
            }

        } catch {
            # 4) Tải lỗi (thường do cần đăng nhập SharePoint) -> mở trình duyệt
            Write-Log "Khong tai duoc truc tiep (co the can dang nhap SharePoint). Dang mo trinh duyet..."
            try {
                $chromePaths = @(
                    "$Env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                    "$Env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe"
                )
                $chromeExe = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
                if (-not $chromeExe) { $chromeExe = "chrome.exe" }
                Start-Process $chromeExe $url
                Write-Log "Da mo Chrome toi trang tai USBZy303. Vui long dang nhap va tai thu cong."
            } catch {
                Write-Log ("Loi khi mo Chrome: {0}. Thu mo bang trinh duyet mac dinh..." -f $_.Exception.Message)
                Start-Process $url
            }
        }

    } catch {
        Write-Log ("Loi khi xu ly USBZy303: {0}" -f $_.Exception.Message)
    }
})


# Handler: tải file ZY303 từ link SharePoint về D:\ và (tuỳ chọn) chạy cài đặt
$buttonZY303.Add_Click({
    try {
        $url      = "https://sasinvn-my.sharepoint.com/:u:/g/personal/nam_tran_sasin_vn/IQASdgPSrQPoRbV7YqU8fz55AVWSMES8f86SVr_LjZPn2r4?download=1"
        $destDir  = "D:\"
        $destFile = Join-Path $destDir "ZY303_Setup.exe"
        $minSize  = 1024 * 100  # 100 KB - ngưỡng để coi file hợp lệ (tránh HTML)

        if (-not (Test-Path $destDir)) {
            New-Item -Path $destDir -ItemType Directory -Force | Out-Null
        }

        # Hàm phụ: cài đặt nếu file hợp lệ
        function Install-IfValid($path, $threshold) {
            if ((Test-Path $path) -and ((Get-Item $path).Length -ge $threshold)) {
                try { Unblock-File -Path $path } catch { }
                Write-Log ("Dang chay cai dat: {0}" -f $path)
                Start-Process $path -Verb RunAs  # UAC prompt yêu cầu xác nhận
                return $true
            } else {
                Write-Log "File khong hop le hoac khong ton tai."
                return $false
            }
        }

        # 1) Nếu đã có file hợp lệ -> cài luôn, bỏ qua tải
        if (Install-IfValid -path $destFile -threshold $minSize) {
            return
        }

        # 2) Chưa có/không hợp lệ -> tiến hành tải
        Write-Log "File chua co/khong hop le. Dang tai ZY303 tu link online..."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        try {
            Invoke-WebRequest -Uri $url -OutFile $destFile -UseBasicParsing -ErrorAction Stop
            Write-Log ("Da tai ZY303 ve: {0}" -f $destFile)

            # 3) Tải xong -> xác minh & cài
            if (-not (Install-IfValid -path $destFile -threshold $minSize)) {
                Write-Log "File tai ve khong hop le (co the la HTML hoac file qua nho)."
            }

        } catch {
            # 4) Tải lỗi (thường do cần đăng nhập SharePoint) -> mở trình duyệt
            Write-Log "Khong tai duoc truc tiep (co the can dang nhap SharePoint). Dang mo trinh duyet..."
            try {
                $chromePaths = @(
                    "$Env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                    "$Env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe"
                )
                $chromeExe = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
                if (-not $chromeExe) { $chromeExe = "chrome.exe" }
                Start-Process $chromeExe $url
                Write-Log "Da mo Chrome toi trang tai ZY303. Vui long dang nhap va tai thu cong."
            } catch {
                Write-Log ("Loi khi mo Chrome: {0}. Thu mo bang trinh duyet mac dinh..." -f $_.Exception.Message)
                Start-Process $url
            }
        }

    } catch {
        Write-Log ("Loi khi xu ly ZY303: {0}" -f $_.Exception.Message)
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
