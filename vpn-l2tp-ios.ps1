
<# 
    VpnMobileconfigBuilder.ps1
    WinForms PowerShell tạo/nhập/xuất danh sách VPN L2TP/IPSec và lưu .mobileconfig cho iOS/macOS.

    Tính năng:
      - Add VPN
      - Delete (áp dụng cho các dòng đã tích OverridePrimary)
      - Import (JSON) — hỗ trợ 1 đối tượng hoặc mảng đối tượng
      - Export (JSON) — chỉ xuất các dòng đã tích
      - Import (Excel) — đọc sheet đầu, header linh hoạt (không phân biệt hoa/thường)
      - Export (Excel) — chỉ xuất các dòng đã tích
      - Save .mobileconfig — chỉ gồm các VPN đã tích
      - Mặc định password: Sasin@2023
      - Nút "con mắt" để xem/ẩn password ở ô nhập
      - Grid hiển thị mật khẩu dạng '•' và ẩn ký tự khi chỉnh sửa cell
      - Cột STT
      - Live search theo DisplayName (không phân biệt hoa/thường)
      - Checkbox “Chọn tất cả” ở header cột OverridePrimary (áp dụng theo bộ lọc đang hiển thị)

    Yêu cầu:
      - Microsoft Excel (COM) để Import/Export Excel.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- UI chạy trong STA runspace để đảm bảo hiển thị ---
$uiScript = {
    [System.Windows.Forms.Application]::EnableVisualStyles()

    function New-GuidString { [System.Guid]::NewGuid().ToString() }
    function Escape-Xml([string]$s) { if ($null -eq $s) { return "" } [System.Security.SecurityElement]::Escape($s) }
    function Sanitize-Identifier([string]$s) {
        if ([string]::IsNullOrWhiteSpace($s)) { return "item" }
        $x = $s.ToLower().Trim()
        $x = ($x -replace '[^a-z0-9\.\-]', '-')
        if ($x -match '^[\.\-]') { $x = "id-$x" }
        $x
    }

    function Build-VpnPayloadXml {
    param(
        [string]$baseIdentifier,
        [string]$displayName,
        [string]$remoteAddress,
        [string]$authName,
        [string]$authPassword,
        [string]$sharedSecret,
        [string]$localIdentifierType = "KeyID",
        [bool]$overridePrimary = $true
    )
        $payloadUUID = New-GuidString
        $sanName     = Sanitize-Identifier $displayName
        $payloadId   = "$baseIdentifier.$sanName.payload"
        $overrideInt = if ($overridePrimary) { 1 } else { 0 }

@"
    <dict>
      <key>PayloadType</key>
      <string>com.apple.vpn.managed</string>
      <key>PayloadVersion</key>
      <integer>1</integer>
      <key>PayloadIdentifier</key>
      <string>$(Escape-Xml $payloadId)</string>
      <key>PayloadUUID</key>
      <string>$(Escape-Xml $payloadUUID)</string>
      <key>PayloadDisplayName</key>
      <string>$(Escape-Xml $displayName)</string>

      <key>UserDefinedName</key>
      <string>$(Escape-Xml $displayName)</string>
      <key>VPNType</key>
      <string>L2TP</string>

      <key>IPSec</key>
      <dict>
        <key>AuthenticationMethod</key>
        <string>SharedSecret</string>
        <key>SharedSecret</key>
        <string>$(Escape-Xml $sharedSecret)</string>
        <key>LocalIdentifierType</key>
        <string>$(Escape-Xml $localIdentifierType)</string>
      </dict>

      <key>PPP</key>
      <dict>
        <key>AuthName</key>
        <string>$(Escape-Xml $authName)</string>
        <key>AuthPassword</key>
        <string>$(Escape-Xml $authPassword)</string>
        <key>CommRemoteAddress</key>
        <string>$(Escape-Xml $remoteAddress)</string>
        <key>OverridePrimary</key>
        <integer>$overrideInt</integer>
      </dict>

      <key>Proxies</key>
      <dict/>
    </dict>
"@
    }

    function New-MobileconfigXml {
    param(
        [string]$profileDisplayName,
        [string]$rootIdentifier,
        [string]$baseIdentifier,
        $vpnRows
    )
        if (-not $vpnRows -or $vpnRows.Count -eq 0) {
            throw "Danh sách VPN rỗng. Hãy tích chọn ít nhất 1 dòng."
        }

        $rootUUID = New-GuidString
        $payloadsXml = ""

        foreach ($row in $vpnRows) {
            $localId  = if ($row.LocalIdentifierType) { $row.LocalIdentifierType } else { "KeyID" }
            $override = [bool]$row.OverridePrimary

            $params = @{
                baseIdentifier      = $baseIdentifier
                displayName         = $row.DisplayName
                remoteAddress       = $row.RemoteAddress
                authName            = $row.AuthName
                authPassword        = $row.AuthPassword
                sharedSecret        = $row.SharedSecret
                localIdentifierType = $localId
                overridePrimary     = $override
            }
            $payloadsXml += (Build-VpnPayloadXml @params)
        }

@"
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
 "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>PayloadType</key>
  <string>Configuration</string>
  <key>PayloadVersion</key>
  <integer>1</integer>
  <key>PayloadIdentifier</key>
  <string>$(Escape-Xml (Sanitize-Identifier $rootIdentifier))</string>
  <key>PayloadUUID</key>
  <string>$(Escape-Xml $rootUUID)</string>
  <key>PayloadDisplayName</key>
  <string>$(Escape-Xml $profileDisplayName)</string>

  <key>PayloadContent</key>
  <array>
$payloadsXml
  </array>
</dict>
</plist>
"@
    }

    # ---------------- UI ----------------
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "VPN Mobileconfig Builder (L2TP/IPSec) - Nam"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(1080, 800)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Profile fields
    $lblProfileName = New-Object System.Windows.Forms.Label
    $lblProfileName.Text = "Profile Display Name:"
    $lblProfileName.Location = New-Object System.Drawing.Point(20, 20)
    $lblProfileName.AutoSize = $true

    $tbProfileName = New-Object System.Windows.Forms.TextBox
    $tbProfileName.Location = New-Object System.Drawing.Point(180, 16)
    $tbProfileName.Width = 300
    $tbProfileName.Text = "VPN Profile"

    $lblRootId = New-Object System.Windows.Forms.Label
    $lblRootId.Text = "Root PayloadIdentifier:"
    $lblRootId.Location = New-Object System.Drawing.Point(20, 50)
    $lblRootId.AutoSize = $true

    $tbRootId = New-Object System.Windows.Forms.TextBox
    $tbRootId.Location = New-Object System.Drawing.Point(180, 46)
    $tbRootId.Width = 300
    $tbRootId.Text = "com.tranvannam.profile.vpn"

    $lblBaseId = New-Object System.Windows.Forms.Label
    $lblBaseId.Text = "Base PayloadIdentifier:"
    $lblBaseId.Location = New-Object System.Drawing.Point(20, 80)
    $lblBaseId.AutoSize = $true

    $tbBaseId = New-Object System.Windows.Forms.TextBox
    $tbBaseId.Location = New-Object System.Drawing.Point(180, 76)
    $tbBaseId.Width = 300
    $tbBaseId.Text = "com.tranvannam.vpn"

    # Group Add VPN
    $grpAdd = New-Object System.Windows.Forms.GroupBox
    $grpAdd.Text = "Add VPN"
    $grpAdd.Location = New-Object System.Drawing.Point(20, 110)
    $grpAdd.Size = New-Object System.Drawing.Size(1030, 160)

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = "Display Name:"
    $lblName.Location = New-Object System.Drawing.Point(15, 30)
    $lblName.AutoSize = $true

    $tbName = New-Object System.Windows.Forms.TextBox
    $tbName.Location = New-Object System.Drawing.Point(110, 26)
    $tbName.Width = 240
    $tbName.Text = "VPN 118HHT"

    $lblRemote = New-Object System.Windows.Forms.Label
    $lblRemote.Text = "Remote Address:"
    $lblRemote.Location = New-Object System.Drawing.Point(370, 30)
    $lblRemote.AutoSize = $true

    $tbRemote = New-Object System.Windows.Forms.TextBox
    $tbRemote.Location = New-Object System.Drawing.Point(480, 26)
    $tbRemote.Width = 260
    $tbRemote.Text = "hcs08ah163k.sn.mynetname.net"

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = "AuthName (user):"
    $lblUser.Location = New-Object System.Drawing.Point(15, 60)
    $lblUser.AutoSize = $true

    $tbUser = New-Object System.Windows.Forms.TextBox
    $tbUser.Location = New-Object System.Drawing.Point(130, 56)
    $tbUser.Width = 220
    $tbUser.Text = "nam.tran"

    $lblPass = New-Object System.Windows.Forms.Label
    $lblPass.Text = "AuthPassword:"
    $lblPass.Location = New-Object System.Drawing.Point(370, 60)
    $lblPass.AutoSize = $true

    $tbPass = New-Object System.Windows.Forms.TextBox
    $tbPass.Location = New-Object System.Drawing.Point(480, 56)
    $tbPass.Width = 220
    $tbPass.UseSystemPasswordChar = $true
    $tbPass.Text = "Sasin@2023"

    # Nút con mắt cho AuthPassword
    $btnEyePass = New-Object System.Windows.Forms.Button
    $btnEyePass.Text = "👁"
    $btnEyePass.Location = New-Object System.Drawing.Point(705, 56)
    $btnEyePass.Width = 35
    $btnEyePass.Height = $tbPass.Height
    $btnEyePass.FlatStyle = 'System'

    $lblSecret = New-Object System.Windows.Forms.Label
    $lblSecret.Text = "SharedSecret:"
    $lblSecret.Location = New-Object System.Drawing.Point(15, 90)
    $lblSecret.AutoSize = $true

    $tbSecret = New-Object System.Windows.Forms.TextBox
    $tbSecret.Location = New-Object System.Drawing.Point(110, 86)
    $tbSecret.Width = 220
    $tbSecret.UseSystemPasswordChar = $true
    $tbSecret.Text = "Sasin@2023"

    # Nút con mắt cho SharedSecret
    $btnEyeSecret = New-Object System.Windows.Forms.Button
    $btnEyeSecret.Text = "👁"
    $btnEyeSecret.Location = New-Object System.Drawing.Point(335, 86)
    $btnEyeSecret.Width = 35
    $btnEyeSecret.Height = $tbSecret.Height
    $btnEyeSecret.FlatStyle = 'System'

    $lblLocalIdType = New-Object System.Windows.Forms.Label
    $lblLocalIdType.Text = "LocalIdentifierType:"
    $lblLocalIdType.Location = New-Object System.Drawing.Point(370, 90)
    $lblLocalIdType.AutoSize = $true

    $cbLocalIdType = New-Object System.Windows.Forms.ComboBox
    $cbLocalIdType.Location = New-Object System.Drawing.Point(480, 86)
    $cbLocalIdType.Width = 120
    $cbLocalIdType.DropDownStyle = "DropDownList"
    [void]$cbLocalIdType.Items.Add("KeyID")
    $cbLocalIdType.SelectedIndex = 0

    $chkOverride = New-Object System.Windows.Forms.CheckBox
    $chkOverride.Text = "OverridePrimary"
    $chkOverride.Location = New-Object System.Drawing.Point(620, 88)
    $chkOverride.Checked = $true
    $chkOverride.AutoSize = $true

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "Add VPN"
    $btnAdd.Location = New-Object System.Drawing.Point(800, 82)
    $btnAdd.Width = 110

    $grpAdd.Controls.AddRange(@(
        $lblName, $tbName, $lblRemote, $tbRemote,
        $lblUser, $tbUser, $lblPass, $tbPass, $btnEyePass,
        $lblSecret, $tbSecret, $btnEyeSecret, $lblLocalIdType, $cbLocalIdType,
        $chkOverride, $btnAdd
    ))

    # --- Ô tìm kiếm DisplayName + label đếm ---
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Search DisplayName:"
    $lblSearch.Location = New-Object System.Drawing.Point(20, 280)
    $lblSearch.AutoSize = $true

    $tbSearch = New-Object System.Windows.Forms.TextBox
    $tbSearch.Location = New-Object System.Drawing.Point(140, 276)
    $tbSearch.Width = 260

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "Result: 0 / Total: 0 | Selected: 0"
    $lblCount.Location = New-Object System.Drawing.Point(420, 280)
    $lblCount.AutoSize = $true

    # Grid + DataTable
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(20, 310)
    $grid.Size = New-Object System.Drawing.Size(1030, 320)
    $grid.AllowUserToAddRows = $false
    $grid.SelectionMode = "FullRowSelect"
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.ReadOnly = $false
    $grid.EditMode = 'EditOnEnter'

    $dt = New-Object System.Data.DataTable
    $dt.CaseSensitive = $false
    [void]$dt.Columns.Add("DisplayName", [string])
    [void]$dt.Columns.Add("RemoteAddress", [string])
    [void]$dt.Columns.Add("AuthName", [string])
    [void]$dt.Columns.Add("AuthPassword", [string])
    [void]$dt.Columns.Add("SharedSecret", [string])
    [void]$dt.Columns.Add("LocalIdentifierType", [string])
    [void]$dt.Columns.Add("OverridePrimary", [bool])
    $grid.DataSource = $dt

    # --- Cột STT ---
    $colStt = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStt.Name = "STT"
    $colStt.HeaderText = "STT"
    $colStt.ReadOnly = $true
    $colStt.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::DisplayedCells
    $grid.Columns.Add($colStt)
    $grid.Columns["STT"].DisplayIndex = 0

    # --- Header checkbox chọn tất cả ---
    $chkSelectAll = New-Object System.Windows.Forms.CheckBox
    $chkSelectAll.Size = New-Object System.Drawing.Size(15, 15)
    $chkSelectAll.Text = ""
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.SetToolTip($chkSelectAll, "Check/unCheck all line show")
    $script:suppressHeaderEvent = $false

    function Place-HeaderCheckbox {
        $idx = $grid.Columns["OverridePrimary"].Index
        $rect = $grid.GetCellDisplayRectangle($idx, -1, $true)
        $chkSelectAll.Location = New-Object System.Drawing.Point($rect.X + [math]::Max(8, $rect.Width/2 - 8), $rect.Y + 4)
        if (-not $grid.Controls.Contains($chkSelectAll)) { $grid.Controls.Add($chkSelectAll) }
    }
    Place-HeaderCheckbox
    $grid.Add_Scroll({ Place-HeaderCheckbox })
    $grid.Add_ColumnWidthChanged({ Place-HeaderCheckbox })
    $grid.Add_SizeChanged({ Place-HeaderCheckbox })

    $chkSelectAll.Add_CheckedChanged({
        if ($script:suppressHeaderEvent) { return }
        $state = $chkSelectAll.Checked
        foreach ($drv in $dt.DefaultView) { $drv["OverridePrimary"] = $state }
        Update-Stt
    })

    # --- Mask mật khẩu khi hiển thị ---
    $grid.Add_CellFormatting({
        param($sender, $e)
        if ($e.ColumnIndex -ge 0 -and $e.RowIndex -ge 0) {
            $colName = $sender.Columns[$e.ColumnIndex].Name
            if ($colName -eq "AuthPassword" -or $colName -eq "SharedSecret") {
                $val = $e.Value
                if ($null -ne $val -and $val -ne "") {
                    $len = [math]::Min(64, $val.ToString().Length)
                    $e.Value = ("•" * $len)
                } else {
                    $e.Value = ""
                }
                $e.FormattingApplied = $true
            }
        }
    })

    # --- Ẩn ký tự khi chỉnh sửa cell password ---
    $grid.Add_EditingControlShowing({
        param($sender, $e)
        $colIndex = $sender.CurrentCell.ColumnIndex
        $colName = $sender.Columns[$colIndex].Name
        if ($colName -eq "AuthPassword" -or $colName -eq "SharedSecret") {
            $tb = $e.Control -as [System.Windows.Forms.TextBox]
            if ($tb) { $tb.UseSystemPasswordChar = $true }
        }
    })

    # Commit checkbox ngay khi click
    $grid.Add_CurrentCellDirtyStateChanged({
        if ($grid.IsCurrentCellDirty) {
            $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    function Get-CheckedRows {
        $checked = @()
        foreach ($drv in $dt.DefaultView) { if ([bool]$drv["OverridePrimary"]) { $checked += $drv } }
        return $checked
    }

    function Update-Stt {
        $i = 1
        foreach ($row in $grid.Rows) {
            if (-not $row.IsNewRow) { $row.Cells["STT"].Value = $i; $i++ }
        }
        $checkedCount = (Get-CheckedRows).Count
        $lblCount.Text = "Result: $($grid.Rows.Count) / Total: $($dt.Rows.Count) | Selected: $checkedCount"

        $allVisChecked = $true; $hasVisRow = $false
        foreach ($row in $grid.Rows) {
            if (-not $row.IsNewRow) {
                $hasVisRow = $true
                if (-not [bool]$row.Cells["OverridePrimary"].Value) { $allVisChecked = $false }
            }
        }
        $script:suppressHeaderEvent = $true
        $chkSelectAll.Checked = ($hasVisRow -and $allVisChecked)
        $script:suppressHeaderEvent = $false
    }

    # Live search theo DisplayName
    $tbSearch.Add_TextChanged({
        $text = $tbSearch.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { $dt.DefaultView.RowFilter = "" }
        else {
            $t = $text.Replace("'", "''")
            $dt.DefaultView.RowFilter = "DisplayName LIKE '%$t%'"
        }
        Update-Stt
    })

    $grid.Add_Sorted({ Update-Stt })
    $grid.Add_DataBindingComplete({ Update-Stt })
    $grid.Add_CellValueChanged({ Update-Stt })

    # Buttons (bổ sung Import/Export Excel)
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "Delete Selected (selected)"
    $btnDelete.Location = New-Object System.Drawing.Point(20, 650)
    $btnDelete.Width = 180

    $btnImportJson = New-Object System.Windows.Forms.Button
    $btnImportJson.Text = "Import (JSON)"
    $btnImportJson.Location = New-Object System.Drawing.Point(220, 650)
    $btnImportJson.Width = 150

    $btnImportExcel = New-Object System.Windows.Forms.Button
    $btnImportExcel.Text = "Import (Excel)"
    $btnImportExcel.Location = New-Object System.Drawing.Point(390, 650)
    $btnImportExcel.Width = 150

    $btnExportJson = New-Object System.Windows.Forms.Button
    $btnExportJson.Text = "Export (JSON) (selected)"
    $btnExportJson.Location = New-Object System.Drawing.Point(560, 650)
    $btnExportJson.Width = 190

    $btnExportExcel = New-Object System.Windows.Forms.Button
    $btnExportExcel.Text = "Export (Excel) (selected)"
    $btnExportExcel.Location = New-Object System.Drawing.Point(770, 650)
    $btnExportExcel.Width = 190

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save .mobileconfig (selected)"
    $btnSave.Location = New-Object System.Drawing.Point(20, 690)
    $btnSave.Width = 240

    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Text = "Exit"
    $btnExit.Location = New-Object System.Drawing.Point(280, 690)
    $btnExit.Width = 120

    # Toggle mắt
    $btnEyePass.Add_Click({ $tbPass.UseSystemPasswordChar = -not $tbPass.UseSystemPasswordChar; $btnEyePass.Text = if ($tbPass.UseSystemPasswordChar) { "👁" } else { "🙈" } })
    $btnEyeSecret.Add_Click({ $tbSecret.UseSystemPasswordChar = -not $tbSecret.UseSystemPasswordChar; $btnEyeSecret.Text = if ($tbSecret.UseSystemPasswordChar) { "👁" } else { "🙈" } })

    # Add VPN
    $btnAdd.Add_Click({
        $d = $tbName.Text.Trim()
        $r = $tbRemote.Text.Trim()
        $u = $tbUser.Text.Trim()
        $p = $tbPass.Text
        $s = $tbSecret.Text
        $l = $cbLocalIdType.SelectedItem
        $o = $chkOverride.Checked

        if ([string]::IsNullOrWhiteSpace($d) -or [string]::IsNullOrWhiteSpace($r) -or
            [string]::IsNullOrWhiteSpace($u) -or [string]::IsNullOrWhiteSpace($p) -or
            [string]::IsNullOrWhiteSpace($s)) {
            [System.Windows.Forms.MessageBox]::Show("Please input All DisplayName, RemoteAddress, AuthName, AuthPassword, SharedSecret.",
                "Thiếu thông tin", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $row = $dt.NewRow()
        $row["DisplayName"]         = $d
        $row["RemoteAddress"]       = $r
        $row["AuthName"]            = $u
        $row["AuthPassword"]        = $p
        $row["SharedSecret"]        = $s
        $row["LocalIdentifierType"] = $l
        $row["OverridePrimary"]     = $o
        $dt.Rows.Add($row)

        $tbPass.Text   = "Sasin@2023"
        $tbSecret.Text = "Sasin@2023"
        $tbPass.UseSystemPasswordChar   = $true
        $tbSecret.UseSystemPasswordChar = $true
        $btnEyePass.Text   = "👁"
        $btnEyeSecret.Text = "👁"

        Update-Stt
    })

    # Delete (đã tích)
    $btnDelete.Add_Click({
        $checked = Get-CheckedRows
        if ($checked.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Not Have Select OverridePrimary.", "Notification",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        foreach ($drv in $checked) { $drv.Row.Delete() }
        Update-Stt
    })

    # Export JSON (đã tích)
    $btnExportJson.Add_Click({
        $checked = Get-CheckedRows
        if ($checked.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Not Have Select export.", "Notification",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $sfd.Title = "Save list VPN (JSON) - only line selected"
        $sfd.FileName = "vpn-selected.json"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $list = @()
                foreach ($drv in $checked) {
                    $r = $drv.Row
                    $list += [pscustomobject]@{
                        DisplayName         = "" + $r["DisplayName"]
                        RemoteAddress       = "" + $r["RemoteAddress"]
                        AuthName            = "" + $r["AuthName"]
                        AuthPassword        = "" + $r["AuthPassword"]
                        SharedSecret        = "" + $r["SharedSecret"]
                        LocalIdentifierType = if ($r["LocalIdentifierType"]) { "" + $r["LocalIdentifierType"] } else { "KeyID" }
                        OverridePrimary     = [bool]$r["OverridePrimary"]
                    }
                }
                ($list | ConvertTo-Json -Depth 5) | Out-File -LiteralPath $sfd.FileName -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Saved JSON: $($sfd.FileName)", "OK",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error export: $($_.Exception.Message)", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            }
        }
    })

    # Save .mobileconfig (đã tích)
    $btnSave.Add_Click({
        try {
            $profileName = $tbProfileName.Text.Trim()
            $rootId      = $tbRootId.Text.Trim()
            $baseId      = $tbBaseId.Text.Trim()

            if ([string]::IsNullOrWhiteSpace($profileName) -or [string]::IsNullOrWhiteSpace($rootId) -or [string]::IsNullOrWhiteSpace($baseId)) {
                throw "Profile Display Name, Root/Base PayloadIdentifier have not Null."
            }

            $checked = Get-CheckedRows
            if ($checked.Count -eq 0) {
                throw "Not Have selected OverridePrimary. Please select line you want in file."
            }

            $rows = @()
            foreach ($drv in $checked) {
                $r = $drv.Row
                $rows += @{
                    DisplayName         = "" + $r["DisplayName"]
                    RemoteAddress       = "" + $r["RemoteAddress"]
                    AuthName            = "" + $r["AuthName"]
                    AuthPassword        = "" + $r["AuthPassword"]
                    SharedSecret        = "" + $r["SharedSecret"]
                    LocalIdentifierType = if ($r["LocalIdentifierType"]) { "" + $r["LocalIdentifierType"] } else { "KeyID" }
                    OverridePrimary     = [bool]$r["OverridePrimary"]
                }
            }

            $xml = New-MobileconfigXml -profileDisplayName $profileName -rootIdentifier $rootId -baseIdentifier $baseId -vpnRows $rows

            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = "Mobileconfig (*.mobileconfig)|*.mobileconfig|All files (*.*)|*.*"
            $sfd.Title = "Save file .mobileconfig (only line selected)"
            $sfd.FileName = (Sanitize-Identifier $profileName) + ".mobileconfig"
            if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                [System.IO.File]::WriteAllText($sfd.FileName, $xml, [System.Text.Encoding]::UTF8)
                [System.Windows.Forms.MessageBox]::Show(
                    "Created file: $($sfd.FileName)`nLetgo AirDrop/email file here in iPhone/iPad to setup.",
                    "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error when create .mobileconfig: $($_.Exception.Message)", "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    # Import JSON (hỗ trợ 1 đối tượng hoặc mảng)
    $btnImportJson.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $ofd.Title = "Select file JSON to import"
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $json = Get-Content -LiteralPath $ofd.FileName -Raw -ErrorAction Stop
                $items = $json | ConvertFrom-Json
                if (-not ($items -is [System.Collections.IEnumerable])) { $items = @($items) }
                if (-not $items -or $items.Count -eq 0) { throw "File JSON not have data." }

                $dt.Rows.Clear()
                foreach ($it in $items) {
                    $row = $dt.NewRow()
                    $row["DisplayName"]         = "" + $it.DisplayName
                    $row["RemoteAddress"]       = "" + $it.RemoteAddress
                    $row["AuthName"]            = "" + $it.AuthName
                    $row["AuthPassword"]        = "" + $it.AuthPassword
                    $row["SharedSecret"]        = "" + $it.SharedSecret
                    $row["LocalIdentifierType"] = if ($it.LocalIdentifierType) { "" + $it.LocalIdentifierType } else { "KeyID" }
                    $row["OverridePrimary"]     = [bool]$it.OverridePrimary
                    $dt.Rows.Add($row)
                }
                $dt.DefaultView.RowFilter = ""; $tbSearch.Text = ""; Update-Stt
                [System.Windows.Forms.MessageBox]::Show("Import success: $($items.Count) List.", "OK",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error import: $($_.Exception.Message)", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            }
        }
    })

    # ====== Import Excel (COM) ======
    $btnImportExcel.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|All files (*.*)|*.*"
        $ofd.Title = "Select file Excel to import"
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $excel = New-Object -ComObject Excel.Application
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Can't initia Excel COM. For sure Microsoft Excel setup done.", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                return
            }
            try {
                $excel.Visible = $false
                $wb = $excel.Workbooks.Open($ofd.FileName)
                $ws = $wb.Worksheets.Item(1)
                $used = $ws.UsedRange
                $rowCount = $used.Rows.Count
                $colCount = $used.Columns.Count
                if ($rowCount -lt 2) { throw "Header has no data (requires at least 1 row header + 1 row data)." }

                # Đọc header (row 1)
                $headers = @()
                for ($c = 1; $c -le $colCount; $c++) { $headers += ("" + $ws.Cells.Item(1, $c).Text).Trim() }

                # Map cột (không phân biệt hoa/thường)
                function Find-ColIndex([string]$name) {
                    for ($i=0; $i -lt $headers.Count; $i++) {
                        if ($headers[$i] -and ($headers[$i] -eq $name -or $headers[$i].ToLower() -eq $name.ToLower())) { return ($i+1) }
                    }
                    return $null
                }
                $idxDisplay = Find-ColIndex "DisplayName"
                $idxRemote  = Find-ColIndex "RemoteAddress"
                $idxAuth    = Find-ColIndex "AuthName"
                $idxPass    = Find-ColIndex "AuthPassword"
                $idxSecret  = Find-ColIndex "SharedSecret"
                $idxLocalId = Find-ColIndex "LocalIdentifierType"
                $idxOverride= Find-ColIndex "OverridePrimary"

                # Kiểm tra tối thiểu
                if (-not $idxDisplay -or -not $idxRemote -or -not $idxAuth) {
                    throw "Missing required columns: DisplayName, RemoteAddress, AuthName."
                }

                $dt.Rows.Clear()
                for ($r = 2; $r -le $rowCount; $r++) {
                    $dn = if ($idxDisplay) { ("" + $ws.Cells.Item($r, $idxDisplay).Text).Trim() } else { "" }
                    $ra = if ($idxRemote)  { ("" + $ws.Cells.Item($r, $idxRemote).Text).Trim() } else { "" }
                    $an = if ($idxAuth)    { ("" + $ws.Cells.Item($r, $idxAuth).Text).Trim() } else { "" }
                    $ap = if ($idxPass)    { ("" + $ws.Cells.Item($r, $idxPass).Text).Trim() } else { "" }
                    $ss = if ($idxSecret)  { ("" + $ws.Cells.Item($r, $idxSecret).Text).Trim() } else { "" }
                    $li = if ($idxLocalId) { ("" + $ws.Cells.Item($r, $idxLocalId).Text).Trim() } else { "KeyID" }
                    $ovRaw = if ($idxOverride) { ("" + $ws.Cells.Item($r, $idxOverride).Text).Trim() } else { "" }

                    if ([string]::IsNullOrWhiteSpace($dn) -and [string]::IsNullOrWhiteSpace($ra) -and [string]::IsNullOrWhiteSpace($an)) { continue }

                    # parse bool OverridePrimary
                    $ov =
                        ($ovRaw -eq $true) -or
                        ($ovRaw -as [string] -match '^(true|1|yes|y)$')

                    $row = $dt.NewRow()
                    $row["DisplayName"]         = $dn
                    $row["RemoteAddress"]       = $ra
                    $row["AuthName"]            = $an
                    $row["AuthPassword"]        = $ap
                    $row["SharedSecret"]        = $ss
                    $row["LocalIdentifierType"] = if ($li) { $li } else { "KeyID" }
                    $row["OverridePrimary"]     = [bool]$ov
                    $dt.Rows.Add($row)
                }

                $wb.Close($false); $excel.Quit()
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($used)
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
                [GC]::Collect(); [GC]::WaitForPendingFinalizers()

                $dt.DefaultView.RowFilter = ""; $tbSearch.Text = ""; Update-Stt
                [System.Windows.Forms.MessageBox]::Show("Import Excel Success.", "OK",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

            } catch {
                try { if ($wb) { $wb.Close($false) } if ($excel) { $excel.Quit() } } catch {}
                [System.Windows.Forms.MessageBox]::Show("Error import Excel: $($_.Exception.Message)", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            }
        }
    })

    # ====== Export Excel (COM) — chỉ xuất dòng đã tích ======
    $btnExportExcel.Add_Click({
        $checked = Get-CheckedRows
        if ($checked.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No lines have been checked to export to Excel.", "Notification",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls"
        $sfd.Title = "Save Excel - only selected rows"
        $sfd.FileName = "vpn-selected.xlsx"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $excel = New-Object -ComObject Excel.Application
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Could not initialize Excel COM. Make sure Microsoft Excel is installed.", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                return
            }
            try {
                $excel.Visible = $false
                $wb = $excel.Workbooks.Add()
                $ws = $wb.Worksheets.Item(1)
                $ws.Name = "VPNs"

                $headers = @("STT","DisplayName","RemoteAddress","AuthName","AuthPassword","SharedSecret","LocalIdentifierType","OverridePrimary")
                for ($c=0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(1, $c+1) = $headers[$c] }

                $rowIdx = 2; $stt = 1
                foreach ($drv in $checked) {
                    $r = $drv.Row
                    $ws.Cells.Item($rowIdx, 1) = $stt
                    $ws.Cells.Item($rowIdx, 2) = "" + $r["DisplayName"]
                    $ws.Cells.Item($rowIdx, 3) = "" + $r["RemoteAddress"]
                    $ws.Cells.Item($rowIdx, 4) = "" + $r["AuthName"]
                    $ws.Cells.Item($rowIdx, 5) = "" + $r["AuthPassword"]  # ⚠ Xuất giá trị thật (bảo mật file!)
                    $ws.Cells.Item($rowIdx, 6) = "" + $r["SharedSecret"]   # ⚠ Xuất giá trị thật
                    $ws.Cells.Item($rowIdx, 7) = "" + $r["LocalIdentifierType"]
                    $ws.Cells.Item($rowIdx, 8) = [bool]$r["OverridePrimary"]
                    $rowIdx++; $stt++
                }

                $ws.Columns.AutoFit() | Out-Null

                $ext = [System.IO.Path]::GetExtension($sfd.FileName).ToLower()
                if ($ext -eq ".xlsx") {
                    $xlOpenXMLWorkbook = 51
                    $wb.SaveAs($sfd.FileName, $xlOpenXMLWorkbook)
                } elseif ($ext -eq ".xls") {
                    $xlWorkbookNormal = -4143
                    $wb.SaveAs($sfd.FileName, $xlWorkbookNormal)
                } else {
                    # mặc định .xlsx
                    $xlOpenXMLWorkbook = 51
                    $wb.SaveAs($sfd.FileName, $xlOpenXMLWorkbook)
                }

                $wb.Close($true); $excel.Quit()
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
                [GC]::Collect(); [GC]::WaitForPendingFinalizers()

                [System.Windows.Forms.MessageBox]::Show("Exported Excel: $($sfd.FileName)", "OK",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

            } catch {
                try { if ($wb) { $wb.Close($false) } if ($excel) { $excel.Quit() } } catch {}
                [System.Windows.Forms.MessageBox]::Show("Error export Excel: $($_.Exception.Message)", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            }
        }
    })

    $btnExit.Add_Click({ $form.Close() })

    $form.Controls.AddRange(@(
        $lblProfileName, $tbProfileName, $lblRootId, $tbRootId, $lblBaseId, $tbBaseId,
        $grpAdd, $lblSearch, $tbSearch, $lblCount, $grid,
        $btnDelete, $btnImportJson, $btnImportExcel, $btnExportJson, $btnExportExcel, $btnSave, $btnExit
    ))

    Update-Stt
    [void]$form.ShowDialog()
}

$rs = [runspacefactory]::CreateRunspace()
$rs.ApartmentState = 'STA'
$rs.ThreadOptions = 'ReuseThread'
$rs.Open()

$ps = [powershell]::Create()
$ps.Runspace = $rs
$null = $ps.AddScript($uiScript)
$null = $ps.Invoke()

$ps.Dispose()
$rs.Close()
$rs.Dispose()
