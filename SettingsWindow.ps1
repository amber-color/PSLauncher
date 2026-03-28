#
# SettingsWindow.ps1 - 設定 GUI（アプリ管理 + 各種設定）
#

# ---------------------------------------------------------------------------
# アプリ追加 / 編集ダイアログ
# ---------------------------------------------------------------------------
function Show-AppEditDialog {
    param(
        $App = $null
    )

    $isNew  = ($null -eq $App)
    $result = $null

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = if ($isNew) { 'アプリ追加' } else { 'アプリ編集' }
    $dlg.Size            = New-Object System.Drawing.Size(540, 340)
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.BackColor       = [System.Drawing.Color]::FromArgb(37, 37, 38)
    $dlg.ForeColor       = [System.Drawing.Color]::White

    function New-DlgLabel { param([string]$Text, [int]$Y)
        $l = New-Object System.Windows.Forms.Label
        $l.Text      = $Text
        $l.Location  = New-Object System.Drawing.Point(12, $Y)
        $l.Size      = New-Object System.Drawing.Size(80, 22)
        $l.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
        return $l
    }

    function New-DlgTextBox { param([int]$Y, [int]$Width = 370)
        $t = New-Object System.Windows.Forms.TextBox
        $t.Location    = New-Object System.Drawing.Point(95, $Y)
        $t.Width       = $Width
        $t.BackColor   = [System.Drawing.Color]::FromArgb(58, 58, 58)
        $t.ForeColor   = [System.Drawing.Color]::White
        $t.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        return $t
    }

    function New-DlgButton { param([string]$Text, [int]$X, [int]$Y, [int]$W = 44)
        $b = New-Object System.Windows.Forms.Button
        $b.Text      = $Text
        $b.Location  = New-Object System.Drawing.Point($X, $Y)
        $b.Size      = New-Object System.Drawing.Size($W, 24)
        $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $b.BackColor = [System.Drawing.Color]::FromArgb(58, 58, 58)
        $b.ForeColor = [System.Drawing.Color]::White
        $b.FlatAppearance.BorderSize = 0
        return $b
    }

    # 表示名
    $lblName = New-DlgLabel '表示名' 16
    $txtName = New-DlgTextBox 14 370
    $txtName.Text = if ($isNew) { '' } else { $App.name }

    # パス / URL
    $lblPath = New-DlgLabel 'パス / URL' 52
    $txtPath = New-DlgTextBox 50 250
    $txtPath.Text = if ($isNew) { '' } else { $App.path }

    $btnBrowseFile   = New-DlgButton '...'      350 50 44
    $btnBrowseFolder = New-DlgButton 'フォルダ' 398 50 68

    $btnBrowseFile.add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title  = '実行ファイルを選択'
        $ofd.Filter = '実行ファイル (*.exe;*.cmd;*.bat)|*.exe;*.cmd;*.bat|すべてのファイル (*.*)|*.*'
        if ($ofd.ShowDialog($dlg) -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtPath.Text = $ofd.FileName
            if ($txtName.Text -eq '') {
                $txtName.Text = [System.IO.Path]::GetFileNameWithoutExtension($ofd.FileName)
            }
        }
    }.GetNewClosure())

    $btnBrowseFolder.add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = 'フォルダを選択'
        if ($fbd.ShowDialog($dlg) -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtPath.Text = $fbd.SelectedPath
            if ($txtName.Text -eq '') {
                $txtName.Text = [System.IO.Path]::GetFileName($fbd.SelectedPath)
            }
        }
    }.GetNewClosure())

    # アイコンパス
    $lblIcon = New-DlgLabel 'アイコン' 88
    $txtIcon = New-DlgTextBox 86 250
    $txtIcon.Text = if ($isNew) { '' } else { $App.iconPath }

    $btnBrowseIcon = New-DlgButton '...' 350 86 44
    $btnBrowseIcon.add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title  = 'アイコンファイルを選択'
        $ofd.Filter = '画像 / アイコン (*.ico;*.png;*.bmp)|*.ico;*.png;*.bmp|すべてのファイル (*.*)|*.*'
        if ($ofd.ShowDialog($dlg) -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtIcon.Text = $ofd.FileName
        }
    }.GetNewClosure())

    # グループ
    $lblGroup = New-DlgLabel 'グループ' 124
    $txtGroup = New-DlgTextBox 122 370
    $txtGroup.Text = if ($isNew) { '' } else { $App.group }

    # ヒント
    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text      = '※ パスには実行ファイル・フォルダ・URL を指定できます。アイコンは省略可。'
    $lblHint.Location  = New-Object System.Drawing.Point(12, 158)
    $lblHint.Size      = New-Object System.Drawing.Size(500, 18)
    $lblHint.ForeColor = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $lblHint.Font      = New-Object System.Drawing.Font('Segoe UI', 8)

    # OK / キャンセル
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text      = 'OK'
    $btnOk.Size      = New-Object System.Drawing.Size(80, 30)
    $btnOk.Location  = New-Object System.Drawing.Point(320, 262)
    $btnOk.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOk.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.FlatAppearance.BorderSize = 0
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text      = 'キャンセル'
    $btnCancel.Size      = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Location  = New-Object System.Drawing.Point(410, 262)
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(58, 58, 58)
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatAppearance.BorderSize = 0
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    $dlg.Controls.AddRange(@(
        $lblName, $txtName,
        $lblPath, $txtPath, $btnBrowseFile, $btnBrowseFolder,
        $lblIcon, $txtIcon, $btnBrowseIcon,
        $lblGroup, $txtGroup,
        $lblHint,
        $btnOk, $btnCancel
    ))

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($txtName.Text.Trim() -eq '' -or $txtPath.Text.Trim() -eq '') {
            [System.Windows.Forms.MessageBox]::Show(
                '表示名とパスは必須です。',
                'PSLauncher',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        } else {
            $result = [ordered]@{
                name     = $txtName.Text.Trim()
                path     = $txtPath.Text.Trim()
                iconPath = $txtIcon.Text.Trim()
                group    = $txtGroup.Text.Trim()
            }
        }
    }

    $dlg.Dispose()
    return $result
}

# ---------------------------------------------------------------------------
# 設定ウィンドウ生成
# ---------------------------------------------------------------------------
function New-SettingsWindow {

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'PSLauncher 設定'
    $form.Size            = New-Object System.Drawing.Size(720, 560)
    $form.MinimumSize     = New-Object System.Drawing.Size(600, 480)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor       = [System.Drawing.Color]::FromArgb(37, 37, 38)
    $form.ForeColor       = [System.Drawing.Color]::White

    # ------------------------------------------------------------------
    # TabControl
    # ------------------------------------------------------------------
    $tab = New-Object System.Windows.Forms.TabControl
    $tab.Dock       = [System.Windows.Forms.DockStyle]::Fill
    $tab.Appearance = [System.Windows.Forms.TabAppearance]::Normal
    $tab.BackColor  = [System.Drawing.Color]::FromArgb(37, 37, 38)

    $pageApps = New-Object System.Windows.Forms.TabPage
    $pageApps.Text      = 'アプリ'
    $pageApps.BackColor = [System.Drawing.Color]::FromArgb(37, 37, 38)
    $pageApps.ForeColor = [System.Drawing.Color]::White

    $pageSettings = New-Object System.Windows.Forms.TabPage
    $pageSettings.Text      = '設定'
    $pageSettings.BackColor = [System.Drawing.Color]::FromArgb(37, 37, 38)
    $pageSettings.ForeColor = [System.Drawing.Color]::White

    $tab.TabPages.AddRange(@($pageApps, $pageSettings))
    $form.Controls.Add($tab)

    # ==================================================================
    # タブ1: アプリ管理
    # ==================================================================
    $btnPanel = New-Object System.Windows.Forms.Panel
    $btnPanel.Dock      = [System.Windows.Forms.DockStyle]::Bottom
    $btnPanel.Height    = 44
    $btnPanel.BackColor = [System.Drawing.Color]::FromArgb(37, 37, 38)
    $pageApps.Controls.Add($btnPanel)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Dock          = [System.Windows.Forms.DockStyle]::Fill
    $lv.View          = [System.Windows.Forms.View]::Details
    $lv.FullRowSelect = $true
    $lv.MultiSelect   = $false
    $lv.GridLines     = $false
    $lv.BackColor     = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $lv.ForeColor     = [System.Drawing.Color]::White
    $lv.BorderStyle   = [System.Windows.Forms.BorderStyle]::None
    $lv.Font          = New-Object System.Drawing.Font('Segoe UI', 10)

    $lv.Columns.Add('表示名',   160) | Out-Null
    $lv.Columns.Add('グループ', 100) | Out-Null
    $lv.Columns.Add('パス',     320) | Out-Null

    $pageApps.Controls.Add($lv)

    function New-ActionButton { param([string]$Text, [int]$X)
        $b = New-Object System.Windows.Forms.Button
        $b.Text      = $Text
        $b.Size      = New-Object System.Drawing.Size(80, 30)
        $b.Location  = New-Object System.Drawing.Point($X, 7)
        $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $b.BackColor = [System.Drawing.Color]::FromArgb(58, 58, 58)
        $b.ForeColor = [System.Drawing.Color]::White
        $b.FlatAppearance.BorderSize = 0
        return $b
    }

    $btnAdd    = New-ActionButton '追加'    8
    $btnEdit   = New-ActionButton '編集'    96
    $btnDelete = New-ActionButton '削除'    184
    $btnUp     = New-ActionButton '↑ 上へ'  280
    $btnDown   = New-ActionButton '↓ 下へ'  368

    $btnPanel.Controls.AddRange(@($btnAdd, $btnEdit, $btnDelete, $btnUp, $btnDown))

    # $lv をキャプチャしてアプリ一覧を再描画するスクリプトブロック
    $loadAppList = {
        $lv.Items.Clear()
        foreach ($app in $global:Config.apps) {
            $item = New-Object System.Windows.Forms.ListViewItem($app.name)
            $item.SubItems.Add($(if ($app.group) { $app.group } else { '' })) | Out-Null
            $item.SubItems.Add($app.path) | Out-Null
            $item.Tag = $app
            $lv.Items.Add($item) | Out-Null
        }
    }.GetNewClosure()

    & $loadAppList

    # 追加
    $btnAdd.add_Click({
        $newApp = Show-AppEditDialog
        if ($null -ne $newApp) {
            if (-not ($global:Config.apps -is [System.Collections.ArrayList])) {
                $global:Config.apps = [System.Collections.ArrayList]@($global:Config.apps)
            }
            $global:Config.apps.Add($newApp) | Out-Null
            Export-Config $global:Config
            & $loadAppList
        }
    }.GetNewClosure())

    # 編集
    $btnEdit.add_Click({
        if ($lv.SelectedItems.Count -eq 0) { return }
        $sel    = $lv.SelectedItems[0]
        $idx    = $sel.Index
        $edited = Show-AppEditDialog -App $sel.Tag
        if ($null -ne $edited) {
            if (-not ($global:Config.apps -is [System.Collections.ArrayList])) {
                $global:Config.apps = [System.Collections.ArrayList]@($global:Config.apps)
            }
            $global:Config.apps[$idx] = $edited
            Export-Config $global:Config
            & $loadAppList
            if ($idx -lt $lv.Items.Count) { $lv.Items[$idx].Selected = $true }
        }
    }.GetNewClosure())

    # ダブルクリックで編集
    $lv.add_DoubleClick({ $btnEdit.PerformClick() }.GetNewClosure())

    # 削除
    $btnDelete.add_Click({
        if ($lv.SelectedItems.Count -eq 0) { return }
        $idx = $lv.SelectedItems[0].Index
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "$($lv.SelectedItems[0].Text) を削除しますか？",
            'PSLauncher',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
            if (-not ($global:Config.apps -is [System.Collections.ArrayList])) {
                $global:Config.apps = [System.Collections.ArrayList]@($global:Config.apps)
            }
            $global:Config.apps.RemoveAt($idx)
            Export-Config $global:Config
            & $loadAppList
        }
    }.GetNewClosure())

    # 上へ
    $btnUp.add_Click({
        if ($lv.SelectedItems.Count -eq 0) { return }
        $idx = $lv.SelectedItems[0].Index
        if ($idx -le 0) { return }
        if (-not ($global:Config.apps -is [System.Collections.ArrayList])) {
            $global:Config.apps = [System.Collections.ArrayList]@($global:Config.apps)
        }
        $tmp = $global:Config.apps[$idx - 1]
        $global:Config.apps[$idx - 1] = $global:Config.apps[$idx]
        $global:Config.apps[$idx]     = $tmp
        Export-Config $global:Config
        & $loadAppList
        $lv.Items[$idx - 1].Selected = $true
    }.GetNewClosure())

    # 下へ
    $btnDown.add_Click({
        if ($lv.SelectedItems.Count -eq 0) { return }
        $idx = $lv.SelectedItems[0].Index
        if ($idx -ge ($global:Config.apps.Count - 1)) { return }
        if (-not ($global:Config.apps -is [System.Collections.ArrayList])) {
            $global:Config.apps = [System.Collections.ArrayList]@($global:Config.apps)
        }
        $tmp = $global:Config.apps[$idx + 1]
        $global:Config.apps[$idx + 1] = $global:Config.apps[$idx]
        $global:Config.apps[$idx]     = $tmp
        Export-Config $global:Config
        & $loadAppList
        $lv.Items[$idx + 1].Selected = $true
    }.GetNewClosure())

    # ==================================================================
    # タブ2: 設定
    # ==================================================================
    $pnl = New-Object System.Windows.Forms.Panel
    $pnl.Dock       = [System.Windows.Forms.DockStyle]::Fill
    $pnl.Padding    = New-Object System.Windows.Forms.Padding(16)
    $pnl.BackColor  = [System.Drawing.Color]::FromArgb(37, 37, 38)
    $pnl.AutoScroll = $true
    $pageSettings.Controls.Add($pnl)

    $cfg = $global:Config.settings
    $y   = 12

    function New-SectionLabel { param([string]$Text, [int]$Y)
        $l = New-Object System.Windows.Forms.Label
        $l.Text      = $Text
        $l.Location  = New-Object System.Drawing.Point(0, $Y)
        $l.Size      = New-Object System.Drawing.Size(500, 22)
        $l.Font      = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $l.ForeColor = [System.Drawing.Color]::FromArgb(0, 150, 255)
        return $l
    }

    function New-SettingLabel { param([string]$Text, [int]$Y)
        $l = New-Object System.Windows.Forms.Label
        $l.Text      = $Text
        $l.Location  = New-Object System.Drawing.Point(16, $Y)
        $l.Size      = New-Object System.Drawing.Size(160, 22)
        $l.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
        return $l
    }

    function New-SettingCombo { param([int]$Y, [string[]]$Items, [string]$Selected)
        $cb = New-Object System.Windows.Forms.ComboBox
        $cb.Location      = New-Object System.Drawing.Point(180, $Y)
        $cb.Width         = 200
        $cb.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        $cb.BackColor     = [System.Drawing.Color]::FromArgb(58, 58, 58)
        $cb.ForeColor     = [System.Drawing.Color]::White
        foreach ($it in $Items) { $cb.Items.Add($it) | Out-Null }
        $cb.SelectedItem  = $Selected
        return $cb
    }

    function New-SettingCheck { param([string]$Text, [int]$Y, [bool]$Checked)
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text      = $Text
        $chk.Location  = New-Object System.Drawing.Point(16, $Y)
        $chk.Size      = New-Object System.Drawing.Size(400, 22)
        $chk.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
        $chk.Checked   = $Checked
        return $chk
    }

    # ---- ホットコーナー ----
    $pnl.Controls.Add((New-SectionLabel 'ホットコーナー' $y)); $y += 28

    $chkHotCornerEnabled = New-SettingCheck 'ホットコーナーを有効にする' $y $cfg.hotCorner.enabled
    $pnl.Controls.Add($chkHotCornerEnabled); $y += 28

    $pnl.Controls.Add((New-SettingLabel 'コーナー位置' $y))
    $cornerMap      = @{ '左上' = 'topLeft'; '右上' = 'topRight'; '左下' = 'bottomLeft'; '右下' = 'bottomRight'; '無効' = 'disabled' }
    $cornerKeys     = @('左上', '右上', '左下', '右下', '無効')
    $cornerSelected = ($cornerMap.GetEnumerator() | Where-Object { $_.Value -eq $cfg.hotCorner.corner } | Select-Object -First 1).Key
    if (-not $cornerSelected) { $cornerSelected = '右下' }
    $cmbCorner = New-SettingCombo $y $cornerKeys $cornerSelected
    $pnl.Controls.Add($cmbCorner); $y += 30

    $pnl.Controls.Add((New-SettingLabel '判定ピクセル数' $y))
    $numPixels = New-Object System.Windows.Forms.NumericUpDown
    $numPixels.Location  = New-Object System.Drawing.Point(180, $y)
    $numPixels.Width     = 80
    $numPixels.Minimum   = 1
    $numPixels.Maximum   = 50
    $numPixels.Value     = [Math]::Max(1, [Math]::Min(50, [int]$cfg.hotCorner.pixels))
    $numPixels.BackColor = [System.Drawing.Color]::FromArgb(58, 58, 58)
    $numPixels.ForeColor = [System.Drawing.Color]::White
    $pnl.Controls.Add($numPixels); $y += 30

    $pnl.Controls.Add((New-SettingLabel 'クールダウン (ms)' $y))
    $numCooldown = New-Object System.Windows.Forms.NumericUpDown
    $numCooldown.Location  = New-Object System.Drawing.Point(180, $y)
    $numCooldown.Width     = 80
    $numCooldown.Minimum   = 200
    $numCooldown.Maximum   = 10000
    $numCooldown.Increment = 100
    $numCooldown.Value     = [Math]::Max(200, [Math]::Min(10000, [int]$cfg.hotCorner.cooldownMs))
    $numCooldown.BackColor = [System.Drawing.Color]::FromArgb(58, 58, 58)
    $numCooldown.ForeColor = [System.Drawing.Color]::White
    $pnl.Controls.Add($numCooldown); $y += 36

    # ---- ホットキー ----
    $pnl.Controls.Add((New-SectionLabel 'ホットキー' $y)); $y += 28

    $chkHotkeyEnabled = New-SettingCheck "ホットキーを有効にする ($($cfg.hotkey.display))" $y $cfg.hotkey.enabled
    $pnl.Controls.Add($chkHotkeyEnabled); $y += 28

    $pnl.Controls.Add((New-SettingLabel 'キー組み合わせ' $y))

    $hotkeyPresets = [ordered]@{
        'Ctrl+Shift+Space' = @{ modifiers = 6;  key = 32  }
        'Ctrl+Alt+Space'   = @{ modifiers = 3;  key = 32  }
        'Win+Space'        = @{ modifiers = 8;  key = 32  }
        'Ctrl+Shift+L'     = @{ modifiers = 6;  key = 76  }
        'Ctrl+Alt+L'       = @{ modifiers = 3;  key = 76  }
        'Ctrl+Shift+F1'    = @{ modifiers = 6;  key = 112 }
    }

    $cmbHotkey = New-SettingCombo $y @($hotkeyPresets.Keys) $cfg.hotkey.display
    if (-not $cmbHotkey.SelectedItem) { $cmbHotkey.SelectedIndex = 0 }
    $pnl.Controls.Add($cmbHotkey); $y += 36

    # ---- 表示設定 ----
    $pnl.Controls.Add((New-SectionLabel '表示設定' $y)); $y += 28

    $pnl.Controls.Add((New-SettingLabel '既定のビュー' $y))
    $cmbDefaultView = New-SettingCombo $y @('grid', 'list') $cfg.defaultView
    $pnl.Controls.Add($cmbDefaultView); $y += 36

    # ---- スタートアップ ----
    $pnl.Controls.Add((New-SectionLabel 'スタートアップ' $y)); $y += 28

    $chkStartup = New-SettingCheck 'Windows 起動時に自動起動する' $y ([bool]$cfg.startup)
    $pnl.Controls.Add($chkStartup); $y += 36

    # ---- 保存ボタン ----
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text      = '保存'
    $btnSave.Size      = New-Object System.Drawing.Size(100, 34)
    $btnSave.Location  = New-Object System.Drawing.Point(0, $y)
    $btnSave.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSave.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnSave.ForeColor = [System.Drawing.Color]::White
    $btnSave.FlatAppearance.BorderSize = 0
    $pnl.Controls.Add($btnSave)

    $btnSave.add_Click({
        $global:Config.settings.hotCorner.enabled    = $chkHotCornerEnabled.Checked
        $global:Config.settings.hotCorner.corner     = $cornerMap[$cmbCorner.SelectedItem]
        $global:Config.settings.hotCorner.pixels     = [int]$numPixels.Value
        $global:Config.settings.hotCorner.cooldownMs = [int]$numCooldown.Value

        $hkPreset = $hotkeyPresets[$cmbHotkey.SelectedItem]
        $global:Config.settings.hotkey.enabled   = $chkHotkeyEnabled.Checked
        $global:Config.settings.hotkey.modifiers = $hkPreset.modifiers
        $global:Config.settings.hotkey.key       = $hkPreset.key
        $global:Config.settings.hotkey.display   = $cmbHotkey.SelectedItem

        $chkHotkeyEnabled.Text = "ホットキーを有効にする ($($cmbHotkey.SelectedItem))"

        $global:Config.settings.defaultView = $cmbDefaultView.SelectedItem

        $global:Config.settings.startup = $chkStartup.Checked
        Set-Startup $chkStartup.Checked

        Export-Config $global:Config

        [System.Windows.Forms.MessageBox]::Show(
            '設定を保存しました。',
            'PSLauncher',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }.GetNewClosure())

    return $form
}

# ---------------------------------------------------------------------------
# スタートアップ登録 / 解除
# ---------------------------------------------------------------------------
function Set-Startup {
    param([bool]$Enable)

    $startupDir = [System.Environment]::GetFolderPath('Startup')
    $lnkPath    = Join-Path $startupDir 'PSLauncher.lnk'
    $targetPs1  = Join-Path $global:ScriptDir 'PSLauncher.ps1'

    if ($Enable) {
        try {
            $wsh      = New-Object -ComObject WScript.Shell
            $shortcut = $wsh.CreateShortcut($lnkPath)
            $shortcut.TargetPath       = 'powershell.exe'
            $shortcut.Arguments        = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$targetPs1`""
            $shortcut.WorkingDirectory = $global:ScriptDir
            $shortcut.WindowStyle      = 7
            $shortcut.Description      = 'PSLauncher'
            $shortcut.Save()
        } catch {
            Write-Warning "スタートアップ登録に失敗しました: $_"
        }
    } else {
        if (Test-Path $lnkPath) {
            try { Remove-Item $lnkPath -Force } catch {
                Write-Warning "スタートアップ解除に失敗しました: $_"
            }
        }
    }
}
