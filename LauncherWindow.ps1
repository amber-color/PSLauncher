#
# LauncherWindow.ps1 - ランチャーウィンドウ（グリッド表示 / リスト表示）
#

function Get-AppIcon {
    param([string]$IconPath, [string]$ExePath, [int]$Size = 48)

    if ($IconPath -and (Test-Path $IconPath)) {
        try {
            $img = [System.Drawing.Image]::FromFile($IconPath)
            return New-Object System.Drawing.Bitmap($img, $Size, $Size)
        } catch {}
    }
    if ($ExePath -and (Test-Path $ExePath)) {
        try {
            $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ExePath)
            $bmp  = New-Object System.Drawing.Bitmap($Size, $Size)
            $g    = [System.Drawing.Graphics]::FromImage($bmp)
            $g.DrawIcon($icon, (New-Object System.Drawing.Rectangle(0,0,$Size,$Size)))
            $g.Dispose(); $icon.Dispose()
            return $bmp
        } catch {}
    }
    $bmp  = New-Object System.Drawing.Bitmap($Size, $Size)
    $g    = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0,120,215))), 0,0,$Size,$Size)
    $initial  = if ($ExePath) { [System.IO.Path]::GetFileNameWithoutExtension($ExePath) } else { '?' }
    $initial  = $initial.Substring(0,1).ToUpper()
    $font     = New-Object System.Drawing.Font('Segoe UI', [int]($Size*0.45), [System.Drawing.FontStyle]::Bold)
    $sf       = New-Object System.Drawing.StringFormat
    $sf.Alignment = $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString($initial, $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF(0,0,$Size,$Size)), $sf)
    $g.Dispose()
    return $bmp
}

function New-LauncherWindow {

    # ----------------------------------------------------------------
    # フォーム
    # ----------------------------------------------------------------
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'PSLauncher'
    $form.Width           = [int]$global:Config.settings.windowWidth
    $form.Height          = [int]$global:Config.settings.windowHeight
    $form.MinimumSize     = New-Object System.Drawing.Size(300, 250)
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::Manual
    $form.BackColor       = [System.Drawing.Color]::FromArgb(28,28,28)
    $form.ForeColor       = [System.Drawing.Color]::White
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.ShowInTaskbar   = $false
    $form.TopMost         = $true

    # ----------------------------------------------------------------
    # ヘッダー (Dock=Top, h=44)
    # ----------------------------------------------------------------
    $header = New-Object System.Windows.Forms.Panel
    $header.Dock      = [System.Windows.Forms.DockStyle]::Top
    $header.Height    = 44
    $header.BackColor = [System.Drawing.Color]::FromArgb(40,40,40)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location    = New-Object System.Drawing.Point(8,10)
    $searchBox.Height      = 24
    $searchBox.BackColor   = [System.Drawing.Color]::FromArgb(58,58,58)
    $searchBox.ForeColor   = [System.Drawing.Color]::FromArgb(160,160,160)
    $searchBox.Font        = New-Object System.Drawing.Font('Segoe UI', 10)
    $searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $searchBox.Text        = '検索...'
    $searchBox.Tag         = $false

    $btnGrid = New-Object System.Windows.Forms.Button
    $btnGrid.Text = '▦'; $btnGrid.Size = New-Object System.Drawing.Size(32,26)
    $btnGrid.Location = New-Object System.Drawing.Point(0,9)
    $btnGrid.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGrid.FlatAppearance.BorderSize = 0
    $btnGrid.ForeColor = [System.Drawing.Color]::White
    $btnGrid.Font = New-Object System.Drawing.Font('Segoe UI', 10)

    $btnList = New-Object System.Windows.Forms.Button
    $btnList.Text = '☰'; $btnList.Size = New-Object System.Drawing.Size(32,26)
    $btnList.Location = New-Object System.Drawing.Point(0,9)
    $btnList.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnList.FlatAppearance.BorderSize = 0
    $btnList.ForeColor = [System.Drawing.Color]::White
    $btnList.Font = New-Object System.Drawing.Font('Segoe UI', 10)

    $header.Controls.AddRange(@($searchBox, $btnGrid, $btnList))
    $form.Controls.Add($header)

    # ----------------------------------------------------------------
    # タブパネル (Dock=Top, h=34) — ヘッダーの直下
    # ----------------------------------------------------------------
    $tabPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $tabPanel.Dock          = [System.Windows.Forms.DockStyle]::Top
    $tabPanel.Height        = 34
    $tabPanel.BackColor     = [System.Drawing.Color]::FromArgb(35,35,35)
    $tabPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $tabPanel.WrapContents  = $false
    $tabPanel.Padding       = New-Object System.Windows.Forms.Padding(4,4,4,0)
    $form.Controls.Add($tabPanel)

    # ----------------------------------------------------------------
    # コンテンツパネル (Dock=Fill)
    # ----------------------------------------------------------------
    $content = New-Object System.Windows.Forms.Panel
    $content.Dock       = [System.Windows.Forms.DockStyle]::Fill
    $content.BackColor  = [System.Drawing.Color]::FromArgb(28,28,28)
    $content.AutoScroll = $true
    $form.Controls.Add($content)

    # ----------------------------------------------------------------
    # 共有状態
    # ----------------------------------------------------------------
    $s = @{
        form           = $form
        searchBox      = $searchBox
        btnGrid        = $btnGrid
        btnList        = $btnList
        tabPanel       = $tabPanel
        content        = $content
        viewMode       = [string]$global:Config.settings.defaultView
        currentTab     = 'すべて'
        lastShownApps  = @()
        dragStartPos   = $null
        dragSourceApp  = $null
        dragSourceCell = $null
    }

    # ----------------------------------------------------------------
    # updateButtonStates
    # ----------------------------------------------------------------
    $s.updateButtonStates = {
        if ($s.viewMode -eq 'grid') {
            $s.btnGrid.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
            $s.btnList.BackColor = [System.Drawing.Color]::FromArgb(58,58,58)
        } else {
            $s.btnList.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
            $s.btnGrid.BackColor = [System.Drawing.Color]::FromArgb(58,58,58)
        }
    }.GetNewClosure()

    # ----------------------------------------------------------------
    # updateHeaderLayout
    # ----------------------------------------------------------------
    $s.updateHeaderLayout = {
        $w = $s.form.ClientSize.Width
        $s.searchBox.Width  = $w - 96
        $s.btnGrid.Location = New-Object System.Drawing.Point(($w-80), 9)
        $s.btnList.Location = New-Object System.Drawing.Point(($w-44), 9)
    }.GetNewClosure()

    # ----------------------------------------------------------------
    # updateTabPanel — グループ一覧からタブボタンを生成
    # ----------------------------------------------------------------
    $s.updateTabPanel = {
        $s.tabPanel.Controls.Clear()

        $groups = @('すべて')
        $seen   = @{}
        foreach ($app in $global:Config.apps) {
            if ($app.group -and $app.group -ne '' -and -not $seen.ContainsKey($app.group)) {
                $groups += $app.group
                $seen[$app.group] = $true
            }
        }

        foreach ($gn in $groups) {
            $gName    = $gn
            $isActive = ($gName -eq $s.currentTab)

            $btn = New-Object System.Windows.Forms.Button
            $btn.Text      = $gName
            $btn.AutoSize  = $true
            $btn.Height    = 26
            $btn.Margin    = New-Object System.Windows.Forms.Padding(2,0,2,0)
            $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btn.FlatAppearance.BorderSize = 1
            $btn.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
            $btn.ForeColor = [System.Drawing.Color]::White
            $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
            if ($isActive) {
                $btn.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
                $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0,120,215)
            } else {
                $btn.BackColor = [System.Drawing.Color]::FromArgb(58,58,58)
                $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80,80,80)
            }

            $btn.add_Click({
                $s.currentTab = $gName
                & $s.updateTabPanel
                & $s.refreshApps
            }.GetNewClosure())

            $s.tabPanel.Controls.Add($btn)
        }
    }.GetNewClosure()

    # ----------------------------------------------------------------
    # showGridView — ドラッグ&ドロップ並び替え付き
    # ----------------------------------------------------------------
    $s.showGridView = {
        param($AppList)

        $s.content.AutoScrollPosition = New-Object System.Drawing.Point(0,0)
        $s.content.Controls.Clear()
        $s.lastShownApps = $AppList

        $iconSize = 48
        $cellW    = $iconSize + 24   # 72
        $cellH    = $iconSize + 28   # 76
        $padLeft  = 8
        $padTop   = 8
        $cols     = [Math]::Max(1, [Math]::Floor(($s.content.ClientSize.Width - $padLeft) / $cellW))
        $col = 0; $row = 0

        foreach ($app in $AppList) {
            $appRef = $app

            $cell = New-Object System.Windows.Forms.Panel
            $cell.Size      = New-Object System.Drawing.Size($cellW, $cellH)
            $cell.Location  = New-Object System.Drawing.Point(
                ($padLeft + $col*$cellW), ($padTop + $row*$cellH))
            $cell.BackColor = [System.Drawing.Color]::FromArgb(28,28,28)
            $cell.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $cell.Tag       = $appRef
            $cell.AllowDrop = $true

            $pic = New-Object System.Windows.Forms.PictureBox
            $pic.Size      = New-Object System.Drawing.Size($iconSize, $iconSize)
            $pic.Location  = New-Object System.Drawing.Point([int](($cellW-$iconSize)/2), 2)
            $pic.SizeMode  = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
            $pic.BackColor = [System.Drawing.Color]::Transparent
            $pic.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $pic.Tag       = $appRef
            $pic.AllowDrop = $true
            try { $pic.Image = Get-AppIcon -IconPath $appRef.iconPath -ExePath $appRef.path -Size $iconSize } catch {}

            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text      = $appRef.name
            $lbl.Size      = New-Object System.Drawing.Size($cellW, 22)
            $lbl.Location  = New-Object System.Drawing.Point(0, ($iconSize+4))
            $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $lbl.ForeColor = [System.Drawing.Color]::FromArgb(210,210,210)
            $lbl.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
            $lbl.BackColor = [System.Drawing.Color]::Transparent
            $lbl.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $lbl.Tag       = $appRef
            $lbl.AllowDrop = $true

            $hoverBg  = [System.Drawing.Color]::FromArgb(55,55,55)
            $normalBg = [System.Drawing.Color]::FromArgb(28,28,28)
            $dragBg   = [System.Drawing.Color]::FromArgb(0,100,190)

            $enterSB = { $cell.BackColor = $hoverBg  }.GetNewClosure()
            $leaveSB = { $cell.BackColor = $normalBg }.GetNewClosure()

            $launchSB = {
                if ($null -ne $s.dragSourceApp) { return }  # ドラッグ中は起動しない
                if ([string]::IsNullOrEmpty($appRef.path)) { return }
                try {
                    Start-Process $appRef.path
                    $global:LauncherForm.Hide()
                } catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        "起動に失敗しました: $($appRef.name)`n$_", 'PSLauncher',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                }
            }.GetNewClosure()

            # ドラッグ開始検出
            $mouseDownSB = {
                if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                    $s.dragStartPos   = New-Object System.Drawing.Point($_.X, $_.Y)
                    $s.dragSourceApp  = $null   # まだ開始していない
                    $s.dragSourceCell = $cell
                }
            }.GetNewClosure()

            $mouseMoveSB = {
                if ($null -eq $s.dragStartPos) { return }
                $ds = [System.Windows.Forms.SystemInformation]::DragSize
                if ([Math]::Abs($_.X - $s.dragStartPos.X) -gt $ds.Width -or
                    [Math]::Abs($_.Y - $s.dragStartPos.Y) -gt $ds.Height) {
                    $s.dragStartPos  = $null
                    $s.dragSourceApp = $appRef
                    $cell.DoDragDrop($appRef, [System.Windows.Forms.DragDropEffects]::Move) | Out-Null
                    $s.dragSourceApp  = $null
                    $s.dragSourceCell = $null
                }
            }.GetNewClosure()

            $mouseUpSB = {
                $s.dragStartPos   = $null
                $s.dragSourceCell = $null
                # dragSourceApp は DoDragDrop が呼ばれなかった場合のみクリア
                if ($null -eq $s.dragSourceCell) { $s.dragSourceApp = $null }
            }.GetNewClosure()

            # ドロップ先のハイライト
            $dragOverSB = {
                $_.Effect = [System.Windows.Forms.DragDropEffects]::Move
                if (-not [object]::ReferenceEquals($appRef, $s.dragSourceApp)) {
                    $cell.BackColor = $dragBg
                }
            }.GetNewClosure()

            $dragLeaveSB = {
                $cell.BackColor = $normalBg
            }.GetNewClosure()

            # ドロップ処理（並び替え確認）
            $dragDropSB = {
                $cell.BackColor = $normalBg
                $srcApp = $s.dragSourceApp
                $tgtApp = $appRef

                if ($null -eq $srcApp -or [object]::ReferenceEquals($srcApp, $tgtApp)) { return }

                $ans = [System.Windows.Forms.MessageBox]::Show(
                    "「$($srcApp.name)」を「$($tgtApp.name)」の前に移動しますか？",
                    'PSLauncher',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question)

                if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }

                # 表示中リストでの位置を特定
                $shown = @($s.lastShownApps)
                $srcIdx = -1; $tgtIdx = -1
                for ($i = 0; $i -lt $shown.Length; $i++) {
                    if ([object]::ReferenceEquals($shown[$i], $srcApp)) { $srcIdx = $i }
                    if ([object]::ReferenceEquals($shown[$i], $tgtApp)) { $tgtIdx = $i }
                }
                if ($srcIdx -lt 0 -or $tgtIdx -lt 0) { return }

                # 表示リスト内で並び替え
                $newShown = [System.Collections.ArrayList]@($shown)
                $newShown.RemoveAt($srcIdx)
                $insIdx = if ($srcIdx -lt $tgtIdx) { $tgtIdx - 1 } else { $tgtIdx }
                $newShown.Insert($insIdx, $srcApp)

                # 全アプリリストに反映（フィルタ外のアプリは位置を保持）
                $allApps  = [System.Collections.ArrayList]@($global:Config.apps)
                $origIdxs = [System.Collections.ArrayList]::new()
                foreach ($sa in $shown) {
                    for ($j = 0; $j -lt $allApps.Count; $j++) {
                        if ([object]::ReferenceEquals($allApps[$j], $sa)) {
                            $origIdxs.Add($j) | Out-Null; break
                        }
                    }
                }
                for ($i = 0; $i -lt $origIdxs.Count; $i++) {
                    $allApps[$origIdxs[$i]] = $newShown[$i]
                }
                $global:Config.apps = @($allApps)
                Export-Config $global:Config
                & $s.refreshApps
            }.GetNewClosure()

            foreach ($ctrl in @($cell, $pic, $lbl)) {
                $ctrl.add_MouseEnter($enterSB)
                $ctrl.add_MouseLeave($leaveSB)
                $ctrl.add_Click($launchSB)
                $ctrl.add_MouseDown($mouseDownSB)
                $ctrl.add_MouseMove($mouseMoveSB)
                $ctrl.add_MouseUp($mouseUpSB)
                $ctrl.add_DragOver($dragOverSB)
                $ctrl.add_DragLeave($dragLeaveSB)
                $ctrl.add_DragDrop($dragDropSB)
            }

            $cell.Controls.AddRange(@($pic, $lbl))
            $s.content.Controls.Add($cell)

            $col++
            if ($col -ge $cols) { $col = 0; $row++ }
        }
    }.GetNewClosure()

    # ----------------------------------------------------------------
    # showListView
    # ----------------------------------------------------------------
    $s.showListView = {
        param($AppList)

        $s.content.Controls.Clear()
        $s.lastShownApps = $AppList

        $lv = New-Object System.Windows.Forms.ListView
        $lv.Dock          = [System.Windows.Forms.DockStyle]::Fill
        $lv.View          = [System.Windows.Forms.View]::Details
        $lv.BackColor     = [System.Drawing.Color]::FromArgb(28,28,28)
        $lv.ForeColor     = [System.Drawing.Color]::FromArgb(210,210,210)
        $lv.FullRowSelect = $true
        $lv.GridLines     = $false
        $lv.BorderStyle   = [System.Windows.Forms.BorderStyle]::None
        $lv.Font          = New-Object System.Drawing.Font('Segoe UI', 10)
        $lv.HeaderStyle   = [System.Windows.Forms.ColumnHeaderStyle]::None

        $lv.Columns.Add('アプリ名', 200) | Out-Null
        $lv.Columns.Add('グループ', 100) | Out-Null
        $lv.Columns.Add('パス',     350) | Out-Null

        $il = New-Object System.Windows.Forms.ImageList
        $il.ImageSize  = New-Object System.Drawing.Size(24,24)
        $il.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit
        foreach ($app in $AppList) {
            try { $il.Images.Add((Get-AppIcon -IconPath $app.iconPath -ExePath $app.path -Size 24)) | Out-Null }
            catch { $il.Images.Add((New-Object System.Drawing.Bitmap(24,24))) | Out-Null }
        }
        $lv.SmallImageList = $il

        $idx = 0
        foreach ($app in $AppList) {
            $item = New-Object System.Windows.Forms.ListViewItem($app.name, $idx)
            $item.SubItems.Add($(if ($app.group) { $app.group } else { '' })) | Out-Null
            $item.SubItems.Add($app.path) | Out-Null
            $item.Tag = $app
            $lv.Items.Add($item) | Out-Null
            $idx++
        }

        $lv.add_DoubleClick({
            if ($lv.SelectedItems.Count -gt 0) {
                $app = $lv.SelectedItems[0].Tag
                if ($null -eq $app -or [string]::IsNullOrEmpty($app.path)) { return }
                try {
                    Start-Process $app.path
                    $global:LauncherForm.Hide()
                } catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        "起動に失敗しました: $($app.name)`n$_", 'PSLauncher',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                }
            }
        }.GetNewClosure())

        $s.content.Controls.Add($lv)
    }.GetNewClosure()

    # ----------------------------------------------------------------
    # refreshApps — タブ・検索でフィルタしてビュー更新
    # ----------------------------------------------------------------
    $s.refreshApps = {
        $searchText = if ([bool]$s.searchBox.Tag) { $s.searchBox.Text } else { '' }

        $apps = @($global:Config.apps)

        if ($s.currentTab -ne 'すべて') {
            $apps = @($apps | Where-Object { $_.group -eq $s.currentTab })
        }
        if ($searchText -ne '') {
            $apps = @($apps | Where-Object {
                ($_.name  -and $_.name  -like "*$searchText*") -or
                ($_.group -and $_.group -like "*$searchText*")
            })
        }

        if ($s.viewMode -eq 'grid') { & $s.showGridView $apps }
        else                        { & $s.showListView $apps }
    }.GetNewClosure()

    # ----------------------------------------------------------------
    # イベント登録
    # ----------------------------------------------------------------
    $form.add_Deactivate({ $s.form.Hide() }.GetNewClosure())

    $form.add_ResizeEnd({
        $global:Config.settings.windowWidth  = $s.form.Width
        $global:Config.settings.windowHeight = $s.form.Height
        Export-Config $global:Config
    }.GetNewClosure())

    $searchBox.add_Enter({
        if (-not [bool]$s.searchBox.Tag) {
            $s.searchBox.Text      = ''
            $s.searchBox.ForeColor = [System.Drawing.Color]::White
            $s.searchBox.Tag       = $true
        }
    }.GetNewClosure())

    $searchBox.add_Leave({
        if ($s.searchBox.Text -eq '') {
            $s.searchBox.Text      = '検索...'
            $s.searchBox.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
            $s.searchBox.Tag       = $false
        }
    }.GetNewClosure())

    $searchBox.add_TextChanged({
        if ([bool]$s.searchBox.Tag) { & $s.refreshApps }
    }.GetNewClosure())

    $btnGrid.add_Click({
        $s.viewMode = 'grid'
        & $s.updateButtonStates
        & $s.refreshApps
    }.GetNewClosure())

    $btnList.add_Click({
        $s.viewMode = 'list'
        & $s.updateButtonStates
        & $s.refreshApps
    }.GetNewClosure())

    $form.add_Resize({
        & $s.updateHeaderLayout
        if ($s.viewMode -eq 'grid') { & $s.refreshApps }
    }.GetNewClosure())

    $form.add_Shown({
        $global:Config = Import-Config
        # currentTab が存在しないグループを指していたらリセット
        $validGroups = @('すべて') + @($global:Config.apps |
            Where-Object { $_.group } | ForEach-Object { $_.group } | Select-Object -Unique)
        if ($s.currentTab -notin $validGroups) { $s.currentTab = 'すべて' }
        & $s.updateHeaderLayout
        & $s.updateButtonStates
        & $s.updateTabPanel
        & $s.refreshApps
        $s.searchBox.Focus()
    }.GetNewClosure())

    return $form
}
