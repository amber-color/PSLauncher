#
# LauncherWindow.ps1 - ランチャーウィンドウ（グリッド表示 / リスト表示）
#
# ── スコープ設計 ──────────────────────────────────────────────────────────
#   PowerShell でイベントハンドラから呼ぶ scriptblock を入れ子にすると、
#   .GetNewClosure() でキャプチャした変数がネスト越しに届かないことがある。
#   対策: 全ての共有状態・操作を $s（ハッシュテーブル）に格納する。
#   $s はオブジェクト参照のため全クロージャが同じインスタンスを共有する。
# ─────────────────────────────────────────────────────────────────────────

# ---------------------------------------------------------------------------
# ヘルパー: アイコンビットマップ取得
# ---------------------------------------------------------------------------
function Get-AppIcon {
    param(
        [string]$IconPath,
        [string]$ExePath,
        [int]$Size = 48
    )

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
            $g.DrawIcon($icon, (New-Object System.Drawing.Rectangle(0, 0, $Size, $Size)))
            $g.Dispose(); $icon.Dispose()
            return $bmp
        } catch {}
    }

    $bmp  = New-Object System.Drawing.Bitmap($Size, $Size)
    $g    = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.FillRectangle(
        (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 120, 215))),
        0, 0, $Size, $Size)
    $initial  = if ($ExePath) { [System.IO.Path]::GetFileNameWithoutExtension($ExePath) } else { '?' }
    $initial  = $initial.Substring(0, 1).ToUpper()
    $fontSize = [int]($Size * 0.45)
    $font     = New-Object System.Drawing.Font('Segoe UI', $fontSize, [System.Drawing.FontStyle]::Bold)
    $sf       = New-Object System.Drawing.StringFormat
    $sf.Alignment     = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString($initial, $font, [System.Drawing.Brushes]::White,
        (New-Object System.Drawing.RectangleF(0, 0, $Size, $Size)), $sf)
    $g.Dispose()
    return $bmp
}

# ---------------------------------------------------------------------------
# ランチャーウィンドウ生成
# ---------------------------------------------------------------------------
function New-LauncherWindow {

    # ------------------------------------------------------------------
    # フォーム
    # ------------------------------------------------------------------
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'PSLauncher'
    $form.Width           = [int]$global:Config.settings.windowWidth
    $form.Height          = [int]$global:Config.settings.windowHeight
    $form.MinimumSize     = New-Object System.Drawing.Size(300, 200)
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::Manual
    $form.BackColor       = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $form.ForeColor       = [System.Drawing.Color]::White
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.ShowInTaskbar   = $false
    $form.TopMost         = $true

    # ------------------------------------------------------------------
    # ヘッダーパネル
    # ------------------------------------------------------------------
    $header = New-Object System.Windows.Forms.Panel
    $header.Dock      = [System.Windows.Forms.DockStyle]::Top
    $header.Height    = 44
    $header.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location    = New-Object System.Drawing.Point(8, 10)
    $searchBox.Height      = 24
    $searchBox.BackColor   = [System.Drawing.Color]::FromArgb(58, 58, 58)
    $searchBox.ForeColor   = [System.Drawing.Color]::FromArgb(160, 160, 160)
    $searchBox.Font        = New-Object System.Drawing.Font('Segoe UI', 10)
    $searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $searchBox.Text        = '検索...'
    $searchBox.Tag         = $false   # プレースホルダ表示中フラグ

    $btnGrid = New-Object System.Windows.Forms.Button
    $btnGrid.Text      = '▦'
    $btnGrid.Size      = New-Object System.Drawing.Size(32, 26)
    $btnGrid.Location  = New-Object System.Drawing.Point(0, 9)
    $btnGrid.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGrid.FlatAppearance.BorderSize = 0
    $btnGrid.ForeColor = [System.Drawing.Color]::White
    $btnGrid.Font      = New-Object System.Drawing.Font('Segoe UI', 10)

    $btnList = New-Object System.Windows.Forms.Button
    $btnList.Text      = '☰'
    $btnList.Size      = New-Object System.Drawing.Size(32, 26)
    $btnList.Location  = New-Object System.Drawing.Point(0, 9)
    $btnList.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnList.FlatAppearance.BorderSize = 0
    $btnList.ForeColor = [System.Drawing.Color]::White
    $btnList.Font      = New-Object System.Drawing.Font('Segoe UI', 10)

    $header.Controls.AddRange(@($searchBox, $btnGrid, $btnList))
    $form.Controls.Add($header)

    # ------------------------------------------------------------------
    # コンテンツパネル
    # ------------------------------------------------------------------
    $content = New-Object System.Windows.Forms.Panel
    $content.Dock       = [System.Windows.Forms.DockStyle]::Fill
    $content.BackColor  = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $content.AutoScroll = $true
    $form.Controls.Add($content)

    # ==================================================================
    # $s: 全クロージャが共有するハッシュテーブル
    #   - コントロール参照 / ビューモード / 操作 scriptblock を格納
    #   - ハッシュテーブルはオブジェクト参照なので全クロージャが同一インスタンスを見る
    # ==================================================================
    $s = @{
        form      = $form
        searchBox = $searchBox
        btnGrid   = $btnGrid
        btnList   = $btnList
        content   = $content
        viewMode  = [string]$global:Config.settings.defaultView
    }

    # ------------------------------------------------------------------
    # 操作 scriptblock（$s 経由でコントロールに触る）
    # GetNewClosure() で $s をキャプチャ → $s はハッシュテーブル参照なので
    # 後から追加したエントリも全クロージャから参照可能
    # ------------------------------------------------------------------

    $s.updateButtonStates = {
        if ($s.viewMode -eq 'grid') {
            $s.btnGrid.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
            $s.btnList.BackColor = [System.Drawing.Color]::FromArgb(58, 58, 58)
        } else {
            $s.btnList.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
            $s.btnGrid.BackColor = [System.Drawing.Color]::FromArgb(58, 58, 58)
        }
    }.GetNewClosure()

    $s.updateHeaderLayout = {
        $w = $s.form.ClientSize.Width
        $s.searchBox.Width  = $w - 90
        $s.btnGrid.Location = New-Object System.Drawing.Point(($w - 80), 9)
        $s.btnList.Location = New-Object System.Drawing.Point(($w - 44), 9)
    }.GetNewClosure()

    $s.showGridView = {
        param($AppList)

        $s.content.Controls.Clear()
        $iconSize = 48
        $cellW    = $iconSize + 24
        $cellH    = $iconSize + 28
        $padLeft  = 8
        $padTop   = 8
        $cols     = [Math]::Max(1, [Math]::Floor(($s.content.ClientSize.Width - $padLeft) / $cellW))
        $col      = 0
        $row      = 0

        foreach ($app in $AppList) {
            $appRef = $app   # ループ変数を確定させるローカルコピー

            $cell = New-Object System.Windows.Forms.Panel
            $cell.Size      = New-Object System.Drawing.Size($cellW, $cellH)
            $cell.Location  = New-Object System.Drawing.Point(
                ($padLeft + $col * $cellW), ($padTop + $row * $cellH))
            $cell.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
            $cell.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $cell.Tag       = $appRef

            $pic = New-Object System.Windows.Forms.PictureBox
            $pic.Size      = New-Object System.Drawing.Size($iconSize, $iconSize)
            $pic.Location  = New-Object System.Drawing.Point([int](($cellW - $iconSize) / 2), 2)
            $pic.SizeMode  = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
            $pic.BackColor = [System.Drawing.Color]::Transparent
            $pic.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $pic.Tag       = $appRef
            try { $pic.Image = Get-AppIcon -IconPath $appRef.iconPath -ExePath $appRef.path -Size $iconSize } catch {}

            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text      = $appRef.name
            $lbl.Size      = New-Object System.Drawing.Size($cellW, 22)
            $lbl.Location  = New-Object System.Drawing.Point(0, ($iconSize + 4))
            $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $lbl.ForeColor = [System.Drawing.Color]::FromArgb(210, 210, 210)
            $lbl.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
            $lbl.BackColor = [System.Drawing.Color]::Transparent
            $lbl.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $lbl.Tag       = $appRef

            $hoverBg  = [System.Drawing.Color]::FromArgb(55, 55, 55)
            $normalBg = [System.Drawing.Color]::FromArgb(28, 28, 28)
            # ループ変数 $cell / $hoverBg / $normalBg をキャプチャ
            $enterSB  = { $cell.BackColor = $hoverBg  }.GetNewClosure()
            $leaveSB  = { $cell.BackColor = $normalBg }.GetNewClosure()
            # $cell と $s をキャプチャ（$s 経由で form を参照）
            $launchSB = {
                $data = $cell.Tag
                try {
                    Start-Process -FilePath $data.path
                    $s.form.Hide()
                } catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        "起動に失敗しました: $($data.name)`n$_", 'PSLauncher',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                }
            }.GetNewClosure()

            foreach ($ctrl in @($cell, $pic, $lbl)) {
                $ctrl.add_MouseEnter($enterSB)
                $ctrl.add_MouseLeave($leaveSB)
                $ctrl.add_Click($launchSB)
            }

            $cell.Controls.AddRange(@($pic, $lbl))
            $s.content.Controls.Add($cell)

            $col++
            if ($col -ge $cols) { $col = 0; $row++ }
        }
    }.GetNewClosure()

    $s.showListView = {
        param($AppList)

        $s.content.Controls.Clear()

        $lv = New-Object System.Windows.Forms.ListView
        $lv.Dock          = [System.Windows.Forms.DockStyle]::Fill
        $lv.View          = [System.Windows.Forms.View]::Details
        $lv.BackColor     = [System.Drawing.Color]::FromArgb(28, 28, 28)
        $lv.ForeColor     = [System.Drawing.Color]::FromArgb(210, 210, 210)
        $lv.FullRowSelect = $true
        $lv.GridLines     = $false
        $lv.BorderStyle   = [System.Windows.Forms.BorderStyle]::None
        $lv.Font          = New-Object System.Drawing.Font('Segoe UI', 10)
        $lv.HeaderStyle   = [System.Windows.Forms.ColumnHeaderStyle]::None

        $lv.Columns.Add('アプリ名', 200) | Out-Null
        $lv.Columns.Add('グループ', 100) | Out-Null
        $lv.Columns.Add('パス',     350) | Out-Null

        $il = New-Object System.Windows.Forms.ImageList
        $il.ImageSize  = New-Object System.Drawing.Size(24, 24)
        $il.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit

        foreach ($app in $AppList) {
            try {
                $il.Images.Add((Get-AppIcon -IconPath $app.iconPath -ExePath $app.path -Size 24)) | Out-Null
            } catch {
                $il.Images.Add((New-Object System.Drawing.Bitmap(24, 24))) | Out-Null
            }
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

        # $lv と $s をキャプチャ
        $lv.add_DoubleClick({
            if ($lv.SelectedItems.Count -gt 0) {
                $app = $lv.SelectedItems[0].Tag
                try {
                    Start-Process -FilePath $app.path
                    $s.form.Hide()
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

    $s.refreshApps = {
        $searchText = if ([bool]$s.searchBox.Tag) { $s.searchBox.Text } else { '' }
        $apps = @($global:Config.apps)
        if ($searchText -ne '') {
            $apps = @($apps | Where-Object {
                ($_.name  -and $_.name  -like "*$searchText*") -or
                ($_.group -and $_.group -like "*$searchText*")
            })
        }
        if ($s.viewMode -eq 'grid') { & $s.showGridView $apps }
        else                        { & $s.showListView $apps }
    }.GetNewClosure()

    # ==================================================================
    # イベント登録（.GetNewClosure() で $s をキャプチャ）
    # ==================================================================

    $form.add_Deactivate({
        $s.form.Hide()
    }.GetNewClosure())

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
            $s.searchBox.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
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
        & $s.updateHeaderLayout
        & $s.updateButtonStates
        & $s.refreshApps
        $s.searchBox.Focus()
    }.GetNewClosure())

    return $form
}
