#
# LauncherWindow.ps1 - ランチャーウィンドウ（グリッド表示 / リスト表示）
#

# ---------------------------------------------------------------------------
# ヘルパー: アイコンビットマップ取得
# ---------------------------------------------------------------------------
function Get-AppIcon {
    param(
        [string]$IconPath,
        [string]$ExePath,
        [int]$Size = 48
    )

    # 1. 明示的なアイコンパス
    if ($IconPath -and (Test-Path $IconPath)) {
        try {
            $img = [System.Drawing.Image]::FromFile($IconPath)
            return New-Object System.Drawing.Bitmap($img, $Size, $Size)
        } catch {}
    }

    # 2. 実行ファイルからアイコン抽出
    if ($ExePath -and (Test-Path $ExePath)) {
        try {
            $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ExePath)
            $bmp  = New-Object System.Drawing.Bitmap($Size, $Size)
            $g    = [System.Drawing.Graphics]::FromImage($bmp)
            $g.DrawIcon($icon, (New-Object System.Drawing.Rectangle(0, 0, $Size, $Size)))
            $g.Dispose()
            $icon.Dispose()
            return $bmp
        } catch {}
    }

    # 3. デフォルト（頭文字）
    $bmp  = New-Object System.Drawing.Bitmap($Size, $Size)
    $g    = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.FillRectangle(
        (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 120, 215))),
        0, 0, $Size, $Size
    )
    $initial = if ($ExePath) { [System.IO.Path]::GetFileNameWithoutExtension($ExePath) } else { '?' }
    $initial = $initial.Substring(0, 1).ToUpper()
    $fontSize = [int]($Size * 0.45)
    $font     = New-Object System.Drawing.Font('Segoe UI', $fontSize, [System.Drawing.FontStyle]::Bold)
    $sf       = New-Object System.Drawing.StringFormat
    $sf.Alignment     = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString($initial, $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF(0, 0, $Size, $Size)), $sf)
    $g.Dispose()
    return $bmp
}

# ---------------------------------------------------------------------------
# ランチャーウィンドウ生成
# ---------------------------------------------------------------------------
function New-LauncherWindow {

    # ------------------------------------------------------------------
    # フォーム本体
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

    # フォーカスが外れたら自動的に隠す
    $form.add_Deactivate({ $form.Hide() })

    # ウィンドウサイズ変更時に設定を更新
    $form.add_ResizeEnd({
        $global:Config.settings.windowWidth  = $form.Width
        $global:Config.settings.windowHeight = $form.Height
        Export-Config $global:Config
    })

    # ------------------------------------------------------------------
    # ヘッダーパネル
    # ------------------------------------------------------------------
    $header = New-Object System.Windows.Forms.Panel
    $header.Dock      = [System.Windows.Forms.DockStyle]::Top
    $header.Height    = 44
    $header.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)

    # 検索ボックス
    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location  = New-Object System.Drawing.Point(8, 10)
    $searchBox.Height    = 24
    $searchBox.BackColor = [System.Drawing.Color]::FromArgb(58, 58, 58)
    $searchBox.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
    $searchBox.Font      = New-Object System.Drawing.Font('Segoe UI', 10)
    $searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $searchBox.Text      = '検索...'
    $searchBox.Tag       = $false  # placeholder表示中フラグ

    $searchBox.add_Enter({
        if (-not [bool]$searchBox.Tag) {
            $searchBox.Text      = ''
            $searchBox.ForeColor = [System.Drawing.Color]::White
            $searchBox.Tag       = $true
        }
    })
    $searchBox.add_Leave({
        if ($searchBox.Text -eq '') {
            $searchBox.Text      = '検索...'
            $searchBox.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
            $searchBox.Tag       = $false
        }
    })

    # グリッドボタン / リストボタン
    function New-ViewButton {
        param([string]$Label, [int]$X)
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text      = $Label
        $btn.Size      = New-Object System.Drawing.Size(32, 26)
        $btn.Location  = New-Object System.Drawing.Point($X, 9)
        $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btn.FlatAppearance.BorderSize = 0
        $btn.ForeColor = [System.Drawing.Color]::White
        $btn.Font      = New-Object System.Drawing.Font('Segoe UI', 10)
        return $btn
    }

    $btnGrid = New-ViewButton '▦' 0   # 位置は後でヘッダーリサイズ時に更新
    $btnList = New-ViewButton '☰' 0

    $header.Controls.AddRange(@($searchBox, $btnGrid, $btnList))
    $form.Controls.Add($header)

    # ------------------------------------------------------------------
    # コンテンツパネル (スクロール可能)
    # ------------------------------------------------------------------
    $content = New-Object System.Windows.Forms.Panel
    $content.Dock        = [System.Windows.Forms.DockStyle]::Fill
    $content.BackColor   = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $content.AutoScroll  = $true
    $form.Controls.Add($content)

    # ------------------------------------------------------------------
    # 現在のビュー状態
    # ------------------------------------------------------------------
    $script:viewMode = $global:Config.settings.defaultView  # 'grid' or 'list'

    function Update-ButtonStates {
        if ($script:viewMode -eq 'grid') {
            $btnGrid.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
            $btnList.BackColor = [System.Drawing.Color]::FromArgb(58, 58, 58)
        } else {
            $btnList.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
            $btnGrid.BackColor = [System.Drawing.Color]::FromArgb(58, 58, 58)
        }
    }

    # ------------------------------------------------------------------
    # グリッドビュー描画
    # ------------------------------------------------------------------
    function Show-GridView {
        param($AppList)

        $content.AutoScrollPosition = New-Object System.Drawing.Point(0, 0)
        $content.Controls.Clear()
        $iconSize  = 48
        $cellW     = $iconSize + 24
        $cellH     = $iconSize + 28
        $padLeft   = 8
        $padTop    = 8

        $cols = [Math]::Max(1, [Math]::Floor(($content.ClientSize.Width - $padLeft) / $cellW))

        $col = 0
        $row = 0

        foreach ($app in $AppList) {
            $appRef = $app  # クロージャ用にローカルコピー

            $cell = New-Object System.Windows.Forms.Panel
            $cell.Size      = New-Object System.Drawing.Size($cellW, $cellH)
            $cell.Location  = New-Object System.Drawing.Point(
                ($padLeft + $col * $cellW),
                ($padTop  + $row * $cellH)
            )
            $cell.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
            $cell.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $cell.Tag       = $appRef

            # アイコン
            $pic = New-Object System.Windows.Forms.PictureBox
            $pic.Size     = New-Object System.Drawing.Size($iconSize, $iconSize)
            $pic.Location = New-Object System.Drawing.Point(
                [int](($cellW - $iconSize) / 2), 2
            )
            $pic.SizeMode   = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
            $pic.BackColor  = [System.Drawing.Color]::Transparent
            $pic.Cursor     = [System.Windows.Forms.Cursors]::Hand
            $pic.Tag        = $appRef

            try {
                $pic.Image = Get-AppIcon -IconPath $appRef.iconPath -ExePath $appRef.path -Size $iconSize
            } catch {}

            # ラベル
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

            # ホバーカラー
            $hoverBg  = [System.Drawing.Color]::FromArgb(55, 55, 55)
            $normalBg = [System.Drawing.Color]::FromArgb(28, 28, 28)

            $enterSB = { $cell.BackColor = $hoverBg }.GetNewClosure()
            $leaveSB = { $cell.BackColor = $normalBg }.GetNewClosure()

            foreach ($ctrl in @($cell, $pic, $lbl)) {
                $ctrl.add_MouseEnter($enterSB)
                $ctrl.add_MouseLeave($leaveSB)
            }

            # クリックで起動
            $launchSB = {
                $data = $cell.Tag
                if ($null -eq $data -or [string]::IsNullOrEmpty($data.path)) { return }
                try {
                    Start-Process -FilePath $data.path
                    $global:LauncherForm.Hide()
                } catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        "起動に失敗しました: $($data.name)`n$_",
                        'PSLauncher',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) | Out-Null
                }
            }.GetNewClosure()

            foreach ($ctrl in @($cell, $pic, $lbl)) {
                $ctrl.add_Click($launchSB)
            }

            $cell.Controls.AddRange(@($pic, $lbl))
            $content.Controls.Add($cell)

            $col++
            if ($col -ge $cols) { $col = 0; $row++ }
        }
    }

    # ------------------------------------------------------------------
    # リストビュー描画
    # ------------------------------------------------------------------
    function Show-ListView {
        param($AppList)

        $content.Controls.Clear()

        $lv = New-Object System.Windows.Forms.ListView
        $lv.Dock        = [System.Windows.Forms.DockStyle]::Fill
        $lv.View        = [System.Windows.Forms.View]::Details
        $lv.BackColor   = [System.Drawing.Color]::FromArgb(28, 28, 28)
        $lv.ForeColor   = [System.Drawing.Color]::FromArgb(210, 210, 210)
        $lv.FullRowSelect = $true
        $lv.GridLines   = $false
        $lv.BorderStyle = [System.Windows.Forms.BorderStyle]::None
        $lv.Font        = New-Object System.Drawing.Font('Segoe UI', 10)
        $lv.HeaderStyle = [System.Windows.Forms.ColumnHeaderStyle]::None

        $lv.Columns.Add('アプリ名', 200) | Out-Null
        $lv.Columns.Add('グループ', 100) | Out-Null
        $lv.Columns.Add('パス',     350) | Out-Null

        # アイコンリスト
        $il = New-Object System.Windows.Forms.ImageList
        $il.ImageSize   = New-Object System.Drawing.Size(24, 24)
        $il.ColorDepth  = [System.Windows.Forms.ColorDepth]::Depth32Bit

        foreach ($app in $AppList) {
            try {
                $il.Images.Add((Get-AppIcon -IconPath $app.iconPath -ExePath $app.path -Size 24)) | Out-Null
            } catch {
                $bmp = New-Object System.Drawing.Bitmap(24, 24)
                $il.Images.Add($bmp) | Out-Null
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

        $lv.add_DoubleClick({
            if ($lv.SelectedItems.Count -gt 0) {
                $app = $lv.SelectedItems[0].Tag
                if ($null -eq $app -or [string]::IsNullOrEmpty($app.path)) { return }
                try {
                    Start-Process -FilePath $app.path
                    $global:LauncherForm.Hide()
                } catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        "起動に失敗しました: $($app.name)`n$_",
                        'PSLauncher',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) | Out-Null
                }
            }
        })

        $content.Controls.Add($lv)
    }

    # ------------------------------------------------------------------
    # アプリリスト更新
    # ------------------------------------------------------------------
    function Refresh-Apps {
        $searchText = if ([bool]$searchBox.Tag) { $searchBox.Text } else { '' }

        $apps = @($global:Config.apps)
        if ($searchText -ne '') {
            $apps = @($apps | Where-Object {
                ($_.name  -and $_.name  -like "*$searchText*") -or
                ($_.group -and $_.group -like "*$searchText*")
            })
        }

        if ($script:viewMode -eq 'grid') {
            Show-GridView $apps
        } else {
            Show-ListView $apps
        }
    }

    # ------------------------------------------------------------------
    # ヘッダーコントロール位置更新
    # ------------------------------------------------------------------
    function Update-HeaderLayout {
        $w = $form.ClientSize.Width
        $searchBox.Width = $w - 96
        $btnGrid.Location = New-Object System.Drawing.Point(($w - 80), 9)
        $btnList.Location = New-Object System.Drawing.Point(($w - 44), 9)
    }

    # ------------------------------------------------------------------
    # イベント接続
    # ------------------------------------------------------------------
    $searchBox.add_TextChanged({
        if ([bool]$searchBox.Tag) { Refresh-Apps }
    })

    $btnGrid.add_Click({
        $script:viewMode = 'grid'
        Update-ButtonStates
        Refresh-Apps
    })

    $btnList.add_Click({
        $script:viewMode = 'list'
        Update-ButtonStates
        Refresh-Apps
    })

    $form.add_Resize({
        Update-HeaderLayout
        if ($script:viewMode -eq 'grid') { Refresh-Apps }
    })

    # 表示時にアプリ一覧を更新
    $form.add_Shown({
        $global:Config = Import-Config
        Update-HeaderLayout
        Update-ButtonStates
        Refresh-Apps
        $searchBox.Focus()
    })

    return $form
}
