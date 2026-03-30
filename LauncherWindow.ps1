#
# LauncherWindow.ps1 - ランチャーウィンドウ（グリッド表示 / リスト表示）
#
# 【スコープ設計】
#   PowerShell の GetNewClosure() は「外側クロージャでキャプチャされた変数」を
#   内側の GetNewClosure() で再キャプチャできない多重クロージャ問題がある。
#   対策:
#     - コントロール参照など New-LauncherWindow スコープの共有状態は $s ハッシュテーブル
#     - showGridView 内部のイベントハンドラから参照が必要な状態は $script: スコープに保持
#       $script:LauncherDragStartPos / DragSourceApp / DragSourceCell
#       $script:LauncherLastShownApps
#       $script:LauncherRefreshApps
#

# ---------------------------------------------------------------------------
# ヘルパー: 16進数カラー文字列 → Color
# ---------------------------------------------------------------------------
function ConvertFrom-HexColor {
    param([string]$Hex)
    $Hex = $Hex.TrimStart('#')
    if ($Hex.Length -ne 6) { return [System.Drawing.Color]::Gray }
    try {
        return [System.Drawing.Color]::FromArgb(
            [Convert]::ToInt32($Hex.Substring(0,2), 16),
            [Convert]::ToInt32($Hex.Substring(2,2), 16),
            [Convert]::ToInt32($Hex.Substring(4,2), 16)
        )
    } catch { return [System.Drawing.Color]::Gray }
}

# ---------------------------------------------------------------------------
# ヘルパー: アイコン取得（URL ファビコン対応）
# ---------------------------------------------------------------------------
function Get-AppIcon {
    param([string]$IconPath, [string]$ExePath, [int]$Size = 48)

    # 明示アイコンパス
    if ($IconPath -and (Test-Path $IconPath)) {
        try {
            $img = [System.Drawing.Image]::FromFile($IconPath)
            return New-Object System.Drawing.Bitmap($img, $Size, $Size)
        } catch {}
    }

    # URL の場合: ファビコン自動取得
    if ($ExePath -match '^https?://') {
        try {
            $uri        = [System.Uri]$ExePath
            $faviconUrl = "$($uri.Scheme)://$($uri.Host)/favicon.ico"
            $req = [System.Net.WebRequest]::Create($faviconUrl)
            $req.Timeout = 3000
            $req.UserAgent = 'PSLauncher/1.0'
            $resp = $req.GetResponse()
            if ($resp.StatusCode -eq [System.Net.HttpStatusCode]::OK) {
                $stream = $resp.GetResponseStream()
                $img    = [System.Drawing.Image]::FromStream($stream)
                $bmp    = New-Object System.Drawing.Bitmap($img, $Size, $Size)
                $stream.Close(); $resp.Close()
                return $bmp
            }
            $resp.Close()
        } catch {}

        # ファビコン取得失敗時: リンクアイコン（青地に "@"）
        $bmp = New-Object System.Drawing.Bitmap($Size, $Size)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0,100,200))),0,0,$Size,$Size)
        $font = New-Object System.Drawing.Font('Segoe UI',[int]($Size*0.45),[System.Drawing.FontStyle]::Bold)
        $sf   = New-Object System.Drawing.StringFormat
        $sf.Alignment = $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $g.DrawString('@',$font,[System.Drawing.Brushes]::White,(New-Object System.Drawing.RectangleF(0,0,$Size,$Size)),$sf)
        $g.Dispose()
        return $bmp
    }

    # 実行ファイルパス: 関連アイコン抽出
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

    # フォールバック: イニシャルアイコン
    $bmp  = New-Object System.Drawing.Bitmap($Size, $Size)
    $g    = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0,120,215))),0,0,$Size,$Size)
    $initial = if ($ExePath) { [System.IO.Path]::GetFileNameWithoutExtension($ExePath) } else { '?' }
    $initial = $initial.Substring(0,1).ToUpper()
    $font    = New-Object System.Drawing.Font('Segoe UI',[int]($Size*0.45),[System.Drawing.FontStyle]::Bold)
    $sf      = New-Object System.Drawing.StringFormat
    $sf.Alignment = $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString($initial,$font,[System.Drawing.Brushes]::White,(New-Object System.Drawing.RectangleF(0,0,$Size,$Size)),$sf)
    $g.Dispose()
    return $bmp
}

# ---------------------------------------------------------------------------
# ランチャーウィンドウ生成
# ---------------------------------------------------------------------------
function New-LauncherWindow {

    # $script: 変数を初期化（多重クロージャ回避用）
    $script:LauncherDragStartPos   = $null
    $script:LauncherDragSourceApp  = $null
    $script:LauncherDragSourceCell = $null
    $script:LauncherLastShownApps  = @()
    $script:LauncherRefreshApps    = $null   # 後で代入

    # ----------------------------------------------------------------
    # テーマ色の読み込み
    # ----------------------------------------------------------------
    $th      = if ($global:Config.settings.theme) { $global:Config.settings.theme } else { @{} }
    $cBg     = ConvertFrom-HexColor (if ($th.bg)     { $th.bg }     else { '#1C1C1C' })
    $cHeader = ConvertFrom-HexColor (if ($th.header) { $th.header } else { '#282828' })
    $cTabBg  = ConvertFrom-HexColor (if ($th.tabBg)  { $th.tabBg }  else { '#232323' })
    $cAccent = ConvertFrom-HexColor (if ($th.accent) { $th.accent } else { '#0078D7' })
    $cText   = ConvertFrom-HexColor (if ($th.text)   { $th.text }   else { '#D2D2D2' })
    $cHover  = ConvertFrom-HexColor (if ($th.hover)  { $th.hover }  else { '#373737' })
    $cInput  = ConvertFrom-HexColor (if ($th.input)  { $th.input }  else { '#3A3A3A' })

    # ----------------------------------------------------------------
    # フォーム
    # ----------------------------------------------------------------
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'PSLauncher'
    $form.Width           = [int]$global:Config.settings.windowWidth
    $form.Height          = [int]$global:Config.settings.windowHeight
    $form.MinimumSize     = New-Object System.Drawing.Size(300, 250)
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::Manual
    $form.BackColor       = $cBg
    $form.ForeColor       = $cText
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.ShowInTaskbar   = $false
    $form.TopMost         = $true

    # ----------------------------------------------------------------
    # 外側コンテナ（Dock=Fill で form を埋める）
    # ----------------------------------------------------------------
    $outer = New-Object System.Windows.Forms.Panel
    $outer.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $outer.BackColor = $cBg
    $form.Controls.Add($outer)

    # ---- ヘッダー (y=0, h=44) ----
    $header = New-Object System.Windows.Forms.Panel
    $header.BackColor = $cHeader

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location    = New-Object System.Drawing.Point(8,10)
    $searchBox.Height      = 24
    $searchBox.BackColor   = $cInput
    $searchBox.ForeColor   = [System.Drawing.Color]::FromArgb(160,160,160)
    $searchBox.Font        = New-Object System.Drawing.Font('Segoe UI',10)
    $searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $searchBox.Text        = '検索...'
    $searchBox.Tag         = $false

    $btnGrid = New-Object System.Windows.Forms.Button
    $btnGrid.Text = '▦'; $btnGrid.Size = New-Object System.Drawing.Size(32,26)
    $btnGrid.Location = New-Object System.Drawing.Point(0,9)
    $btnGrid.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGrid.FlatAppearance.BorderSize = 0
    $btnGrid.ForeColor = $cText
    $btnGrid.Font = New-Object System.Drawing.Font('Segoe UI',10)

    $btnList = New-Object System.Windows.Forms.Button
    $btnList.Text = '☰'; $btnList.Size = New-Object System.Drawing.Size(32,26)
    $btnList.Location = New-Object System.Drawing.Point(0,9)
    $btnList.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnList.FlatAppearance.BorderSize = 0
    $btnList.ForeColor = $cText
    $btnList.Font = New-Object System.Drawing.Font('Segoe UI',10)

    $header.Controls.AddRange(@($searchBox,$btnGrid,$btnList))
    $outer.Controls.Add($header)

    # ---- タブパネル (y=44, h=34) ----
    $tabPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $tabPanel.BackColor     = $cTabBg
    $tabPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $tabPanel.WrapContents  = $false
    $tabPanel.Padding       = New-Object System.Windows.Forms.Padding(4,4,4,0)
    $outer.Controls.Add($tabPanel)

    # ---- コンテンツ (y=78, h=残り) — FlowLayoutPanel で自動折り返し配置 ----
    $content = New-Object System.Windows.Forms.FlowLayoutPanel
    $content.BackColor     = $cBg
    $content.AutoScroll    = $true
    $content.WrapContents  = $true
    $content.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $content.Padding       = New-Object System.Windows.Forms.Padding(6,6,0,0)
    $outer.Controls.Add($content)

    # ----------------------------------------------------------------
    # 共有状態 $s
    # ----------------------------------------------------------------
    # Color 構造体は $s ハッシュテーブルに格納すると null になる PowerShell の挙動があるため
    # $cBg 等はローカル変数のままにし、各クロージャが GetNewClosure() で直接キャプチャする
    $s = @{
        form       = $form
        outer      = $outer
        header     = $header
        searchBox  = $searchBox
        btnGrid    = $btnGrid
        btnList    = $btnList
        tabPanel   = $tabPanel
        content    = $content
        viewMode   = [string]$global:Config.settings.defaultView
        currentTab = 'すべて'
    }

    # ----------------------------------------------------------------
    # updateButtonStates
    # ----------------------------------------------------------------
    $s.updateButtonStates = {
        if ($s.viewMode -eq 'grid') {
            $s.btnGrid.BackColor = $cAccent
            $s.btnList.BackColor = $cInput
        } else {
            $s.btnList.BackColor = $cAccent
            $s.btnGrid.BackColor = $cInput
        }
    }.GetNewClosure()

    # ----------------------------------------------------------------
    # updateLayout — ヘッダー/タブ/コンテンツを手動配置
    # ----------------------------------------------------------------
    $s.updateLayout = {
        $w = $s.form.ClientSize.Width
        $h = $s.form.ClientSize.Height
        $s.header.SetBounds(0, 0,  $w, 44)
        $s.tabPanel.SetBounds(0, 44, $w, 34)
        $s.content.SetBounds(0, 78, $w, [Math]::Max(0, $h - 78))
        $s.searchBox.Width  = $w - 96
        $s.btnGrid.Location = New-Object System.Drawing.Point(($w-80),9)
        $s.btnList.Location = New-Object System.Drawing.Point(($w-44),9)
    }.GetNewClosure()

    # ----------------------------------------------------------------
    # tabClickHandler — タブボタン共通クリックハンドラ（第1レベルクロージャ）
    # ----------------------------------------------------------------
    $s.tabClickHandler = {
        $s.currentTab = ([System.Windows.Forms.Button]$args[0]).Text
        & $s.updateTabPanel
        & $s.updateLayout
        & $s.refreshApps
    }.GetNewClosure()

    # ----------------------------------------------------------------
    # updateTabPanel — グループからタブボタンを生成
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
            $btn.Font      = New-Object System.Drawing.Font('Segoe UI',9)
            $btn.ForeColor = $cText
            $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
            if ($isActive) {
                $btn.BackColor = $cAccent
                $btn.FlatAppearance.BorderColor = $cAccent
            } else {
                $btn.BackColor = $cInput
                $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80,80,80)
            }

            $btn.add_Click($s.tabClickHandler)
            $s.tabPanel.Controls.Add($btn)
        }
    }.GetNewClosure()

    # ----------------------------------------------------------------
    # showGridView — ドラッグ&ドロップ並び替え付き
    # ----------------------------------------------------------------
    $s.showGridView = {
        param($AppList)

        $s.content.Padding = New-Object System.Windows.Forms.Padding(6,6,0,0)
        $s.content.SuspendLayout()
        $s.content.Controls.Clear()
        $s.content.AutoScrollPosition = New-Object System.Drawing.Point(0,0)
        $script:LauncherLastShownApps = $AppList

        $iconSize = 48
        $cellW    = $iconSize + 24   # 72
        $cellH    = $iconSize + 28   # 76

        foreach ($app in $AppList) {
            $appRef = $app

            $cell = New-Object System.Windows.Forms.Panel
            $cell.Size      = New-Object System.Drawing.Size($cellW,$cellH)
            $cell.Margin    = New-Object System.Windows.Forms.Padding(2,2,2,2)
            $cell.BackColor = $cBg
            $cell.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $cell.Tag       = $appRef
            $cell.AllowDrop = $true

            $pic = New-Object System.Windows.Forms.PictureBox
            $pic.Size      = New-Object System.Drawing.Size($iconSize,$iconSize)
            $pic.Location  = New-Object System.Drawing.Point([int](($cellW-$iconSize)/2),2)
            $pic.SizeMode  = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
            $pic.BackColor = [System.Drawing.Color]::Transparent
            $pic.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $pic.Tag       = $appRef
            $pic.AllowDrop = $true
            try { $pic.Image = Get-AppIcon -IconPath $appRef.iconPath -ExePath $appRef.path -Size $iconSize } catch {}

            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text      = $appRef.name
            $lbl.Size      = New-Object System.Drawing.Size($cellW,22)
            $lbl.Location  = New-Object System.Drawing.Point(0,($iconSize+4))
            $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $lbl.ForeColor = $cText
            $lbl.Font      = New-Object System.Drawing.Font('Segoe UI',8)
            $lbl.BackColor = [System.Drawing.Color]::Transparent
            $lbl.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $lbl.Tag       = $appRef
            $lbl.AllowDrop = $true

            $hoverBg  = $cHover
            $normalBg = $cBg
            $dragBg   = [System.Drawing.Color]::FromArgb(0,100,190)

            # --- ホバー ---
            $enterSB = { $cell.BackColor = $hoverBg  }.GetNewClosure()
            $leaveSB = { $cell.BackColor = $normalBg }.GetNewClosure()

            # --- クリック起動（ドラッグ中は無視）---
            $launchSB = {
                if ($null -ne $script:LauncherDragSourceApp) { return }
                if ([string]::IsNullOrEmpty($appRef.path)) { return }
                try {
                    Start-Process $appRef.path
                    $global:LauncherForm.Hide()
                } catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        "起動に失敗しました: $($appRef.name)`n$_",'PSLauncher',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                }
            }.GetNewClosure()

            # --- ドラッグ開始検出 ---
            $mouseDownSB = {
                $e = $args[1]
                if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                    $script:LauncherDragStartPos   = New-Object System.Drawing.Point($e.X,$e.Y)
                    $script:LauncherDragSourceApp  = $null
                    $script:LauncherDragSourceCell = $cell
                }
            }.GetNewClosure()

            $mouseMoveSB = {
                if ($null -eq $script:LauncherDragStartPos) { return }
                $e = $args[1]
                $ds = [System.Windows.Forms.SystemInformation]::DragSize
                if ([Math]::Abs($e.X - $script:LauncherDragStartPos.X) -gt $ds.Width -or
                    [Math]::Abs($e.Y - $script:LauncherDragStartPos.Y) -gt $ds.Height) {
                    $script:LauncherDragStartPos  = $null
                    $script:LauncherDragSourceApp = $appRef
                    $cell.DoDragDrop($appRef,[System.Windows.Forms.DragDropEffects]::Move) | Out-Null
                    $script:LauncherDragSourceApp  = $null
                    $script:LauncherDragSourceCell = $null
                }
            }.GetNewClosure()

            $mouseUpSB = {
                $script:LauncherDragStartPos   = $null
                $script:LauncherDragSourceApp  = $null
                $script:LauncherDragSourceCell = $null
            }.GetNewClosure()

            # --- ドロップ先のハイライト ---
            $dragOverSB = {
                $args[1].Effect = [System.Windows.Forms.DragDropEffects]::Move
                if (-not [object]::ReferenceEquals($appRef, $script:LauncherDragSourceApp)) {
                    $cell.BackColor = $dragBg
                }
            }.GetNewClosure()

            $dragLeaveSB = {
                $cell.BackColor = $normalBg
            }.GetNewClosure()

            # --- ドロップ処理（確認 → 並び替え → 保存）---
            $dragDropSB = {
                $cell.BackColor = $normalBg
                $srcApp = $script:LauncherDragSourceApp
                $tgtApp = $appRef

                if ($null -eq $srcApp -or [object]::ReferenceEquals($srcApp,$tgtApp)) { return }

                $ans = [System.Windows.Forms.MessageBox]::Show(
                    "「$($srcApp.name)」を「$($tgtApp.name)」の前に移動しますか？",
                    'PSLauncher',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question)

                if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }

                $shown  = @($script:LauncherLastShownApps)
                $srcIdx = -1; $tgtIdx = -1
                for ($i = 0; $i -lt $shown.Length; $i++) {
                    if ([object]::ReferenceEquals($shown[$i],$srcApp)) { $srcIdx = $i }
                    if ([object]::ReferenceEquals($shown[$i],$tgtApp)) { $tgtIdx = $i }
                }
                if ($srcIdx -lt 0 -or $tgtIdx -lt 0) { return }

                $newShown = [System.Collections.ArrayList]@($shown)
                $newShown.RemoveAt($srcIdx)
                $insIdx = if ($srcIdx -lt $tgtIdx) { $tgtIdx - 1 } else { $tgtIdx }
                $newShown.Insert($insIdx,$srcApp)

                $allApps  = [System.Collections.ArrayList]@($global:Config.apps)
                $origIdxs = [System.Collections.ArrayList]::new()
                foreach ($sa in $shown) {
                    for ($j = 0; $j -lt $allApps.Count; $j++) {
                        if ([object]::ReferenceEquals($allApps[$j],$sa)) {
                            $origIdxs.Add($j) | Out-Null; break
                        }
                    }
                }
                for ($i = 0; $i -lt $origIdxs.Count; $i++) {
                    $allApps[$origIdxs[$i]] = $newShown[$i]
                }
                $global:Config.apps = @($allApps)
                Export-Config $global:Config
                & $script:LauncherRefreshApps
            }.GetNewClosure()

            foreach ($ctrl in @($cell,$pic,$lbl)) {
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

            $cell.Controls.AddRange(@($pic,$lbl))
            $s.content.Controls.Add($cell)
        }
        $s.content.ResumeLayout($true)
        $s.content.Refresh()
    }.GetNewClosure()

    # ----------------------------------------------------------------
    # showListView
    # ----------------------------------------------------------------
    $s.showListView = {
        param($AppList)

        # リストビュー用: パディングをゼロにして全領域を使う
        $s.content.Padding = New-Object System.Windows.Forms.Padding(0)
        $s.content.Controls.Clear()
        $script:LauncherLastShownApps = $AppList

        $lv = New-Object System.Windows.Forms.ListView
        # FlowLayoutPanel 内では Dock=Fill が効かないため明示的にサイズ指定
        $lv.Size          = New-Object System.Drawing.Size($s.content.ClientSize.Width, $s.content.ClientSize.Height)
        $lv.Margin        = New-Object System.Windows.Forms.Padding(0)
        $lv.View          = [System.Windows.Forms.View]::Details
        $lv.BackColor     = $cBg
        $lv.ForeColor     = $cText
        $lv.FullRowSelect = $true
        $lv.GridLines     = $false
        $lv.BorderStyle   = [System.Windows.Forms.BorderStyle]::None
        $lv.Font          = New-Object System.Drawing.Font('Segoe UI',10)
        $lv.HeaderStyle   = [System.Windows.Forms.ColumnHeaderStyle]::None

        $lv.Columns.Add('アプリ名',200) | Out-Null
        $lv.Columns.Add('グループ',100) | Out-Null
        $lv.Columns.Add('パス',    350) | Out-Null

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
            $item = New-Object System.Windows.Forms.ListViewItem($app.name,$idx)
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
                        "起動に失敗しました: $($app.name)`n$_",'PSLauncher',
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

    # $script: に公開して showGridView 内の内側クロージャから呼べるようにする
    $script:LauncherRefreshApps = $s.refreshApps

    # ----------------------------------------------------------------
    # イベント登録
    # ----------------------------------------------------------------
    $form.add_Deactivate({ $s.form.Hide() }.GetNewClosure())

    $form.add_ResizeEnd({
        $global:Config.settings.windowWidth  = $s.form.Width
        $global:Config.settings.windowHeight = $s.form.Height
        Export-Config $global:Config
        # リサイズ完了後に一度だけ再描画（グリッド折り返し・リスト列幅を正確に更新）
        if ($s.form.Visible) { & $s.refreshApps }
    }.GetNewClosure())

    $searchBox.add_Enter({
        if (-not [bool]$s.searchBox.Tag) {
            $s.searchBox.Text      = ''
            $s.searchBox.ForeColor = $cText
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
        if (-not $s.form.Visible) { return }
        & $s.updateLayout
        # リストビューの場合は ListView のサイズを直接更新（再描画コストを避けるため）
        if ($s.viewMode -eq 'list' -and $s.content.Controls.Count -gt 0) {
            $lvCtrl = $s.content.Controls[0]
            if ($lvCtrl -is [System.Windows.Forms.ListView]) {
                $lvCtrl.Size = New-Object System.Drawing.Size(
                    $s.content.ClientSize.Width,
                    $s.content.ClientSize.Height)
            }
        }
    }.GetNewClosure())

    # VisibleChanged: 表示されるたびにコンフィグ再読み込み・アイコン更新
    $form.add_VisibleChanged({
        if (-not $s.form.Visible) { return }
        $global:Config = Import-Config
        $validGroups = @('すべて') + @($global:Config.apps |
            Where-Object { $_.group } | ForEach-Object { $_.group } | Select-Object -Unique)
        if ($s.currentTab -notin $validGroups) { $s.currentTab = 'すべて' }
        & $s.updateLayout
        & $s.updateButtonStates
        & $s.updateTabPanel
        & $s.refreshApps
        $s.searchBox.Focus()
    }.GetNewClosure())

    return $form
}
