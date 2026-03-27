#Requires -Version 5.0
<#
.SYNOPSIS
    PSLauncher - タスクトレイ常駐型ランチャーアプリ
.DESCRIPTION
    Windows 11 上で動作するタスクトレイ常駐型ランチャー。
    ホットコーナー・ホットキー・トレイアイコンダブルクリックでランチャーを表示する。
.NOTES
    起動方法:
        powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "PSLauncher.ps1"
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------------------------------------
# スクリプトパス
# ---------------------------------------------------------------------------
$global:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$global:ConfigPath = Join-Path $global:ScriptDir 'config.json'

# サブスクリプト読み込み
. (Join-Path $global:ScriptDir 'LauncherWindow.ps1')
. (Join-Path $global:ScriptDir 'SettingsWindow.ps1')

# ---------------------------------------------------------------------------
# 多重起動防止 (Mutex)
# ---------------------------------------------------------------------------
$mutexName  = 'Global\PSLauncher_SingleInstance'
$global:AppMutex = New-Object System.Threading.Mutex($false, $mutexName)
if (-not $global:AppMutex.WaitOne(0)) {
    [System.Windows.Forms.MessageBox]::Show(
        'PSLauncher はすでに起動しています。',
        'PSLauncher',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
    exit 1
}

# ---------------------------------------------------------------------------
# Win32 API (RegisterHotKey / ホットキー受信フォーム)
# ---------------------------------------------------------------------------
if (-not ([System.Management.Automation.PSTypeName]'PSLauncher.HotkeyHelper').Type) {
    Add-Type -ReferencedAssemblies System.Windows.Forms,System.Drawing -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PSLauncher {
    public class HotkeyHelper : Form {
        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_ID   = 9001;

        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public event EventHandler HotkeyPressed;

        public HotkeyHelper() {
            this.ShowInTaskbar  = false;
            this.FormBorderStyle = FormBorderStyle.None;
            this.WindowState    = FormWindowState.Minimized;
            this.Opacity        = 0;
            this.Size           = new System.Drawing.Size(1, 1);
        }

        public bool Register(int modifiers, int vk) {
            // ハンドル生成を保証
            IntPtr dummy = this.Handle;
            return RegisterHotKey(this.Handle, HOTKEY_ID, modifiers, vk);
        }

        public void Unregister() {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
        }

        protected override void WndProc(ref Message m) {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID) {
                if (HotkeyPressed != null) { HotkeyPressed(this, EventArgs.Empty); }
            }
            base.WndProc(ref m);
        }

        protected override void SetVisibleCore(bool value) {
            // フォームを画面に表示しない
            base.SetVisibleCore(false);
        }
    }
}
'@
}

# ---------------------------------------------------------------------------
# 設定ユーティリティ
# ---------------------------------------------------------------------------
function Get-DefaultConfig {
    return [ordered]@{
        apps     = @()
        settings = [ordered]@{
            hotCorner  = [ordered]@{
                enabled    = $true
                corner     = 'bottomRight'  # topLeft / topRight / bottomLeft / bottomRight / disabled
                pixels     = 5
                cooldownMs = 1500
            }
            hotkey     = [ordered]@{
                enabled   = $true
                modifiers = 6     # Ctrl(2) + Shift(4)
                key       = 32    # VK_SPACE
                display   = 'Ctrl+Shift+Space'
            }
            defaultView   = 'grid'   # grid / list
            startup       = $false
            windowWidth   = 640
            windowHeight  = 480
        }
    }
}

function ConvertFrom-PSCustomObject {
    param([Parameter(ValueFromPipeline)]$InputObject)
    process {
        if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
            $ht = [ordered]@{}
            foreach ($prop in $InputObject.PSObject.Properties) {
                $ht[$prop.Name] = ConvertFrom-PSCustomObject $prop.Value
            }
            return $ht
        }
        elseif ($InputObject -is [System.Object[]] -or $InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
            return @($InputObject | ForEach-Object { ConvertFrom-PSCustomObject $_ })
        }
        return $InputObject
    }
}

function Import-Config {
    if (Test-Path $global:ConfigPath) {
        try {
            $raw    = Get-Content $global:ConfigPath -Raw -Encoding UTF8
            $loaded = ConvertFrom-PSCustomObject ($raw | ConvertFrom-Json)
            # apps が null の場合は空配列に正規化
            if ($null -eq $loaded.apps) { $loaded.apps = @() }
            return $loaded
        } catch {
            Write-Warning "config.json の読み込みに失敗しました: $_"
        }
    }
    return Get-DefaultConfig
}

function Export-Config {
    param($Config)
    try {
        $Config | ConvertTo-Json -Depth 10 | Set-Content $global:ConfigPath -Encoding UTF8
    } catch {
        Write-Warning "config.json の保存に失敗しました: $_"
    }
}

# ---------------------------------------------------------------------------
# グローバル状態
# ---------------------------------------------------------------------------
$global:Config        = Import-Config
$global:LauncherForm  = $null

# ---------------------------------------------------------------------------
# ランチャー表示 / 非表示トグル
# ---------------------------------------------------------------------------
function Show-Launcher {
    # フォームが破棄済みなら再生成
    if ($null -eq $global:LauncherForm -or $global:LauncherForm.IsDisposed) {
        $global:LauncherForm = New-LauncherWindow
    }

    if ($global:LauncherForm.Visible) {
        $global:LauncherForm.Hide()
        return
    }

    # マウス位置の近くに表示（画面外にはみ出さないよう調整）
    $mousePos  = [System.Windows.Forms.Cursor]::Position
    $screen    = [System.Windows.Forms.Screen]::FromPoint($mousePos)
    $wa        = $screen.WorkingArea
    $fw        = $global:LauncherForm.Width
    $fh        = $global:LauncherForm.Height

    $x = [Math]::Min($mousePos.X, $wa.Right  - $fw)
    $y = [Math]::Min($mousePos.Y, $wa.Bottom - $fh)
    $x = [Math]::Max($x, $wa.Left)
    $y = [Math]::Max($y, $wa.Top)

    $global:LauncherForm.Location = New-Object System.Drawing.Point($x, $y)
    $global:LauncherForm.Show()
    $global:LauncherForm.Activate()
}

# ---------------------------------------------------------------------------
# 設定ウィンドウを開く
# ---------------------------------------------------------------------------
function Open-Settings {
    $settingsForm = New-SettingsWindow
    $settingsForm.ShowDialog() | Out-Null
    $settingsForm.Dispose()

    # 設定変更を反映
    $global:Config = Import-Config
    Apply-HotCornerSetting
    Apply-HotkeySetting
}

# ---------------------------------------------------------------------------
# トレイアイコン生成
# ---------------------------------------------------------------------------
function New-TrayIcon {
    $bmp = New-Object System.Drawing.Bitmap(16, 16)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.FillRectangle(
        (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 120, 215))),
        0, 0, 16, 16
    )
    $g.DrawString(
        'L',
        (New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)),
        [System.Drawing.Brushes]::White,
        1, 1
    )
    $g.Dispose()
    return [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
}

$trayIcon      = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon = New-TrayIcon
$trayIcon.Text = 'PSLauncher'

# コンテキストメニュー
$cms = New-Object System.Windows.Forms.ContextMenuStrip
$cms.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$cms.ForeColor = [System.Drawing.Color]::White

function Add-MenuItem {
    param($Strip, $Text, $Action)
    $item = New-Object System.Windows.Forms.ToolStripMenuItem
    $item.Text      = $Text
    $item.ForeColor = [System.Drawing.Color]::White
    $item.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $item.add_Click($Action)
    $Strip.Items.Add($item) | Out-Null
    return $item
}

Add-MenuItem $cms 'ランチャーを開く' { Show-Launcher }   | Out-Null
Add-MenuItem $cms '設定'             { Open-Settings }    | Out-Null
$cms.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
Add-MenuItem $cms '終了' {
    $global:HotCornerTimer.Stop()
    $global:HotkeyHelper.Unregister()
    $trayIcon.Visible = $false
    $trayIcon.Dispose()
    [System.Windows.Forms.Application]::Exit()
    try { $global:AppMutex.ReleaseMutex() } catch {}
} | Out-Null

$trayIcon.ContextMenuStrip = $cms
$trayIcon.add_DoubleClick({ Show-Launcher })
$trayIcon.Visible = $true

# ---------------------------------------------------------------------------
# ホットキー
# ---------------------------------------------------------------------------
$global:HotkeyHelper = New-Object PSLauncher.HotkeyHelper
$global:HotkeyHelper.add_HotkeyPressed({ Show-Launcher })

function Apply-HotkeySetting {
    $global:HotkeyHelper.Unregister()
    $hk = $global:Config.settings.hotkey
    if ($hk.enabled) {
        $ok = $global:HotkeyHelper.Register([int]$hk.modifiers, [int]$hk.key)
        if (-not $ok) {
            Write-Warning "ホットキーの登録に失敗しました ($($hk.display))"
        }
    }
}

Apply-HotkeySetting

# ---------------------------------------------------------------------------
# ホットコーナー検出タイマー
# ---------------------------------------------------------------------------
$global:LastHotCornerFired = [DateTime]::MinValue

$global:HotCornerTimer          = New-Object System.Windows.Forms.Timer
$global:HotCornerTimer.Interval = 100  # 100ms ごとにポーリング

$global:HotCornerTimer.add_Tick({
    $hc = $global:Config.settings.hotCorner
    if (-not $hc.enabled -or $hc.corner -eq 'disabled') { return }

    $cooldown = [int]$hc.cooldownMs
    $pixels   = [int]$hc.pixels
    $now      = [DateTime]::Now

    if (($now - $global:LastHotCornerFired).TotalMilliseconds -lt $cooldown) { return }

    $pos    = [System.Windows.Forms.Cursor]::Position
    $screen = [System.Windows.Forms.Screen]::FromPoint($pos)
    $b      = $screen.Bounds

    $hit = switch ($hc.corner) {
        'topLeft'     { $pos.X -le ($b.Left   + $pixels) -and $pos.Y -le ($b.Top    + $pixels) }
        'topRight'    { $pos.X -ge ($b.Right  - $pixels) -and $pos.Y -le ($b.Top    + $pixels) }
        'bottomLeft'  { $pos.X -le ($b.Left   + $pixels) -and $pos.Y -ge ($b.Bottom - $pixels) }
        'bottomRight' { $pos.X -ge ($b.Right  - $pixels) -and $pos.Y -ge ($b.Bottom - $pixels) }
        default       { $false }
    }

    if ($hit) {
        $global:LastHotCornerFired = $now
        Show-Launcher
    }
})

function Apply-HotCornerSetting {
    $hc = $global:Config.settings.hotCorner
    if ($hc.enabled -and $hc.corner -ne 'disabled') {
        $global:HotCornerTimer.Start()
    } else {
        $global:HotCornerTimer.Stop()
    }
}

Apply-HotCornerSetting

# ---------------------------------------------------------------------------
# アプリケーションループ開始
# ---------------------------------------------------------------------------
$appContext = New-Object System.Windows.Forms.ApplicationContext
[System.Windows.Forms.Application]::Run($appContext)

# 終了後クリーンアップ
try { $global:AppMutex.ReleaseMutex() } catch {}
