# PSLauncher

Windows 11 上で動作するタスクトレイ常駐型ランチャーアプリです。
PowerShell + Windows Forms のみで実装しており、外部ツールや exe のインストールは不要です。

---

## ファイル構成

```
PSLauncher/
├── PSLauncher.ps1       # エントリポイント
├── LauncherWindow.ps1   # ランチャーウィンドウ（グリッド／リスト表示）
├── SettingsWindow.ps1   # 設定 GUI
├── config.json          # 設定データ（自動生成・編集可）
└── README.md            # 本ファイル
```

---

## 動作要件

| 項目 | 要件 |
|------|------|
| OS | Windows 10 / 11 |
| PowerShell | 5.0 以上（Windows 11 標準搭載） |
| 外部ツール | 不要（Windows 標準機能のみ使用） |

---

## 起動方法

### 通常起動（コンソールウィンドウあり）

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\path\to\PSLauncher\PSLauncher.ps1"
```

### コンソール非表示で起動（推奨）

```powershell
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\path\to\PSLauncher\PSLauncher.ps1"
```

---

## スタートアップへの登録

2 通りの方法があります。

### 方法 A: 設定画面から登録（推奨）

1. タスクトレイの「L」アイコンを右クリック → **設定** を選択
2. **設定** タブを開く
3. 「Windows 起動時に自動起動する」チェックボックスをオンにして **保存**

スタートアップフォルダ (`shell:startup`) に `PSLauncher.lnk` が自動作成されます。

### 方法 B: 手動でショートカットを作成

1. `Win + R` → `shell:startup` を入力して Enter
2. 開いたフォルダに新しいショートカットを作成
   - **対象**: `powershell.exe`
   - **引数**: `-ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\path\to\PSLauncher\PSLauncher.ps1"`
3. ショートカットのプロパティ → **実行時の大きさ** を **最小化** に設定

---

## 機能一覧

### ランチャー起動トリガー（3 種類）

| トリガー | 内容 |
|----------|------|
| **ホットコーナー** | 画面の指定した隅にマウスを移動するとランチャーが開く |
| **ホットキー** | デフォルト `Ctrl + Shift + Space`（設定で変更可） |
| **トレイアイコン** | タスクトレイの「L」アイコンをダブルクリック |

### ランチャーウィンドウ

- **グリッド表示**と**リスト表示**をワンクリックで切り替え
- 検索ボックスによるインクリメンタル絞り込み
- フォーカスが外れると自動的に非表示

### 設定画面（タスクトレイ右クリック → 設定）

- アプリの**追加・編集・削除・並び替え**
- ホットコーナーの有効化・コーナー位置・判定ピクセル数・クールダウン時間
- ホットキーの有効化・キー組み合わせ選択
- 既定のビュー（グリッド / リスト）
- スタートアップ自動起動 ON/OFF

---

## 各アプリエントリの設定項目

| 項目 | 説明 |
|------|------|
| 表示名 | ランチャーに表示される名前（必須） |
| 実行パス | `exe` / `cmd` / `bat` など（必須） |
| アイコン | 画像ファイルのパス（省略時は実行ファイルから自動取得） |
| グループ | 任意のカテゴリ文字列（検索対象になる） |

---

## config.json について

設定画面から操作すると自動更新されます。手動で編集する場合は UTF-8 で保存してください。

```json
{
  "apps": [
    {
      "name": "メモ帳",
      "path": "C:\\Windows\\System32\\notepad.exe",
      "iconPath": "",
      "group": "ユーティリティ"
    }
  ],
  "settings": {
    "hotCorner": {
      "enabled": true,
      "corner": "bottomRight",
      "pixels": 5,
      "cooldownMs": 1500
    },
    "hotkey": {
      "enabled": true,
      "modifiers": 6,
      "key": 32,
      "display": "Ctrl+Shift+Space"
    },
    "defaultView": "grid",
    "startup": false,
    "windowWidth": 640,
    "windowHeight": 480
  }
}
```

`corner` の値: `topLeft` / `topRight` / `bottomLeft` / `bottomRight` / `disabled`

`modifiers` の値 (ビット OR):
- `MOD_ALT = 1`
- `MOD_CONTROL = 2`
- `MOD_SHIFT = 4`
- `MOD_WIN = 8`

---

## 注意事項

- ホットキー `Ctrl+Shift+Space` は他のアプリと競合する場合があります。競合時は設定画面で別のキーに変更してください。
- `ExecutionPolicy` が `Restricted` の環境では `Bypass` または `RemoteSigned` に変更が必要です。
- アイコン取得に失敗した場合は実行ファイル名の頭文字が表示されます。
