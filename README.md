# excel-date-addin

設定した数字と `yymmdd`(今日の日付)を連結し、選択中のセルに**文字列**として出力する Excel Office.js アドインです。

例: 設定値が `1234`、日付が 2026-10-17 の場合 → `1234261017`

## 機能

- タスクペインで任意の数字を保存(`Office.context.roamingSettings`)
- 「出力」ボタンでアクティブセルに `設定値 + yymmdd` を出力
- 区切り文字なし、書式は文字列(`@`)で先頭の0も保持

## ファイル構成

| ファイル | 説明 |
|---|---|
| `manifest.xml` | アドインのマニフェスト |
| `taskpane.html` | タスクペインのUI |
| `taskpane.js` | 設定保存と出力のロジック |
| `assets/` | アイコン(`icon-16.png` / `icon-32.png` / `icon-64.png` / `icon-80.png`)※要追加 |

## セットアップ

1. **アイコンを追加**: `assets/` フォルダに上記のPNGを配置します。
2. **GitHub Pages を有効化**: `Settings` → `Pages` → Source を `Deploy from a branch`、Branch を `main` / `(root)` にします。
3. 公開URL `https://mr2okama.github.io/excel-date-addin/` で `taskpane.html` が開けることを確認します。
4. 別のURLで公開する場合は、`manifest.xml` 内のURLをすべて置換してください。

## Excel への読み込み(サイドロード)

- **Excel on the web**: `挿入` → `アドイン` → `個人用アドイン` → `マニフェストのアップロード` で `manifest.xml` を選択
- **Windows**: `manifest.xml` を共有フォルダに置き、信頼されたアドインカタログに登録して `挿入` → `個人用アドイン` → `共有フォルダー` から読み込み

## 使い方

1. ホームタブの「番号出力」でタスクペインを開く
2. 数字を入力して「設定を保存」
3. 出力したいセルを選択して「出力」

## 注意

- 設定値は数字のみ受け付けます。
- 複数セルを選択している場合は、アクティブな1セルにのみ出力されます。
- `manifest.xml` の `<Id>` は固有のGUIDです。複製して使う場合は新しいGUIDに変更してください。
