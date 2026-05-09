# knowledgevba VBA モジュール自動インポータ

GitHub Pages サイト（https://ai-crafted-portfolio.github.io/knowledgevba/）から VBA モジュール
26 本を自動ダウンロード → 既存 .xlsm に Excel COM 経由で一括 import → SetupSheetsAndButtons 実行
までを 1 クリックで完了させるツールです。

---

## 動作環境

- Windows 10/11
- Microsoft Excel（VBA 有効化）
- PowerShell 5.1 以上（Windows 標準）

---

## ★ 事前設定（一度だけ必須）

Excel を起動し、以下を ON にしてください。OFF だと PowerShell から `VBProject.VBComponents.Import`
が呼べず、自動 import が失敗します。

1. Excel → ファイル → オプション
2. トラスト センター → トラスト センターの設定
3. マクロの設定 → **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をチェック**
4. OK で閉じる

この設定はマシン単位ではなく Excel ユーザプロファイル単位で保存されます（一度設定すれば次回以降は自動）。

---

## 使い方

### 方法 1: ドラッグ & ドロップ（推奨）

1. **対象 .xlsm を `Install-KnowledgevbaModules.bat` の上にドラッグ & ドロップ**
2. PowerShell が起動して 26 モジュールをダウンロード + import + 初期化
3. 完了したら `pause` で待機（Enter で閉じる）
4. 対象 .xlsm を再度開いて確認

### 方法 2: コマンドプロンプトから

```cmd
cd C:\path\to\knowledgevba_installer
Install-KnowledgevbaModules.bat "C:\path\to\your_book.xlsm"
```

### 方法 3: Cowork デモモード（modAutoInit + modDemoSeeder 含む）

```cmd
Install-KnowledgevbaModules.bat "C:\path\to\demo_book.xlsm" /demo
```

`/demo` フラグを付けると ThisWorkbook が **OnTime auto-init 版**（コンテンツ有効化後 1〜2 秒で
シート + ボタン + デモナレッジが自動投入される版）に置き換わります。

省略時は **本番版**（auto-init なし、起動時にユーザが Alt+F8 で SetupSheetsAndButtons を実行する想定）。

---

## 何をしているか（中身）

PowerShell スクリプトの動作:

1. `https://raw.githubusercontent.com/ai-crafted-portfolio/knowledgevba/main/docs/source/<layer>/<name>.md`
   から各モジュール .md を取得
2. 各 .md の ` ```vba ... ``` ` コードブロックを抽出
3. .bas / .cls ヘッダ（`Attribute VB_Name = "..."` / `VERSION 1.0 CLASS` 等）を補完
4. **Sanitize-ForCp932 関数で CP932 非互換文字を `?` に置換**
   - サイトの md ソースに `▶` `²` `≤` `–` 等が混入していた場合でも、Excel VBE が読める Shift_JIS に
     落とし込めるよう、import 直前に強制サニタイズします（ADR 0035）
5. **改行を CRLF に正規化**
   - `LF` のままだと VBE Import で「ステートメントの最後」コンパイルエラーが発生するため、
     `\r?\n` を `\r\n` に統一してから保存します
6. **Shift_JIS（CP932）エンコード**で `%TEMP%\knowledgevba_modules\` に保存
   - 理由: ユーザ入力が SJIS 環境（Excel 日本語ロケール）の場合、ソースも SJIS で保存されないと
     文字化けが起きる（ADR 0035）
7. Excel COM で対象 .xlsm を開き、`VBProject.VBComponents.Import()` で各モジュール投入
8. ThisWorkbook は `Import` 不可能なので `CodeModule.AddFromString()` で本体だけ差し替え
9. `Application.Run("SetupSheetsAndButtons", false)` で 14 シート + 29 ボタンを自動生成
10. 保存して Excel を quit
11. `%TEMP%\knowledgevba_modules\` を清掃（`-KeepWorkDir` 付ければ残せる）

---

## トラブルシューティング

### Q1. 「access to the VBA Project is not trusted」エラー

★ 事前設定の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が OFF です。
Excel オプションで ON にしてください。

### Q2. 「Compile error: invalid procedure call」または「ステートメントの最後」が出る

サイトの一部モジュールに `▶` `²` `≤` などの CP932 非互換文字が残っている、または改行が LF の
ままになっている可能性。本スクリプトは Sanitize-ForCp932 と CRLF 正規化を import 前に実施
しますが、サニタイズ後も残ったコンパイルエラーがあれば該当行をスクショ + モジュール名を
連絡してください（2026-05-08 の cp932 commits で網羅修正済の想定）。

### Q3. 26 モジュール全部 import されない

bat ウィンドウのコンソール出力を見ると、`+ imported XXX` のログが出ます。
出ていないモジュールは:
- ダウンロード失敗（サイト 404 / ネットワーク）
- 該当 .md にコードブロックが無い（フロントマター修正待ち）
のどちらか。出力で `[WARN]` 行があるはずなのでそれを見てください。

### Q4. 既存ナレッジデータが消える？

**消えません**。このスクリプトは VBA モジュールの差替えだけで、シート上のデータは触りません。
ただし `SetupSheetsAndButtons` は不足しているシート + ボタンを追加するので、
**新規シートが増える可能性**はあります（既存シートは保持）。

### Q5. もう一度実行するとどうなる？

何度実行しても OK（idempotent）。同名モジュールは削除 → 再 import で常に最新版に揃います。
ThisWorkbook も DeleteLines + AddFromString で完全置換。

---

## ファイル構成

```
installer/
  ├─ Install-KnowledgevbaModules.bat   ← これをダブルクリック / xlsm を D&D
  ├─ Install-KnowledgevbaModules.ps1   ← 本体スクリプト
  └─ README.md                         ← このファイル
```

---

## 関連 ADR

サイトの `https://ai-crafted-portfolio.github.io/knowledgevba/` および
`C:\decisions\adr\` 配下:

- ADR 0008: VBA 配布パターン（モジュール手動 import + Setup マクロ）
- ADR 0026: real Excel COM ビルド経路（aspose binary stub の罠回避）
- ADR 0027: クラスモジュール内 Public Const 禁止
- ADR 0028: Function/Sub Exit 文の整合性
- ADR 0035: VBA source の CP932 互換性（ユーザ入力 SJIS 整合）

このインストーラ自体の方針も実質 ADR-0008 + 0026 + 0035 の組合せです。
