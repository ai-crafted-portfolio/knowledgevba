# インストーラ

knowledgevba の VBA モジュール 26 本を、既存の `.xlsm` に **1 クリックで一括 import + 初期化**
するためのインストーラです。`.bat` に `.xlsm` をドラッグ & ドロップするだけで、
ダウンロード → CP932 サニタイズ → CRLF 正規化 → VBComponents.Import → SetupSheetsAndButtons
までを完走します。

---

## ダウンロード

以下 3 ファイルを同じフォルダに保存してください（Windows のエクスプローラから右クリック →
「名前を付けてリンク先を保存」）。

| ファイル | リンク |
| --- | --- |
| `Install-KnowledgevbaModules.bat` | [raw を保存](https://raw.githubusercontent.com/ai-crafted-portfolio/knowledgevba/main/installer/Install-KnowledgevbaModules.bat) |
| `Install-KnowledgevbaModules.ps1` | [raw を保存](https://raw.githubusercontent.com/ai-crafted-portfolio/knowledgevba/main/installer/Install-KnowledgevbaModules.ps1) |
| `README.md`（使い方の詳細） | [raw を保存](https://raw.githubusercontent.com/ai-crafted-portfolio/knowledgevba/main/installer/README.md) |

PowerShell 経由でまとめてダウンロードする場合:

```powershell
$base = "https://raw.githubusercontent.com/ai-crafted-portfolio/knowledgevba/main/installer"
$dst  = "$env:USERPROFILE\Downloads\knowledgevba_installer"
New-Item -ItemType Directory -Force -Path $dst | Out-Null
foreach ($f in 'Install-KnowledgevbaModules.bat','Install-KnowledgevbaModules.ps1','README.md') {
    Invoke-WebRequest "$base/$f" -OutFile "$dst\$f"
}
explorer $dst
```

---

## 動作環境

- Windows 10 / 11
- Microsoft Excel（VBA 有効化）
- PowerShell 5.1 以上（Windows 標準）

---

## 事前設定（一度だけ必須）

PowerShell から `VBProject.VBComponents.Import` を呼ぶには、Excel の以下のオプションが ON
になっている必要があります。OFF だとインストーラが「access to the VBA Project is not trusted」
で停止します。

1. Excel を起動 → **ファイル** → **オプション**
2. **トラスト センター** → **トラスト センターの設定**
3. **マクロの設定** → **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をチェック**
4. **OK** で閉じる

この設定は Excel のユーザプロファイル単位で保存されます（マシン単位ではない）。
一度設定すれば次回以降は自動で有効です。

---

## 使い方

### 方法 1: ドラッグ & ドロップ（推奨）

1. インストール対象の `.xlsm` を **`Install-KnowledgevbaModules.bat` の上にドラッグ & ドロップ**
2. PowerShell が起動して 26 モジュールをダウンロード + import + 初期化
3. 完了したら `pause` で待機（Enter で閉じる）
4. 対象 `.xlsm` を再度開いて、シートとボタンが生成されていることを確認

### 方法 2: コマンドプロンプトから

```cmd
cd %USERPROFILE%\Downloads\knowledgevba_installer
Install-KnowledgevbaModules.bat "C:\path\to\your_book.xlsm"
```

### 方法 3: Cowork デモモード

```cmd
Install-KnowledgevbaModules.bat "C:\path\to\demo_book.xlsm" /demo
```

`/demo` フラグを付けると、ThisWorkbook が **OnTime auto-init 版** に置き換わります
（コンテンツ有効化の 1〜2 秒後にシート + ボタン + デモナレッジが自動投入される）。
省略時は **本番版**（auto-init なし、起動時にユーザが手動で `SetupSheetsAndButtons` を実行する想定）。

---

## 何をしているか

PowerShell スクリプトは内部で以下を実行します。

1. `https://raw.githubusercontent.com/ai-crafted-portfolio/knowledgevba/main/docs/source/<layer>/<name>.md`
   から各モジュール .md を取得
2. 各 .md の ` ```vba ... ``` ` コードブロックを抽出
3. .bas / .cls ヘッダ（`Attribute VB_Name = "..."` / `VERSION 1.0 CLASS` 等）を補完
4. **Sanitize-ForCp932 関数で CP932 非互換文字（`▶` `²` `≤` `–` 等）を `?` に置換**
5. **改行を `\r\n`（CRLF）に正規化**（VBE Import が LF を受理しないため）
6. **Shift_JIS（CP932）エンコード**で `%TEMP%\knowledgevba_modules\` に保存
7. Excel COM で対象 .xlsm を開き `VBProject.VBComponents.Import()` で投入
8. ThisWorkbook は `CodeModule.AddFromString()` で本体だけ差し替え（Import 不可のため）
9. `Application.Run("SetupSheetsAndButtons", false)` で 14 シート + 29 ボタンを生成
10. 保存して Excel を quit、作業フォルダを清掃

---

## トラブルシューティング

### Q1. 「access to the VBA Project is not trusted」エラー

事前設定の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が OFF です。
Excel オプションで ON にしてください（上記 [事前設定](#一度だけ必須) を参照）。

### Q2. 「Compile error: invalid procedure call」または「ステートメントの最後」が出る

サイトの md ソースに CP932 非互換文字が残っている、または改行が LF のままになっている可能性。
本スクリプトは Sanitize-ForCp932 と CRLF 正規化を import 前に実施しますが、サニタイズ後も残った
コンパイルエラーがあれば該当行をスクショ + モジュール名を連絡してください
（2026-05-08 の cp932 commits で網羅修正済の想定です）。

### Q3. 26 モジュール全部 import されない

bat ウィンドウのコンソール出力を見ると `+ imported XXX` のログが出ます。
出ていないモジュールは:

- ダウンロード失敗（サイト 404 / ネットワーク）
- 該当 .md にコードブロックが無い（フロントマター修正待ち）

のどちらか。出力で `[WARN]` 行があるはずなのでそれを確認してください。

### Q4. 既存ナレッジデータが消える？

**消えません**。このスクリプトは VBA モジュールの差替えだけで、シート上のデータは触りません。
ただし `SetupSheetsAndButtons` は不足しているシート + ボタンを追加するので、
**新規シートが増える可能性**はあります（既存シートは保持）。

### Q5. もう一度実行するとどうなる？

何度実行しても OK（idempotent）。同名モジュールは削除 → 再 import で常に最新版に揃います。
ThisWorkbook も DeleteLines + AddFromString で完全置換されます。

---

## 関連 ADR

- ADR 0008: VBA 配布パターン（モジュール手動 import + Setup マクロ）
- ADR 0026: real Excel COM ビルド経路（aspose binary stub の罠回避）
- ADR 0027: クラスモジュール内 Public Const 禁止
- ADR 0028: Function/Sub Exit 文の整合性
- ADR 0035: VBA source の CP932 互換性（ユーザ入力 SJIS 整合）

このインストーラ自体の方針は実質 ADR-0008 + 0026 + 0035 の組合せです。
