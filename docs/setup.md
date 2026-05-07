---
title: セットアップ
description: VBA モジュール手動 import、初回セットアップマクロ実行、マクロ有効化など、利用開始までの手順
tags:
  - excel
  - vba
  - setup
---

# セットアップ

このページでは、ナレッジ管理ツールを **新規利用開始するための準備手順** を集約しています。日常操作（検索・登録・編集など）は [操作手順](operations.md) を参照してください。

## 0. 前提環境

- Microsoft Excel 2019 / 2021 / 365（Windows 版推奨）
- 書込権限のあるローカルフォルダ（ネットワークドライブ直下は避ける）
- マクロ有効化が許可されているプロファイル

!!! warning "Excel for Mac の制約"
    `Forms.*` ProgID が無いため `Macro_ShowSearchResultPreview`（spec 駆動 UserForm）は動作しません。検索のみなら可。

---

## 1. 配布パッケージを使う場合（モジュール手動 import）

職場本番展開向け。配布物は `release_v2/` ZIP（VBA モジュール 24 個 + サンプル `kb_images/` + README）で `.xlsm` 本体は含みません。

### Step 1.1 マクロ有効化と新規 .xlsm 作成

1. Excel を起動
2. `[ファイル] → [新規] → [空白のブック]`
3. `[ファイル] → [名前を付けて保存]` で **必ず `.xlsm`（Excel マクロ有効ブック）** として保存
    - ファイルの種類で「**Excel マクロ有効ブック (`*.xlsm`)**」を選択
    - 任意のファイル名（例: `ナレッジ管理.xlsm`）

### Step 1.2 マクロ設定の確認

`[ファイル] → [オプション] → [トラスト センター] → [トラスト センターの設定] → [マクロの設定]` で次を確認・設定します。

- **マクロの設定**: 「警告して、すべてのマクロを無効にする」または「電子署名されたマクロを除き、すべてのマクロを無効にする」
- **VBA プロジェクト オブジェクト モデルへのアクセスを信頼する**: **ON**（詳細プレビュー UserForm 機能を使う場合に必須）

### Step 1.3 VBE を開いて全モジュールをインポート

1. `Alt + F11` で VBE（Visual Basic Editor）を開く
2. 左の「プロジェクト」ツリーから対象ブックを右クリック → `[ファイルのインポート...]`
3. 以下のフォルダ別ファイル群を順次インポート

=== "エントリポイント層 (6)"

    - `modEntryFormat.bas`
    - `modEntryKnowledge.bas`
    - `modEntryMain.bas`
    - `modEntrySearch.bas`
    - `modEntrySettings.bas`
    - `modSpecExamples.bas`

=== "ビジネスロジック層 (9)"

    - `clsControlSpec.cls`
    - `clsFieldMigrator.cls`
    - `clsFormSpec.cls`
    - `clsFormatManager.cls`
    - `clsKnowledgeManager.cls`
    - `clsSearchEngine.cls`
    - `clsStorageResolver.cls`
    - `clsTaskController.cls`
    - `modFormBuilder.bas`

=== "ユーティリティ層 (5)"

    - `modCommon.bas`
    - `modDateUtil.bas`
    - `modFileIO.bas`
    - `modImageRender.bas`
    - `modStringUtil.bas`

=== "ログ層 (2)"

    - `clsLogger.cls`
    - `clsLogEntry.cls`

=== "インストーラ層 (1)"

    - `modSetup.bas`

各モジュールの中身は [ソースコード](source/index.md) ページから直接コピペすることもできます（コピーボタン付き）。

### Step 1.4 ThisWorkbook の中身を貼り付け（インポートではない）

`ThisWorkbook.cls` だけは「インポート」ではなく **中身のコピペ** が必要です。VBE は新しい `ThisWorkbook` を作れず、既存の `ThisWorkbook` に上書きする必要があるためです。

1. VBE 左ペインで **`Microsoft Excel Objects` → `ThisWorkbook`** をダブルクリック
2. 右側のコードペインを開く（現状は空のはず）
3. [ThisWorkbook.cls (本番版)](source/special/thisworkbook.md) をブラウザで開いてコピーボタンを押す
4. 先頭から下記の **9 行のヘッダーを除いた**残り全部を選択してコピー

    ```text
    VERSION 1.0 CLASS
    BEGIN
      MultiUse = -1  'True
    END
    Attribute VB_Name = "ThisWorkbook"
    Attribute VB_GlobalNameSpace = False
    Attribute VB_Creatable = False
    Attribute VB_PredeclaredId = True
    Attribute VB_Exposed = True
    ```

5. VBE の `ThisWorkbook` コードペインに貼り付け
6. 確認: 先頭が `Option Explicit` で始まっていれば OK

### Step 1.5 コンパイル確認

VBE の `[デバッグ] → [VBAProject のコンパイル]` を実行。エラーが出なければ次へ。

### Step 1.6 セットアップマクロを実行

1. Excel に戻る（`Alt + F11` で再度切り替え）
2. `[開発] → [マクロ]` または **`Alt + F8`** でマクロダイアログを開く
3. リストから **`SetupSheetsAndButtons`** を選択
4. **`[実行]`** をクリック
5. 完了ダイアログ「セットアップ完了。」が出たら成功

このマクロは以下を実行します。

- 必要シート 14 個（`メイン`, `検索`, `ログ` 等）を自動生成
- ログシートのヘッダー行（日時/モジュール名/関数名/メッセージ種別/メッセージ内容）を書込
- 設定シート C6 にテストモードフラグ（`"FALSE"` = 本番モード）を設定
- メインシート B25:E26 にタスクボタン用ラベルを書込
- **フォームコントロールボタン 29 個** を全シートに配置 + マクロ割り当て
- メインシート以外を非表示（起動時はメインのみ表示）
- 既定の空 Sheet1 を削除

### Step 1.7 保存して再オープン

1. **`Ctrl + S`** で保存
2. Excel を一度閉じる
3. 再度同じ `.xlsm` を開く → `Workbook_Open` が走り、メインシートが表示される

メインシート行 25-26 に 8 個のボタン（`▶初回セットアップ` 等）が並んでいれば成功です。これでセットアップは完了です。日常操作は [操作手順](operations.md) へ。

---

## 2. デモパッケージを使う場合（`knowledge_test_v2_demo.xlsm`）

Cowork デモ用。`Workbook_Open` イベントが自動でセットアップとデモデータ投入を実行するので、ユーザは **コンテンツの有効化を押して待つだけ** です。

### Step 2.1 パッケージをローカルにコピー

`Downloads\knowledge_test_v2_demo\` 一式を任意の書込権限のあるローカルフォルダにコピーします（例: `C:\Users\<あなた>\Documents\knowledge_demo\`）。`data/` と `kb_images/` は `.xlsm` と同じ階層を維持してください。

### Step 2.2 .xlsm を開く

`knowledge_test_v2_demo_v6.xlsm` をダブルクリック → 黄色の警告帯が出たら **`[コンテンツの有効化]`** をクリック。

!!! tip "インターネットゾーンブロック対策"
    ファイル右クリック → プロパティ → セキュリティ欄に「許可する」チェックボックスがあれば ON にしてから開いてください。Excel 365 のインターネットゾーンブロックで `Workbook_Open` が走らないことがあります。

### Step 2.3 自動初期化を待つ

`[コンテンツの有効化]` 後、`Workbook_Open` イベントが自動で:

1. **シート未生成なら** `SetupSheetsAndButtons` を実行（14 シート + 29 ボタン + 検索 14 行目 9 列ヘッダ生成）
2. **デモデータ未投入なら** `Macro_SeedDemoData` を実行（DEMO-MEMO フォーマット登録 + dataFolder 設定 + ナレッジ一覧 5 件投入）
3. ログクリア + 起動ログ書込
4. メインシートのみ表示

完了すると「セットアップ完了。」「デモデータの初期化が完了しました。」のメッセージが順に出ます（各 `OK` で閉じる）。

!!! info "冪等性"
    既にセットアップ済の状態（検索シート存在 & ナレッジ一覧 row 3 にデータあり）で開けば自動初期化はスキップされます。何度開いても破綻しません。

セットアップが完了したら [操作手順](operations.md) で日常操作を試してください。

---

## 3. セットアップ時のトラブルシューティング

### Q1. 開いた直後の自動初期化が走らない

- **原因 A**: マクロが無効。**対処**: `[コンテンツの有効化]` をクリック。
- **原因 B**: `Workbook_Open` がエラーで止まった（MsgBox で「自動初期化に失敗しました: ...」が出る）。**対処**: `Alt + F8` → `SetupSheetsAndButtons` 手動実行 → `Macro_SeedDemoData` 手動実行。

### Q2. 「コンパイルエラー」が出る

VBE（`Alt + F11`） → `[デバッグ] → [VBAProject のコンパイル]` で詳細を確認します。

1. 赤くなる行のモジュール名 / 行番号 / メッセージを記録
2. （一時回避）該当モジュールを VBE 上で右クリック → 解放 → 再 Import
3. 全モジュールが揃っているか Step 1.3 のリストと突き合わせ

### Q3. `SetupSheetsAndButtons` マクロが見つからない

`modSetup.bas` のインポート漏れの可能性。`Alt + F11` の VBE 左ペインで `インストーラ層 → modSetup`（または標準モジュールに `modSetup`）があるか確認してください。

### Q4. 「コンパイルエラー: メソッドまたはデータ メンバーが見つかりません」

どこかのモジュールがインポート漏れの可能性。VBE 左ペインで以下が全て揃っているか確認:

- 標準モジュール（`mod*`）: 12 個（`modCommon`, `modDateUtil`, `modEntry系6`, `modFileIO`, `modImageRender`, `modSetup`, `modSpecExamples`, `modStringUtil`, `modFormBuilder`）
- クラスモジュール（`cls*`）: 11 個（`clsControlSpec`, `clsFieldMigrator`, `clsFormSpec`, `clsFormatManager`, `clsKnowledgeManager`, `clsLogger`, `clsLogEntry`, `clsSearchEngine`, `clsStorageResolver`, `clsTaskController`）
- Microsoft Excel Objects: `ThisWorkbook`（中身が貼られている）

### Q5. 起動時に「インデックスが有効範囲にありません」

`SetupSheetsAndButtons` の実行を忘れていないか確認してください。シートが無いと `Workbook_Open` がメインシートを参照できず本エラーになります。

### Q6. ボタンが既存の表示と重なって見えにくい

新規 `.xlsm` の既定の列幅・行高さのため、ボタンが小さく見える場合があります。各シートの該当行（例: メイン 25-26 行目）の高さを増やすか、列幅を広げると見栄えが改善します。本マクロは機能配置のみで装飾は最小化しています。

### Q7. セットアップを最初からやり直したい

`SetupSheetsAndButtons` は **何度実行しても安全**（既存ボタンは削除→再配置、シートは存在チェック後スキップ）です。再度マクロを実行してください。シート自体を消したい場合は、対象シートを手動削除してから `SetupSheetsAndButtons` を再実行すれば再生成されます。

---

## 4. セットアップが触る範囲・触らない範囲

| 項目 | セットアップマクロが触るか |
|---|---|
| 必要シート 14 個の作成 | あり（既存スキップ、不在分のみ作成） |
| ログシート 1 行目ヘッダー | あり（上書き） |
| 設定シート C6（テストモード） | あり（`"FALSE"` 設定） |
| メインシート B24:E26 | あり（案内テキスト + ボタン用ラベル書込） |
| 各シートのフォームコントロールボタン | あり（同名は削除→再配置、29 個） |
| 各シートの初期可視性 | あり（メインのみ表示、他は非表示） |
| 空の既定 Sheet1 | あり（削除、データ・シェイプがあれば残す） |
| その他のセル内容 | なし |
| ユーザフォーム | なし |
| Excel シートの書式・色 | なし |


---

## TODO・制約・既知の限界

### 制約

- VBA 子プロセス禁止（Shell/Run/WScript.Shell/Exec）— 職場 PC ポリシー (ADR-0002)
- ChromaDB 等の外部 RAG 連携無効化 — Cowork 隔離 (ADR-0004)
- クラスモジュール (.cls) 内の `Public Const/Type/Declare/Static` 禁止 (ADR-0027)
- aspose-cells-python 単独では VBA binary stub のみ生成、real Excel COM 必須 (ADR-0026)
- mkdocs Material のモバイル UX が完璧ではない（ナビ collapse は OK、図解の細部は要確認）

### 既知の限界

- Excel 単体動作のため Web/モバイルでは利用不可
- 同時編集なし（ファイルベース、Git 等での並行編集が必要）
- 検索性能はモジュール数 O(n) 線形（数千ナレッジまでは実用、それ以上は ChromaDB 移行検討）

### TODO（v15 以降のロードマップ）

- M-5: modImageRender の RowHeight 副作用排除
- D-3: clsKnowledgeManager / clsSearchEngine / clsTaskController に Worksheet DI 追加
- D-5: 責務固有の Const をクラス側へ移動
- Minor m-1: SHEET_* / FIELD_TYPE_* の Enum 化
- ベンチマーク取得（A+ 到達条件）
- 単体テストカバレッジ整備（A+ 到達条件）

詳細は ADR ([0001-0033](https://github.com/ai-crafted-portfolio/knowledgevba)) を参照。
