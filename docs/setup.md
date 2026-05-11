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

!!! tip "全体の流れ"
    新規 `.xlsm` 作成 → トラストセンターで信頼設定 → VBE で **48 モジュール** を Import + `ThisWorkbook` 中身コピペ → コンパイル確認 → `SetupSheetsAndButtons` を 1 回実行 → 保存・再オープン、の順で進みます。所要時間は初回 15 〜 25 分程度です。

!!! warning "旧手順との違い（重要）"
    旧版手順は **23 モジュール / 29 ボタン** でしたが、現バージョンは **48 モジュール / 68 ボタン** に拡張されています。
    **画面層 14 クラス（`clsMainScreen` 〜 `clsLogScreen`）が新しく必要** になりました。
    画面層を 1 つでも import 漏れすると、対応する画面のボタンが配置されません。Step 1.3 のリストを **全件** import してください。

## 0. 前提環境

- Microsoft Excel 2019 / 2021 / 365（Windows 版推奨）
- 書込権限のあるローカルフォルダ（ネットワークドライブ直下は避ける）
- マクロ有効化が許可されているプロファイル

!!! warning "Excel for Mac の制約"
    `Forms.*` ProgID が無いため `Macro_ShowSearchResultPreview`（spec 駆動 UserForm）は動作しません。検索のみなら可。

---

## 1. 配布パッケージを使う場合（モジュール手動 import）

配布物は VBA モジュール **48 個 + `ThisWorkbook`** のみで、`.xlsm` 本体は含みません。各モジュールの中身は本サイトの [ソースコード](source/index.md) ページから直接コピペできます（コピーボタン付き）。

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
3. 以下のフォルダ別ファイル群を **全 48 件** 順次インポート

=== "エントリポイント層 (6)"

    | ファイル | 役割 |
    |---|---|
    | [`modEntryMain.bas`](source/entrypoint/modentrymain.md) | 本番モード起点 / 共通ログヘルパ |
    | [`modEntrySearch.bas`](source/entrypoint/modentrysearch.md) | 検索シート上のマクロボタン受け口 |
    | [`modEntryKnowledge.bas`](source/entrypoint/modentryknowledge.md) | ナレッジ一覧シートのマクロボタン受け口 |
    | [`modEntryFormat.bas`](source/entrypoint/modentryformat.md) | フォーマット一覧シートのマクロボタン受け口 |
    | [`modEntrySettings.bas`](source/entrypoint/modentrysettings.md) | 設定シートのマクロボタン受け口 (dataFolder 切替) |
    | [`modSpecExamples.bas`](source/entrypoint/modspecexamples.md) | clsFormSpec DSL 利用例（詳細プレビュー UserForm 構築） |

=== "ビジネスロジック層 (21)"

    | ファイル | 役割 |
    |---|---|
    | [`clsButtonSpec.cls`](source/business-logic/clsbuttonspec.md) | シート上ボタン 1 個分の宣言情報を保持する値オブジェクト |
    | [`clsControlSpec.cls`](source/business-logic/clscontrolspec.md) | コントロール仕様 DSL (Label / TextBox / Image / Button) |
    | [`clsFieldMigrator.cls`](source/business-logic/clsfieldmigrator.md) | フォーマット定義変更時のスキーマ移行 |
    | [`clsFieldSpec.cls`](source/business-logic/clsfieldspec.md) | シート上フィールド 1 件分の宣言情報を保持する値オブジェクト |
    | [`clsFormSpec.cls`](source/business-logic/clsformspec.md) | 詳細プレビュー UserForm 仕様 DSL（ルート） |
    | [`clsFormatManager.cls`](source/business-logic/clsformatmanager.md) | フォーマット一覧シート管理 / フィールド DSL 解決 |
    | [`clsKnowledgeManager.cls`](source/business-logic/clsknowledgemanager.md) | ナレッジ一覧シートの行管理 |
    | [`clsLogger.cls`](source/business-logic/clslogger.md) | ログエントリの蓄積 / シート出力 / 外部ファイル出力 (TRACE 含む) |
    | [`clsScreenSpec.cls`](source/business-logic/clsscreenspec.md) | 1 画面分の宣言情報（タイトル / セクション / ボタン / フィールド） |
    | [`clsSearchEngine.cls`](source/business-logic/clssearchengine.md) | スコアリング検索 / 結果描画 / サムネ画像描画フック |
    | [`clsSectionSpec.cls`](source/business-logic/clssectionspec.md) | 1 画面内の 1 セクション（帯）の宣言情報 |
    | [`clsSetupOrchestrator.cls`](source/business-logic/clssetuporchestrator.md) | セットアップ全工程を統括する制御クラス |
    | [`clsSheetRenderer.cls`](source/business-logic/clssheetrenderer.md) | シート上にボタン / ラベル / 帯を物理配置する描画クラス |
    | [`clsStorageResolver.cls`](source/business-logic/clsstorageresolver.md) | dataFolder / kb_images の解決 |
    | [`clsTaskController.cls`](source/business-logic/clstaskcontroller.md) | ナレッジ操作のトランザクション制御 |
    | [`clsUserFormRenderer.cls`](source/business-logic/clsuserformrenderer.md) | UserForm 描画用の将来用スタブ |
    | [`IScreenRenderer.cls`](source/business-logic/iscreenrenderer.md) | 画面描画の抽象インターフェイス |
    | [`modFactory.bas`](source/business-logic/modfactory.md) | 画面層クラスと Renderer のファクトリ |
    | [`modFormBuilder.bas`](source/business-logic/modformbuilder.md) | clsFormSpec を実行時 UserForm に展開するビルダ |
    | [`modScreenRender.bas`](source/business-logic/modscreenrender.md) | 標準画面の描画委譲（共通処理） |
    | [`modScreenSpecRegistry.bas`](source/business-logic/modscreenspecregistry.md) | M-01〜M-14 の各画面 spec 定義 |

=== "ユーティリティ層 (6)"

    | ファイル | 役割 |
    |---|---|
    | [`clsLogEntry.cls`](source/utility/clslogentry.md) | 1 行分のログレコード（値オブジェクト） |
    | [`modCommon.bas`](source/utility/modcommon.md) | 全モジュール共通定数（シート名 / 列番号 / 行番号 / ログレベル / 外部ログパス） |
    | [`modDateUtil.bas`](source/utility/moddateutil.md) | 日付・時刻処理の純粋関数群 |
    | [`modFileIO.bas`](source/utility/modfileio.md) | Shift_JIS + CRLF ファイル I/O / フォルダ操作 |
    | [`modImageRender.bas`](source/utility/modimagerender.md) | Shapes.AddPicture によるサムネ画像配置 |
    | [`modStringUtil.bas`](source/utility/modstringutil.md) | 文字列処理の純粋関数群 |

=== "画面層 (14) ★必須★"

    | ファイル | 担当画面 |
    |---|---|
    | [`clsMainScreen.cls`](source/screen/clsmainscreen.md) | M-01 メイン画面（タスクボタン 12 個） |
    | [`clsFormatListScreen.cls`](source/screen/clsformatlistscreen.md) | M-02 フォーマット一覧 |
    | [`clsFormatDesignScreen.cls`](source/screen/clsformatdesignscreen.md) | M-03 フォーマットデザイン |
    | [`clsFormatPreviewScreen.cls`](source/screen/clsformatpreviewscreen.md) | M-04 フォーマットプレビュー |
    | [`clsKnowledgeRegisterScreen.cls`](source/screen/clsknowledgeregisterscreen.md) | M-05 ナレッジ登録 |
    | [`clsKnowledgeEditScreen.cls`](source/screen/clsknowledgeeditscreen.md) | M-06 ナレッジ編集 |
    | [`clsKnowledgeListScreen.cls`](source/screen/clsknowledgelistscreen.md) | M-07 ナレッジ一覧 |
    | [`clsSearchScreen.cls`](source/screen/clssearchscreen.md) | M-08 検索 |
    | [`clsKnowledgeViewScreen.cls`](source/screen/clsknowledgeviewscreen.md) | M-09 ナレッジ表示 |
    | [`clsStorageScreen.cls`](source/screen/clsstoragescreen.md) | M-10 格納先設定 |
    | [`clsSystemSettingsScreen.cls`](source/screen/clssystemsettingsscreen.md) | M-11 システム設定 |
    | [`clsMigrationScreen.cls`](source/screen/clsmigrationscreen.md) | M-12 フィールド反映 |
    | [`clsFileFormatScreen.cls`](source/screen/clsfileformatscreen.md) | M-13 ファイル形式 |
    | [`clsLogScreen.cls`](source/screen/clslogscreen.md) | M-14 操作ログ |

    !!! danger "画面層 14 クラスは 1 つでも欠落させない"
        画面層フォルダの 14 クラスを 1 つでも import 漏れすると、`clsSetupOrchestrator` が `modFactory.CreateScreen` 経由で生成に失敗し、対応する画面のボタンが配置されません。
        VBE 左ペインのモジュール一覧に **`clsMainScreen` から `clsLogScreen` まで 14 個揃っているか必ず確認** してください。

=== "インストーラ層 (1)"

    | ファイル | 役割 |
    |---|---|
    | [`modSetup.bas`](source/infrastructure/modsetup.md) | `SetupSheetsAndButtons` のエントリポイント。中身は `clsSetupOrchestrator` に委譲 |

各モジュールの中身は対応するソースページからコピペできます。新規モジュールを VBE で挿入してから貼り付けるか、`.bas`/`.cls` ファイルに保存して `[ファイルのインポート...]` で取り込んでください。

### Step 1.4 ThisWorkbook の中身を貼り付け（インポートではない）

`ThisWorkbook.cls` だけは「インポート」ではなく **中身のコピペ** が必要です。VBE は新しい `ThisWorkbook` を作れず、既存の `ThisWorkbook` に上書きする必要があるためです。

1. VBE 左ペインで **`Microsoft Excel Objects` → `ThisWorkbook`** をダブルクリック
2. 右側のコードペインを開く（現状は空のはず）
3. [ThisWorkbook.cls](source/special/thisworkbook.md) をブラウザで開いてコピーボタンを押す
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

!!! tip "「メソッドまたはデータ メンバーが見つかりません」が出たら"
    画面層 14 クラス、特に `clsSheetRenderer` / `clsSetupOrchestrator` / `clsMainScreen` の Import 漏れが最も多い原因です。
    VBE 左ペインのモジュール一覧の合計が **48 + ThisWorkbook = 49** になっているかを必ず確認してください。

### Step 1.6 セットアップマクロを実行

1. Excel に戻る（`Alt + F11` で再度切り替え）
2. `[開発] → [マクロ]` または **`Alt + F8`** でマクロダイアログを開く
3. リストから **`SetupSheetsAndButtons`** を選択
4. **`[実行]`** をクリック
5. 完了ダイアログ「セットアップ完了。」が出たら成功

このマクロは以下を実行します。

- 必要シート 14 個（`メイン`, `検索`, `ログ` 等）を自動生成
- ログシートのヘッダー行（日時 / モジュール名 / 関数名 / メッセージ種別 / メッセージ内容）を書込
- 設定シート C6 にテストモードフラグ（`"FALSE"` = 本番モード）を設定
- メインシート B10:K18 にタスクボタン用ラベルを書込
- **フォームコントロールボタン 68 個** を全 14 シートに配置 + マクロ割り当て
    - メイン: **12 個**（▶検索 / ▶ナレッジ登録 / ▶ナレッジ修正 / ▶ナレッジ一覧 / ▶フォーマット管理 / ▶フィールド反映 / ▶格納先設定 / ▶システム設定 / ▶ログ確認 / ▶ファイル形式 / ▶初回セットアップ / ▶ヘルプ）
    - 業務 13 シート: 計 **56 個**（各シート 2〜7 ボタン + `Btn_BackToMain`）
- メインシート以外を非表示（起動時はメインのみ表示）
- 既定の空 Sheet1 を削除

実行中の各工程は **ログシート + 外部ログファイル** にトレース記録されます。失敗時は `step=[XYZ] errNum=N desc=...` の形式で 1 行に集約されます。

### Step 1.7 保存して再オープン

1. **`Ctrl + S`** で保存
2. Excel を一度閉じる
3. 再度同じ `.xlsm` を開く → `Workbook_Open` が走り、メインシートが表示される

メインシート行 10〜18 列 B〜K に 12 個のボタンが並んでいれば成功です。これでセットアップは完了です。日常操作は [操作手順](operations.md) へ。

---

## 2. セットアップ時のトラブルシューティング

### Q1. `SetupSheetsAndButtons` マクロが見つからない

`modSetup.bas` のインポート漏れの可能性。`Alt + F11` の VBE 左ペインで `インストーラ層 → modSetup`（または標準モジュールに `modSetup`）があるか確認してください。

### Q2. 「コンパイルエラー: メソッドまたはデータ メンバーが見つかりません」

どこかのモジュールがインポート漏れの可能性が高いです。VBE 左ペインで以下が全て揃っているか確認してください。

- **標準モジュール**（`mod*` プレフィックス、`.bas`）: **16 個**
    - エントリポイント層 6 個（`modEntryMain` / `modEntrySearch` / `modEntryKnowledge` / `modEntryFormat` / `modEntrySettings` / `modSpecExamples`）
    - ビジネスロジック層 4 個（`modFactory` / `modFormBuilder` / `modScreenRender` / `modScreenSpecRegistry`）
    - ユーティリティ層 5 個（`modCommon` / `modDateUtil` / `modFileIO` / `modImageRender` / `modStringUtil`）
    - インストーラ層 1 個（`modSetup`）
- **クラスモジュール**（`cls*` または `I*` プレフィックス、`.cls`）: **32 個**
    - ビジネスロジック層 17 個（`clsButtonSpec` / `clsControlSpec` / `clsFieldMigrator` / `clsFieldSpec` / `clsFormSpec` / `clsFormatManager` / `clsKnowledgeManager` / `clsLogger` / `clsScreenSpec` / `clsSearchEngine` / `clsSectionSpec` / `clsSetupOrchestrator` / `clsSheetRenderer` / `clsStorageResolver` / `clsTaskController` / `clsUserFormRenderer` / `IScreenRenderer`）
    - ユーティリティ層 1 個（`clsLogEntry`）
    - **画面層 14 個**（`clsMainScreen` / `clsFormatListScreen` / `clsFormatDesignScreen` / `clsFormatPreviewScreen` / `clsKnowledgeRegisterScreen` / `clsKnowledgeEditScreen` / `clsKnowledgeListScreen` / `clsSearchScreen` / `clsKnowledgeViewScreen` / `clsStorageScreen` / `clsSystemSettingsScreen` / `clsMigrationScreen` / `clsFileFormatScreen` / `clsLogScreen`）
- **Microsoft Excel Objects**: `ThisWorkbook`（中身が貼られている）

合計 **48 + ThisWorkbook = 49**。

!!! danger "頻発: 画面層 14 個の Import 漏れ"
    旧手順から移行した利用者が、画面層フォルダの 14 クラスを丸ごと忘れているケースが多発します。`clsMainScreen` / `clsLogScreen` などがモジュール一覧に無ければ、画面層フォルダから追加 import を実施してください。

### Q3. `SetupSheetsAndButtons` 実行中にエラーで止まる

セットアップマクロは各工程をログシート + 外部ログファイルにトレースします。エラー発生時は最後の `step=[XYZ] errNum=N desc=...` を確認してください。

代表的なステップ名と対処:

| step= | 内容 | 主な原因 |
|---|---|---|
| `EnsureSheets` | 14 シートの存在確認・作成 | シート保護 / Workbook 状態異常 |
| `BuildSpecs` | 各画面 spec の構築 | `modScreenSpecRegistry` の import 漏れ |
| `RenderSheet` | ボタン物理配置 | `clsSheetRenderer` / 画面層クラスの import 漏れ |
| `WireOnAction` | ボタンへのマクロ割当 | エントリポイント層 import 漏れ |

### Q4. 起動時に「インデックスが有効範囲にありません」

`SetupSheetsAndButtons` の実行を忘れていないか確認してください。シートが無いと `Workbook_Open` がメインシートを参照できず本エラーになります。

### Q5. ボタンが既存の表示と重なって見えにくい

新規 `.xlsm` の既定の列幅・行高さのため、ボタンが小さく見える場合があります。各シートの該当行（例: メイン 10〜18 行目）の高さを増やすか、列幅を広げると見栄えが改善します。本マクロは機能配置のみで装飾は最小化しています。

### Q6. ボタンが「`Btn_xxxx` を実行できません」とエラー

ブックを「名前を付けて保存」でリネームした後にボタンを押すと発生することがあります。`SetupSheetsAndButtons` を再実行すれば、新しいブック名に対する OnAction が再設定されます。

### Q7. セットアップを最初からやり直したい

`SetupSheetsAndButtons` は **何度実行しても安全**（既存ボタンは削除→再配置、シートは存在チェック後スキップ）です。再度マクロを実行してください。シート自体を消したい場合は、対象シートを手動削除してから `SetupSheetsAndButtons` を再実行すれば再生成されます。

### Q8. ログを詳しく見たい

セッション単位のログはブック内の **「ログ」シート** に残ります（`Workbook_Open` 時に自動クリア）。
動作中の Trace / Error は外部ログファイル（既定 `C:\<書込み可能なフォルダ>\runtime.log`）に随時 Append されます。出力先パスは `modCommon` の `EXTERNAL_LOG_PATH` 定数で変更できます。

---

## 3. セットアップが触る範囲・触らない範囲

| 項目 | セットアップマクロが触るか |
|---|---|
| 必要シート 14 個の作成 | あり（既存スキップ、不在分のみ作成） |
| ログシート 1 行目ヘッダー | あり（上書き） |
| 設定シート C6（テストモード） | あり（`"FALSE"` 設定） |
| メインシート B10:K18 | あり（案内テキスト + ボタン用ラベル書込） |
| 各シートのフォームコントロールボタン | あり（同名は削除→再配置、計 68 個） |
| 各シートの初期可視性 | あり（メインのみ表示、他は非表示） |
| 空の既定 Sheet1 | あり（削除、データ・シェイプがあれば残す） |
| 外部ログファイル | あり（Append） |
| その他のセル内容 | なし |
| ユーザフォーム | なし |
| Excel シートの書式・色 | なし（ボタン配置セルの色塗りは行う） |

---

## 4. 旧手順からの移行メモ

- **旧手順は 23 モジュール / 29 ボタン**、**現手順は 48 モジュール / 68 ボタン** です。
- 旧 `.xlsm` がある場合は、**新規 `.xlsm` を作り直して 48 モジュールを再 import** することを推奨します（増分 import は競合の元）。
- メイン画面のボタンは **8 個 → 12 個** に増えました（`▶ファイル形式` / `▶初回セットアップ` / `▶ヘルプ` 等が追加）。
- ログ機能が強化されており、セッションログ（シート）+ 外部ログファイルの 2 系統に同時記録されます。
