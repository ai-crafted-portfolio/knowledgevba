---
title: ソースコード
description: VBA モジュール 48 本 + ThisWorkbook の全ソースをコピペ可能な形で公開
---

# ソースコード

VBA モジュール **48 本 + ThisWorkbook = 49 ファイル** の全ソースをコピペ可能な形で公開しています。各ページ右上のコピーボタンから VBE（Visual Basic Editor、Excel に標準搭載されているコードエディタ。`Alt+F11` で開く）に直接貼り付けできます。

!!! tip "用語の前置き"
    本サイトでは VBA 慣習に従い、ファイル名のプレフィックスで層を区別しています。`mod` で始まるものは **標準モジュール**（手続き・関数を集めた `.bas` ファイル）、`cls` で始まるものは **クラスモジュール**（オブジェクト指向で状態を保持する `.cls` ファイル）、`I` で始まるものは **インターフェイスクラス**、`ThisWorkbook` は Excel が自動で持つ **特殊なドキュメントモジュール** です。

## 配布パターン

このサイト上のソースコードを VBE に貼り付け、`Alt+F8` から `SetupSheetsAndButtons` を 1 回実行すれば、**14 シート + 68 ボタン** が自動生成されます。詳細手順は [セットアップ](../setup.md) を参照してください。

## 一覧（層別）

### エントリポイント層 (6)

利用者の操作（マクロボタン押下 / Workbook イベント）を受け取り、ビジネスロジック層に処理を委譲する薄い受け口層です。標準モジュール（`mod` プレフィックス）で実装し、引数なしまたは `ByVal` の単純型のみを受け取る Public Sub だけを置きます。エラーハンドリングと MsgBox 表示など UI 寄りの副作用はこの層に集約し、下位の層は副作用を持たない構造にしています。

| ファイル | 役割 |
|---|---|
| [modEntryMain.bas](entrypoint/modentrymain.md) | 本番モード起点 / 共通ログヘルパ |
| [modEntrySearch.bas](entrypoint/modentrysearch.md) | 検索シート上のマクロボタン受け口 |
| [modEntryKnowledge.bas](entrypoint/modentryknowledge.md) | ナレッジ一覧シートのマクロボタン受け口 |
| [modEntryFormat.bas](entrypoint/modentryformat.md) | フォーマット一覧シートのマクロボタン受け口 |
| [modEntrySettings.bas](entrypoint/modentrysettings.md) | 設定シートのマクロボタン受け口 (dataFolder 切替) |
| [modSpecExamples.bas](entrypoint/modspecexamples.md) | clsFormSpec DSL 利用例 (詳細プレビュー UserForm 構築) |

### ビジネスロジック層 (21)

検索エンジン・ナレッジ管理・タスク制御・ストレージ解決・フォーマット定義・UserForm 動的生成・画面 spec 駆動描画・ログ蓄積などのドメインロジックを担当する層です。クラスモジュール（`cls` プレフィックス）でステートを保持し、エントリポイント層からの呼出に応答します。`modFormBuilder` / `modFactory` / `modScreenRender` / `modScreenSpecRegistry` は標準モジュールとして配置しています。

| ファイル | 役割 |
|---|---|
| [clsButtonSpec.cls](business-logic/clsbuttonspec.md) | シート上ボタン 1 個分の宣言情報を保持する値オブジェクト |
| [clsControlSpec.cls](business-logic/clscontrolspec.md) | コントロール仕様 DSL (Label / TextBox / Image / Button) |
| [clsFieldMigrator.cls](business-logic/clsfieldmigrator.md) | フォーマット定義変更時のスキーマ移行 |
| [clsFieldSpec.cls](business-logic/clsfieldspec.md) | シート上フィールド 1 件分の宣言情報を保持する値オブジェクト |
| [clsFormSpec.cls](business-logic/clsformspec.md) | 詳細プレビュー UserForm 仕様 DSL (ルート) |
| [clsFormatManager.cls](business-logic/clsformatmanager.md) | フォーマット一覧シート管理 / フィールド DSL 解決 |
| [clsKnowledgeManager.cls](business-logic/clsknowledgemanager.md) | ナレッジ一覧シートの行管理 |
| [clsLogger.cls](business-logic/clslogger.md) | ログエントリの蓄積 / シート出力 / 外部ファイル出力 (TRACE 含む) |
| [clsScreenSpec.cls](business-logic/clsscreenspec.md) | 1 画面分の宣言情報を保持するルート spec |
| [clsSearchEngine.cls](business-logic/clssearchengine.md) | スコアリング検索 / 結果描画 / サムネ画像描画フック |
| [clsSectionSpec.cls](business-logic/clssectionspec.md) | 1 画面内の 1 セクション(帯)の宣言情報 |
| [clsSetupOrchestrator.cls](business-logic/clssetuporchestrator.md) | セットアップ全工程を統括する制御クラス |
| [clsSheetRenderer.cls](business-logic/clssheetrenderer.md) | シート上にボタン / ラベル / 帯を物理配置する描画クラス |
| [clsStorageResolver.cls](business-logic/clsstorageresolver.md) | dataFolder / kb_images の解決 |
| [clsTaskController.cls](business-logic/clstaskcontroller.md) | ナレッジ操作のトランザクション制御 |
| [clsUserFormRenderer.cls](business-logic/clsuserformrenderer.md) | UserForm 描画用の将来用スタブ |
| [IScreenRenderer.cls](business-logic/iscreenrenderer.md) | 画面描画の抽象インターフェイス |
| [modFactory.bas](business-logic/modfactory.md) | 画面層クラスと Renderer のファクトリ |
| [modFormBuilder.bas](business-logic/modformbuilder.md) | clsFormSpec を実行時 UserForm に展開するビルダ |
| [modScreenRender.bas](business-logic/modscreenrender.md) | 標準画面の描画委譲（共通処理） |
| [modScreenSpecRegistry.bas](business-logic/modscreenspecregistry.md) | M-01〜M-14 の各画面 spec をハードコードで定義するレジストリ |

### ユーティリティ層 (6)

副作用を持たない純粋関数群（文字列・日付・ファイル I/O・画像描画）と共通定数（シート名・列番号・ログレベル・外部ログパスなど）を集める層です。**上位の層を一切呼ばない** ことを規約として、循環依存を防ぎます。すべての層から自由に呼び出せます。`clsLogEntry` だけは値オブジェクトとしてここに置いています。

| ファイル | 役割 |
|---|---|
| [clsLogEntry.cls](utility/clslogentry.md) | 1 行分のログレコード (値オブジェクト) |
| [modCommon.bas](utility/modcommon.md) | 全モジュール共通定数 (シート名 / 列番号 / 行番号 / ログレベル / 外部ログパス) |
| [modDateUtil.bas](utility/moddateutil.md) | 日付・時刻処理の純粋関数群 |
| [modFileIO.bas](utility/modfileio.md) | Shift_JIS + CRLF ファイル I/O / フォルダ操作 |
| [modImageRender.bas](utility/modimagerender.md) | Shapes.AddPicture によるサムネ画像配置 (Object late binding 化済) |
| [modStringUtil.bas](utility/modstringutil.md) | 文字列処理の純粋関数群 |

### 画面層 (14)

各業務画面（M-01 メイン 〜 M-14 操作ログ）の構築・再描画を担当するクラス群です。`IScreenRenderer` 実装を Init 注入で受け取り、`clsScreenSpec` を `modScreenRender` に委譲して標準描画（タイトル帯・セクション帯・ボタン群・フィールドラベル・一覧ヘッダ・「←メインに戻る」ボタン）を一括処理します。

| ファイル | 担当画面 |
|---|---|
| [clsMainScreen.cls](screen/clsmainscreen.md) | M-01 メイン画面（12 個のタスクボタン配置） |
| [clsFormatListScreen.cls](screen/clsformatlistscreen.md) | M-02 フォーマット一覧 |
| [clsFormatDesignScreen.cls](screen/clsformatdesignscreen.md) | M-03 フォーマットデザイン |
| [clsFormatPreviewScreen.cls](screen/clsformatpreviewscreen.md) | M-04 フォーマットプレビュー |
| [clsKnowledgeRegisterScreen.cls](screen/clsknowledgeregisterscreen.md) | M-05 ナレッジ登録 |
| [clsKnowledgeEditScreen.cls](screen/clsknowledgeeditscreen.md) | M-06 ナレッジ編集 |
| [clsKnowledgeListScreen.cls](screen/clsknowledgelistscreen.md) | M-07 ナレッジ一覧 |
| [clsSearchScreen.cls](screen/clssearchscreen.md) | M-08 検索 |
| [clsKnowledgeViewScreen.cls](screen/clsknowledgeviewscreen.md) | M-09 ナレッジ表示 |
| [clsStorageScreen.cls](screen/clsstoragescreen.md) | M-10 格納先設定 |
| [clsSystemSettingsScreen.cls](screen/clssystemsettingsscreen.md) | M-11 システム設定 |
| [clsMigrationScreen.cls](screen/clsmigrationscreen.md) | M-12 フィールド反映 |
| [clsFileFormatScreen.cls](screen/clsfileformatscreen.md) | M-13 ファイル形式 |
| [clsLogScreen.cls](screen/clslogscreen.md) | M-14 操作ログ |

### インストーラ層 (1)

新規 `.xlsm` に必要なシート 14 個とフォームコントロールボタン 68 個を VBA だけで自動生成するセットアップマクロを置く層です。本ツールは `.xlsm` 本体を配布せずモジュールだけ配布する方針のため、シート構造とボタン配置はインストール時に各 PC で生成します。`ThisWorkbook` からのみ呼ばれます。

| ファイル | 役割 |
|---|---|
| [modSetup.bas](infrastructure/modsetup.md) | `SetupSheetsAndButtons` のエントリ。中身は `clsSetupOrchestrator` に委譲 |

### 特殊モジュール (1)

Excel が `.xlsm` ごとに自動で持つドキュメントモジュールです。`Workbook_Open` 等の Excel イベントを受け取り、必要なら初回起動時のセットアップマクロを起動します。インポートではなく中身のコピペで反映する点が他のモジュールと異なります（VBE が新規 `ThisWorkbook` を作れないため）。

| ファイル | 役割 |
|---|---|
| [ThisWorkbook.cls](special/thisworkbook.md) | 本番版 ThisWorkbook (Workbook_Open / Workbook_BeforeClose イベント / メイン画面初期表示) |
