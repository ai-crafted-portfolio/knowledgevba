---
title: ソースコード
description: 現行 v2.1 の配布対象 VBA モジュールの全ソースをコピー可能な形で公開
---

# ソースコード

knowledgevba v2.1 の配布対象 VBA モジュール **67 本** の全ソースを、
コピー可能な形で層別に掲載しています。各ページのコードブロック右上のボタンから
コードを取得し、ご自身の `.xlsm` の VBE（Visual Basic Editor。Excel 上で
`Alt + F11` で開きます）に取り込んでご利用ください。取り込み手順の全体像は
[導入手順](../setup.md)にまとめています。

!!! note "テスト用モジュールは含みません"
    開発時に使う自動テスト用のモジュールは配布対象ではないため、本一覧には
    含めていません。掲載しているのは、ツールを動かすために実際に取り込む
    モジュールだけです。

!!! info "ファイル名のプレフィックスについて"
    VBA の慣習に従い、`mod` で始まるものは**標準モジュール**（手続き・関数の
    集まり = `.bas`）、`cls` で始まるものは**クラスモジュール**（状態を持つ
    オブジェクト = `.cls`）、`I` で始まるものは**インターフェースクラス**、
    `ThisWorkbook` は Excel がブックごとに自動で持つ**特殊なモジュール**です。

## 3 ブック構成とモジュール配置

本ツールは **登録修正.xlsm / 検索.xlsm / 管理.xlsm** の 3 ブックに分かれています。
多くのモジュールは 3 ブックに共通で取り込みますが、画面クラス・各画面の入口・
`ThisWorkbook` はブックごとに異なります。各ページの「配置ブック」欄で、その
モジュールをどのブックに取り込むかを確認できます。

## 層別一覧

### インストーラ層（1）

空のブックに必要なシートを用意するセットアップ入口です。

| モジュール | 配置ブック | 役割 |
|---|---|---|
| [modSetup.bas](infrastructure/modsetup.md) | 3 ブック共通 | 3 ブック共通のセットアップ入口。各ブックに必要なシート（LOG・各画面）を確認し、無ければ作成する |

### エントリポイント層（5）

シート上のボタン操作や起動イベントを受け取り、業務ロジックへ橋渡しする薄い受け口です。

| モジュール | 配置ブック | 役割 |
|---|---|---|
| [modEntryFormat.bas](entrypoint/modentryformat.md) | 管理.xlsm | フォーマット設計画面のボタン処理。フォーマット定義セルとデータの相互変換、保存・読込ワークフロー |
| [modEntryKnowledge.bas](entrypoint/modentryknowledge.md) | 登録修正.xlsm | ナレッジ登録・修正画面のボタン処理。入力セルとナレッジデータの相互変換、保存・読込ワークフロー |
| [modEntrySearch.bas](entrypoint/modentrysearch.md) | 検索.xlsm | ナレッジ検索画面のボタン処理。検索条件セルの読み取りと結果グリッドへの書き出し |
| [modEntrySettings.bas](entrypoint/modentrysettings.md) | 管理.xlsm | 格納先設定（M-10）・設定（M-11）画面の入口。設定値のシート反映・再読込・破棄 |
| [modSpecExamples.bas](entrypoint/modspecexamples.md) | 3 ブック共通 | clsFormSpec を使った UserForm 組み立てのデモ用コード |

### 画面層（15）

M-02〜M-14 の各画面の組み立てと再描画を担うクラス群です。レイアウトの実体は UI 定義ファイルにあり、ここでは描画の指揮だけを行います。

| モジュール | 配置ブック | 役割 |
|---|---|---|
| [clsBackupMgmtScreen.cls](screen/clsbackupmgmtscreen.md) | 管理.xlsm | M-14 と同じ系統のバックアップ管理画面クラス |
| [clsErrorBandScreen.cls](screen/clserrorbandscreen.md) | 管理.xlsm | M-13 のエラーバンド（4 列）表示を担う連動画面クラス |
| [clsFileFormatScreen.cls](screen/clsfileformatscreen.md) | 管理.xlsm | M-13 ファイル形式設定画面の構築・再描画 |
| [clsFormatDesignScreen.cls](screen/clsformatdesignscreen.md) | 管理.xlsm | M-03 フォーマット設計画面の構築・再描画 |
| [clsFormatListScreen.cls](screen/clsformatlistscreen.md) | 管理.xlsm | M-02 フォーマット一覧画面の構築・再描画 |
| [clsFormatPreviewScreen.cls](screen/clsformatpreviewscreen.md) | 管理.xlsm | M-04 フォーマットプレビュー画面の構築・再描画 |
| [clsKnowledgeEditScreen.cls](screen/clsknowledgeeditscreen.md) | 登録修正.xlsm | M-06 ナレッジ修正画面の構築・再描画 |
| [clsKnowledgeListScreen.cls](screen/clsknowledgelistscreen.md) | 検索.xlsm | M-07 ナレッジ一覧画面の構築・再描画 |
| [clsKnowledgeRegisterScreen.cls](screen/clsknowledgeregisterscreen.md) | 登録修正.xlsm | M-05 ナレッジ登録画面の構築・再描画 |
| [clsKnowledgeViewScreen.cls](screen/clsknowledgeviewscreen.md) | 検索.xlsm | M-09 ナレッジ表示画面の構築・再描画 |
| [clsLogScreen.cls](screen/clslogscreen.md) | 管理.xlsm | M-14 操作ログ画面の構築・再描画 |
| [clsMigrationScreen.cls](screen/clsmigrationscreen.md) | 管理.xlsm | M-12 フィールド反映画面の構築・再描画 |
| [clsSearchScreen.cls](screen/clssearchscreen.md) | 検索.xlsm | M-08 ナレッジ検索画面の構築・再描画 |
| [clsStorageScreen.cls](screen/clsstoragescreen.md) | 管理.xlsm | M-10 格納先設定画面の構築・再描画 |
| [clsSystemSettingsScreen.cls](screen/clssystemsettingsscreen.md) | 管理.xlsm | M-11 設定画面の構築・再描画 |

### ビジネスロジック層（19）

ナレッジ・フォーマット・検索・ログなど、ツールの中心的な業務処理を担う層です。

| モジュール | 配置ブック | 役割 |
|---|---|---|
| [clsButtonSpec.cls](business-logic/clsbuttonspec.md) | 3 ブック共通 | ボタン 1 個分の仕様を保持する値オブジェクト |
| [clsControlSpec.cls](business-logic/clscontrolspec.md) | 3 ブック共通 | UserForm 上のコントロール 1 個の宣言情報を保持する値オブジェクト |
| [clsFieldMigrator.cls](business-logic/clsfieldmigrator.md) | 3 ブック共通 | フォーマット変更時に既存ナレッジへフィールドを反映し、消失リスク時にバックアップを取る |
| [clsFieldSpec.cls](business-logic/clsfieldspec.md) | 3 ブック共通 | 入力フィールド 1 件の仕様を保持する値オブジェクト |
| [clsFormatManager.cls](business-logic/clsformatmanager.md) | 3 ブック共通 | フォーマット定義の作成・編集・削除と、保存時の既存ナレッジ自動反映の連動 |
| [clsFormSpec.cls](business-logic/clsformspec.md) | 3 ブック共通 | UserForm 1 つの宣言情報を保持する値オブジェクト |
| [clsKnowledgeManager.cls](business-logic/clsknowledgemanager.md) | 3 ブック共通 | ナレッジの登録・読込・採番・楽観ロックなどの業務ロジック |
| [clsLogger.cls](business-logic/clslogger.md) | 3 ブック共通 | 操作ログをログシートへ出力。詳細度フィルタと行数上限ローテーション |
| [clsScreenSpec.cls](business-logic/clsscreenspec.md) | 3 ブック共通 | 1 画面分の構成情報（タイトル・セクション・ボタン・フィールド）を保持する値オブジェクト |
| [clsSearchEngine.cls](business-logic/clssearchengine.md) | 3 ブック共通 | ナレッジ検索の中核。番号直接・キーワード・日付範囲フィルタとスコアリング |
| [clsSectionSpec.cls](business-logic/clssectionspec.md) | 3 ブック共通 | 画面内のセクション帯 1 個の仕様を保持する値オブジェクト |
| [clsSetupOrchestrator.cls](business-logic/clssetuporchestrator.md) | 3 ブック共通 | ブック起動時のセットアップ一括処理（設定読込→ログ初期化→シート構築→保護→起動シート表示） |
| [clsSheetRenderer.cls](business-logic/clssheetrenderer.md) | 3 ブック共通 | 画面描画インターフェースのシート実装。レイアウト適用を modUILoader に委譲する |
| [clsStorageResolver.cls](business-logic/clsstorageresolver.md) | 3 ブック共通 | 格納先設定に基づきファイル参照リンクを解決する |
| [clsUserFormRenderer.cls](business-logic/clsuserformrenderer.md) | 3 ブック共通 | 画面描画インターフェースの UserForm 実装（将来用の入口） |
| [IScreenRenderer.cls](business-logic/iscreenrenderer.md) | 3 ブック共通 | 画面描画の抽象インターフェース（8 メソッド） |
| [modFactory.bas](business-logic/modfactory.md) | 3 ブック共通 | 描画クラス・画面クラスのインスタンス生成を集約するファクトリ |
| [modFormBuilder.bas](business-logic/modformbuilder.md) | 3 ブック共通 | clsFormSpec の宣言情報から UserForm を動的に組み立てる |
| [modScreenRender.bas](business-logic/modscreenrender.md) | 3 ブック共通 | 各画面クラス共通のシート描画入口とログ補助 |

### ユーティリティ層（24）

外部ファイルの入出力、スタンザ解析、文字列・日付処理など、上位の層から共通利用される下支えの層です。

| モジュール | 配置ブック | 役割 |
|---|---|---|
| [clsCellAddrHelper.cls](utility/clscelladdrhelper.md) | 3 ブック共通 | セル番地の計算（列文字変換・オフセット等）を行うヘルパー |
| [clsCellBinding.cls](utility/clscellbinding.md) | 3 ブック共通 | セル 1 個の読み書きを抽象化する薄いラッパー |
| [clsCellIO.cls](utility/clscellio.md) | 3 ブック共通 | セル値の読み書きヘルパー（ワークシート／辞書モックの両対応） |
| [clsGridIO.cls](utility/clsgridio.md) | 3 ブック共通 | グリッド（表）形式の一括読み書きヘルパー |
| [clsLogEntry.cls](utility/clslogentry.md) | 3 ブック共通 | ログ 1 行分の値オブジェクト |
| [ClsStanzaSection.cls](utility/clsstanzasection.md) | 3 ブック共通 | スタンザ 1 セクションを表す値オブジェクト |
| [ClsStanzaValidationIssue.cls](utility/clsstanzavalidationissue.md) | 3 ブック共通 | スタンザ検証の指摘 1 件を表す値オブジェクト |
| [ClsStanzaValidationResult.cls](utility/clsstanzavalidationresult.md) | 3 ブック共通 | スタンザ検証結果を集約する値オブジェクト |
| [modCommon.bas](utility/modcommon.md) | 3 ブック共通 | 全ブック共通の定数群 |
| [modConfigHolder.bas](utility/modconfigholder.md) | 3 ブック共通 | 設定値をメモリに保持し、各層へ取得用のメソッドで提供する |
| [modConfigLoader.bas](utility/modconfigloader.md) | 3 ブック共通 | 起動時に <ブック名>_config.txt を読み込み modConfigHolder へ渡す（読み取り専用） |
| [modDateUtil.bas](utility/moddateutil.md) | 3 ブック共通 | 日付・時刻処理の純粋関数群 |
| [modFileIO.bas](utility/modfileio.md) | 3 ブック共通 | Shift_JIS ＋ CRLF の低レベルファイル入出力 |
| [modFormatColumns.bas](utility/modformatcolumns.md) | 3 ブック共通 | フォーマット一覧の列番号定数 |
| [modFormatLoader.bas](utility/modformatloader.md) | 3 ブック共通 | フォーマット定義 .txt の読み書き（書き込みは管理.xlsm のみ） |
| [modImageRender.bas](utility/modimagerender.md) | 3 ブック共通 | ナレッジに添付された画像のサムネ・詳細表示 |
| [modKnowledgeFileIO.bas](utility/modknowledgefileio.md) | 3 ブック共通 | ナレッジ .txt の読み書き・バックアップ（書き込みは登録修正.xlsm のみ） |
| [modLogIds.bas](utility/modlogids.md) | 3 ブック共通 | ログ ID 定数の定義 |
| [modSheetMap.bas](utility/modsheetmap.md) | 3 ブック共通 | シート名と画面 ID の対応表 |
| [modStanzaIO.bas](utility/modstanzaio.md) | 3 ブック共通 | スタンザ形式（[セクション] ＋ key=value）の汎用 read / write |
| [modStringUtil.bas](utility/modstringutil.md) | 3 ブック共通 | 文字列処理の純粋関数群 |
| [modUIConfig.bas](utility/moduiconfig.md) | 3 ブック共通 | 画面の見た目に関する既定値（フォールバック用） |
| [modUILayoutExtractor.bas](utility/moduilayoutextractor.md) | 3 ブック共通 | UI レイアウトを UI 定義 .txt として書き出す開発用ツール |
| [modUILoader.bas](utility/moduiloader.md) | 3 ブック共通 | UI 定義 .txt を読み込み、シートにレイアウトを適用する |

### 特殊モジュール（3）

Excel がブックごとに自動で持つ ThisWorkbook モジュールです。起動・終了イベントを受け取ります。

| モジュール | 配置ブック | 役割 |
|---|---|---|
| [ThisWorkbook（検索.xlsm）](special/thisworkbook-search.md) | 検索.xlsm | 検索.xlsm 用の ThisWorkbook。起動時の設定読込・セットアップ・終了ログ |
| [ThisWorkbook（登録修正.xlsm）](special/thisworkbook-register.md) | 登録修正.xlsm | 登録修正.xlsm 用の ThisWorkbook。起動時の設定読込・セットアップ・終了ログ |
| [ThisWorkbook（管理.xlsm）](special/thisworkbook-admin.md) | 管理.xlsm | 管理.xlsm 用の ThisWorkbook。起動時の設定読込・セットアップ・終了ログ |
