---
title: ソースコード
description: VBA モジュール 28 本の全ソースをコピペ可能な形で公開（v13 最新）
---

# ソースコード

VBA モジュール 28 本（本番 25 + 特殊 1 + Cowork デモ専用 3 −1 重複）の全ソースをコピペ可能な形で公開しています。各ページ右上のコピーボタンから VBE（Visual Basic Editor、Excel に標準搭載されているコードエディタ。`Alt+F11` で開く）に直接貼り付けできます。

!!! tip "用語の前置き"
    本サイトでは VBA 慣習に従い、ファイル名のプレフィックスで層を区別しています。`mod` で始まるものは **標準モジュール**（手続き・関数を集めた `.bas` ファイル）、`cls` で始まるものは **クラスモジュール**（オブジェクト指向で状態を保持する `.cls` ファイル）、`ThisWorkbook` は Excel が自動で持つ **特殊なドキュメントモジュール** です。

> **このページが対応する版**: `release_v2 + image_ext + v13` (2026-05-07)。v8〜v13 の修正をすべて反映済みです。最新コミット情報は本ページ末尾の「バージョン」セクションを参照してください。

!!! note "公開可視性"
    本ページ群も `noindex,nofollow` を設定しています。職場利用想定のため、検索エンジンへの indexing は意図的に抑止しています。

## 配布パターン

[ADR-0008 モジュール配布パターン](../architecture.md#5-adr) に従って次の手順で展開します。

1. このサイト上のソースコードを VBE に貼り付け（標準モジュール / クラスモジュール / ThisWorkbook の区別はページ内「配置先」セクション参照）
2. `Alt+F8` から `SetupSheetsAndButtons` を 1 回実行 → 14 シート + 29 ボタン自動生成
3. 検索 → 投入 → 表示の通常運用

詳細は [操作手順](../operations.md) を参照してください。

## 一覧（層別）

### エントリポイント層 (6)

利用者の操作（マクロボタン押下 / Workbook イベント）を受け取り、ビジネスロジック層に処理を委譲する薄い受け口層です。標準モジュール（`mod` プレフィックス）で実装し、引数なしまたは `ByVal` の単純型のみを受け取る Public Sub だけを置きます。エラーハンドリングと MsgBox 表示など UI 寄りの副作用はこの層に集約し、下位の層は副作用を持たない構造にしています。

| ファイル | 役割 |
|---|---|
| [modEntryMain.bas](entrypoint/modentrymain.md) | 本番モード起点 / 共通 BuildLogger 提供 |
| [modEntrySearch.bas](entrypoint/modentrysearch.md) | 検索シート上のマクロボタン受け口 |
| [modEntryKnowledge.bas](entrypoint/modentryknowledge.md) | ナレッジ一覧シートのマクロボタン受け口 |
| [modEntryFormat.bas](entrypoint/modentryformat.md) | フォーマット一覧シートのマクロボタン受け口 |
| [modEntrySettings.bas](entrypoint/modentrysettings.md) | 設定シートのマクロボタン受け口 (dataFolder 切替) |
| [modSpecExamples.bas](entrypoint/modspecexamples.md) | clsFormSpec DSL 利用例 (詳細プレビュー UserForm 構築) |

### ビジネスロジック層 (9)

検索エンジン・ナレッジ管理・タスク制御・ストレージ解決・フォーマット定義・UserForm 動的生成といったドメインロジックを担当する層です。クラスモジュール（`cls` プレフィックス）でステートを保持し、エントリポイント層（`modEntry*`）からの呼出に応答します。`modFormBuilder` だけは UserForm を実行時生成する性質上、標準モジュールとして配置しています。

| ファイル | 役割 |
|---|---|
| [clsSearchEngine.cls](business-logic/clssearchengine.md) | スコアリング検索 / 結果描画 / サムネ画像描画フック |
| [clsKnowledgeManager.cls](business-logic/clsknowledgemanager.md) | ナレッジ一覧シートの行管理 |
| [clsFormatManager.cls](business-logic/clsformatmanager.md) | フォーマット一覧シート管理 / フィールド DSL 解決 |
| [clsTaskController.cls](business-logic/clstaskcontroller.md) | ナレッジ操作のトランザクション制御 |
| [clsStorageResolver.cls](business-logic/clsstorageresolver.md) | dataFolder / kb_images の解決 |
| [clsFieldMigrator.cls](business-logic/clsfieldmigrator.md) | フォーマット定義変更時のスキーマ移行 |
| [clsFormSpec.cls](business-logic/clsformspec.md) | 詳細プレビュー UserForm 仕様 DSL (ルート) |
| [clsControlSpec.cls](business-logic/clscontrolspec.md) | コントロール仕様 DSL (Label / TextBox / Image / Button) |
| [modFormBuilder.bas](business-logic/modformbuilder.md) | clsFormSpec を実行時 UserForm に展開するビルダ |

### ユーティリティ層 (6)

副作用を持たない純粋関数群（文字列・日付・ファイル I/O・画像描画）と共通定数（シート名・列番号など）を集める層です。**上位の層を一切呼ばない** ことを規約として、循環依存を防ぎます。すべての層から自由に呼び出せます。

| ファイル | 役割 |
|---|---|
| [modCommon.bas](utility/modcommon.md) | 全モジュール共通定数 (シート名 / 列番号 / 行番号など) |
| [modFileIO.bas](utility/modfileio.md) | Shift_JIS + CRLF ファイル I/O / フォルダ操作 |
| [modStringUtil.bas](utility/modstringutil.md) | 文字列処理の純粋関数群 |
| [modDateUtil.bas](utility/moddateutil.md) | 日付・時刻処理の純粋関数群 |
| [modImageRender.bas](utility/modimagerender.md) | Shapes.AddPicture によるサムネ画像配置 (Object late binding 化済) |
| [modFormatColumns.bas](utility/modformatcolumns.md) | フォーマット一覧シート列番号定数 (v12 で clsFormatManager から切出し / ADR-0027) |

### インストーラ層 (1)

新規 `.xlsm` に必要なシート 14 個とフォームコントロールボタン 29 個を VBA だけで自動生成するセットアップマクロを置く層です。本ツールは `.xlsm` 本体を配布せずモジュールだけ配布する方針（[ADR-0008](../architecture.md#5-adr)）のため、シート構造とボタン配置はインストール時に各 PC で生成します。`ThisWorkbook` からのみ呼ばれます。

| ファイル | 役割 |
|---|---|
| [modSetup.bas](infrastructure/modsetup.md) | 14 シート + 29 ボタンを 1 回実行で自動生成 (ADR-0008) |

### ログ層 (2)

実行時ログの蓄積（メモリ上）と、ログシート（5 列構成）への出力をまとめた層です。`clsLogger` がエントリ群を保持し、`clsLogEntry` が 1 行分のレコードを表現します。各層から呼び出せる位置に置いていますが、ユーティリティ層と同様に副作用は最小です。

| ファイル | 役割 |
|---|---|
| [clsLogger.cls](logger/clslogger.md) | ログエントリの蓄積 / シート出力 / TS タイムスタンプ |
| [clsLogEntry.cls](logger/clslogentry.md) | 1 行分のログレコード |

### 特殊モジュール (1)

Excel が `.xlsm` ごとに自動で持つドキュメントモジュールです。`Workbook_Open` 等の Excel イベントを受け取り、必要なら初回起動時のセットアップマクロを起動します。インポートではなく中身のコピペで反映する点が他のモジュールと異なります（VBE が新規 `ThisWorkbook` を作れないため）。

| ファイル | 役割 |
|---|---|
| [ThisWorkbook.cls](special/thisworkbook.md) | 本番版 ThisWorkbook (Workbook イベントハンドラ / 自動初期化なし) |
