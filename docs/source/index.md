---
title: ソースコード
description: VBA モジュール 28 本の全ソースをコピペ可能な形で公開（v13 最新）
---

# ソースコード

VBA モジュール 28 本（本番 25 + 特殊 1 + Cowork デモ専用 3 −1 重複）の全ソースをコピペ可能な形で公開しています。各ページ右上のコピーボタンから VBE に直接貼り付けできます。

> **このページが対応する版**: `release_v2 + image_ext + v13` (2026-05-07)。v8〜v13 の修正をすべて反映済みです。最新コミット情報は本ページ末尾の「バージョン」セクションを参照してください。

!!! note "公開可視性"
    本ページ群も `noindex,nofollow` を設定しています。職場利用想定のため、検索エンジンへの indexing は意図的に抑止しています。

## 層構成（全体像）

このツールは 5 つの層 + 特殊モジュール + Cowork デモ専用層に分かれて実装されています。各層は**単一責任の原則**に従い、上位層は下位層に依存し、その逆は禁止です。

| 層 | 主担当 | プレフィックス |
|---|---|---|
| エントリポイント層 | ユーザ操作（ボタン）→ ビジネスロジック呼出 | modEntry* |
| ビジネスロジック層 | 検索 / ナレッジ管理 / タスク制御 / フォーマット定義 | cls* / modFormBuilder |
| ユーティリティ層 | ファイル I/O / 文字列 / 日付 / 画像描画 / 共通定数 | mod* |
| インストーラ層 | 初回セットアップ（シート + ボタン生成） | modSetup |
| ログ層 | 実行履歴記録 | clsLogger / clsLogEntry |
| 特殊モジュール | Excel 標準ドキュメントモジュール | ThisWorkbook |
| Cowork デモ専用 | 自動初期化 + デモデータ投入（本番配布外） | modAutoInit / modDemoSeeder |

依存関係の詳細は [アーキテクチャ](../architecture.md) を参照してください。

## 配布パターン

[ADR-0008 モジュール配布パターン](../architecture.md#5-adr) に従って次の手順で展開します。

1. このサイト上のソースコードを VBE に貼り付け（標準モジュール / クラスモジュール / ThisWorkbook の区別はページ内「配置先」セクション参照）
2. `Alt+F8` から `SetupSheetsAndButtons` を 1 回実行 → 14 シート + 29 ボタン自動生成
3. 検索 → 投入 → 表示の通常運用

詳細は [操作手順](../operations.md) を参照してください。

## 一覧（層別）

### エントリポイント層 (6)

利用者の操作（マクロボタン押下 / Workbook イベント）を受け取り、ビジネスロジック層に処理を委譲する薄い受け口層。標準モジュール（mod プレフィックス）で実装され、副作用（MsgBox 表示・エラーメッセージ表示等）はこの層に集約します。ビジネスロジック層は副作用を持たないため、UI 連動部分はすべてここで吸収する設計です。

| ファイル | 役割 |
|---|---|
| [modEntryMain.bas](entrypoint/modentrymain.md) | 本番モード起点 / 共通 BuildLogger 提供 |
| [modEntrySearch.bas](entrypoint/modentrysearch.md) | 検索シート上のマクロボタン受け口 |
| [modEntryKnowledge.bas](entrypoint/modentryknowledge.md) | ナレッジ一覧シートのマクロボタン受け口 |
| [modEntryFormat.bas](entrypoint/modentryformat.md) | フォーマット一覧シートのマクロボタン受け口 |
| [modEntrySettings.bas](entrypoint/modentrysettings.md) | 設定シートのマクロボタン受け口 (dataFolder 切替) |
| [modSpecExamples.bas](entrypoint/modspecexamples.md) | clsFormSpec DSL 利用例 (詳細プレビュー UserForm 構築) |

### ビジネスロジック層 (9)

検索 / 投入 / フォーマット / プレビュー UserForm 構築といったドメインロジックを担当。クラスモジュール（cls プレフィックス）でステートを保持し、エントリポイント層からの呼出に応答します。各クラスは 1 責務に特化し、互いの依存は最小限に抑えてあります。

主なクラス:

- **clsSearchEngine**: スコアリング検索本体（タイトル ×3 + フィールド ×2 ブースト）
- **clsKnowledgeManager**: ナレッジ一覧シートの行 CRUD
- **clsFormatManager**: フォーマット定義（カラム配置 / バリデーション）の管理
- **clsTaskController**: ナレッジ投入のトランザクション制御
- **clsStorageResolver**: dataFolder / kb_images パスの解決
- **clsFieldMigrator**: フォーマット定義変更時のスキーママイグレーション
- **clsFormSpec / clsControlSpec**: 動的 UserForm 生成のための仕様定義（DSL）

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

副作用を最小化した純粋関数群と全モジュール共通の定数。再利用性を確保するため、シート名 / 列番号 / 行番号といった「変動しがちな数値」はすべて `modCommon` および `modFormatColumns` に集約しています。

主なモジュール:

- **modCommon**: 全モジュール共通の定数（シート名・列番号・行番号など）
- **modFileIO**: Shift_JIS + CRLF ファイル I/O ヘルパ
- **modStringUtil**: 文字列処理の純粋関数群
- **modDateUtil**: 日付処理ヘルパ（TS タイムスタンプ等）
- **modImageRender**: Shapes.AddPicture によるサムネ画像配置（Object late binding 化済）
- **modFormatColumns**: フォーマット一覧シートの列番号定数（v12 で clsFormatManager から切出し、ADR-0027）

| ファイル | 役割 |
|---|---|
| [modCommon.bas](utility/modcommon.md) | 全モジュール共通定数 (シート名 / 列番号 / 行番号など) |
| [modFileIO.bas](utility/modfileio.md) | Shift_JIS + CRLF ファイル I/O / フォルダ操作 |
| [modStringUtil.bas](utility/modstringutil.md) | 文字列処理の純粋関数群 |
| [modDateUtil.bas](utility/moddateutil.md) | 日付・時刻処理の純粋関数群 |
| [modImageRender.bas](utility/modimagerender.md) | Shapes.AddPicture によるサムネ画像配置 (Object late binding 化済) |
| [modFormatColumns.bas](utility/modformatcolumns.md) | フォーマット一覧シート列番号定数 (v12 で clsFormatManager から切出し / ADR-0027) |

### インストーラ層 (1)

VBA だけで 14 シート + 29 ボタンを自動生成するセットアップマクロ。idempotent 設計（複数回実行しても破綻しない）のため、既存シートの差分追加にも使えます。これにより配布物に `.xlsx` テンプレートを含めず、`.bas` / `.cls` のみで配布可能になっています。

| ファイル | 役割 |
|---|---|
| [modSetup.bas](infrastructure/modsetup.md) | 14 シート + 29 ボタンを 1 回実行で自動生成 (ADR-0008) |

### ログ層 (2)

実行時ログの蓄積とシート出力。`m_nextRow` キャッシュで O(1) 書込性能を確保（v13、ADR M-3）。`clsLogger` がエントリを蓄積し、`clsLogEntry` が 1 行分のレコードを表現します。

| ファイル | 役割 |
|---|---|
| [clsLogger.cls](logger/clslogger.md) | ログエントリの蓄積 / シート出力 / TS タイムスタンプ |
| [clsLogEntry.cls](logger/clslogentry.md) | 1 行分のログレコード |

### 特殊モジュール (1)

Excel が自動で持つドキュメントモジュール。本番版は Workbook イベントハンドラのみで、自動初期化は行いません（Cowork デモ版とは別物）。

| ファイル | 役割 |
|---|---|
| [ThisWorkbook.cls](special/thisworkbook.md) | 本番版 ThisWorkbook (Workbook イベントハンドラ / 自動初期化なし) |

### Cowork デモ専用 (3)

職場 PC ではなく Cowork デモ環境専用の補助モジュール。**本番配布物には含めません。** Workbook_Open での自動初期化 + デモナレッジ 5 件の自動投入を行うため、デモを再現する場合だけインポートしてください。

| ファイル | 役割 |
|---|---|
| [ThisWorkbook.cls (デモ版)](cowork-demo/thisworkbook.md) | Workbook_Open で modAutoInit を 1 ティック遅延起動 |
| [modAutoInit.bas](cowork-demo/modautoinit.md) | OnTime 経由でセットアップ + シードを自動実行 |
| [modDemoSeeder.bas](cowork-demo/moddemoseeder.md) | フォーマット / dataFolder / ナレッジ 5 件のワンクリック投入 |

## ライセンス・注意

- 個人利用前提（職場 PC 内の社内ノウハウ蓄積を想定）
- VBA 子プロセス禁止 ([ADR-0002](../architecture.md#5-adr)) — `Shell` / `WScript.Shell.Run` 等は使わない
- ChromaDB 連携無効化 ([ADR-0004](../architecture.md#5-adr)) — `clsSearchEngine` 1 ファイルだけ差し替えれば txt 走査 → Range 走査に切り替えられる構造

## バージョン

ベースは `release_v2 + image_ext rev1`（2026-05-04 / 本番 24 モジュール）に Cowork デモ用 v6 ビルド（2026-05-06）の追加 3 モジュールを合わせた構成。さらに v8〜v13 の累積修正（2026-05-07）を反映しています。

### v8〜v13 で反映された主な修正

| 版 | 内容 | 関連 ADR / モジュール |
|---|---|---|
| v8〜v10 | コンパイル/参照不整合の解消、Setup の冪等性強化 | `modSetup` 改修 |
| **v11** | バイナリ互換性の解決（VBE 側で発生していたインポート不整合の修正） | `modSetup`, `modImageRender` |
| **v12** | `Public Const FL_COL_*` をクラスから標準モジュールへ切出し（実 Excel が `Public Const` をクラス内で reject する仕様への対応） | **新規 `modFormatColumns.bas`** / ADR 0027 |
| **v13** | `clsFieldMigrator.MigrateFields` を Sub → Function に変更し、`Exit Sub` を `Exit Function` に修正（戻り値型整合） | `clsFieldMigrator` / ADR 0028 |
| その他 | `modImageRender` を Object late binding 化、`modSetup` に `silent` 引数を追加、`ThisWorkbook` の OnTime 経由 auto-init を簡素化、Cowork デモ専用の `modAutoInit` を分離 | `modImageRender`, `modSetup`, `ThisWorkbook (デモ版)`, `modAutoInit` |

### UAT ステータス

v13 で UAT 完走済（2026-05-07）。本ページのソースコードは UAT 通過時点のものです。
