---
title: ソースコード
description: VBA モジュール 27 本の全ソースをコピペ可能な形で公開
---

# ソースコード

VBA モジュール 27 本（本番 24 + 特殊 1 + Cowork デモ専用 3 −1 重複）の全ソースをコピペ可能な形で公開しています。各ページ右上のコピーボタンから VBE に直接貼り付けできます。

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

利用者の操作（マクロボタン押下 / Workbook イベント）を受け取り、ビジネスロジック層に処理を委譲する薄い受け口層。

| ファイル | 役割 |
|---|---|
| [modEntryMain.bas](entrypoint/modentrymain.md) | 本番モード起点 / 共通 BuildLogger 提供 |
| [modEntrySearch.bas](entrypoint/modentrysearch.md) | 検索シート上のマクロボタン受け口 |
| [modEntryKnowledge.bas](entrypoint/modentryknowledge.md) | ナレッジ一覧シートのマクロボタン受け口 |
| [modEntryFormat.bas](entrypoint/modentryformat.md) | フォーマット一覧シートのマクロボタン受け口 |
| [modEntrySettings.bas](entrypoint/modentrysettings.md) | 設定シートのマクロボタン受け口 (dataFolder 切替) |
| [modSpecExamples.bas](entrypoint/modspecexamples.md) | clsFormSpec DSL 利用例 (詳細プレビュー UserForm 構築) |

### ビジネスロジック層 (9)

検索 / 投入 / フォーマット / プレビュー UserForm 構築といった主要機能を実装する層。

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

### ユーティリティ層 (5)

副作用を持たない純粋関数群と共通定数。

| ファイル | 役割 |
|---|---|
| [modCommon.bas](utility/modcommon.md) | 全モジュール共通定数 (シート名 / 列番号 / 行番号など) |
| [modFileIO.bas](utility/modfileio.md) | Shift_JIS + CRLF ファイル I/O / フォルダ操作 |
| [modStringUtil.bas](utility/modstringutil.md) | 文字列処理の純粋関数群 |
| [modDateUtil.bas](utility/moddateutil.md) | 日付・時刻処理の純粋関数群 |
| [modImageRender.bas](utility/modimagerender.md) | Shapes.AddPicture によるサムネ画像配置 |

### インストーラ層 (1)

VBA だけで 14 シート + 29 ボタンを自動生成するセットアップマクロ。

| ファイル | 役割 |
|---|---|
| [modSetup.bas](infrastructure/modsetup.md) | 14 シート + 29 ボタンを 1 回実行で自動生成 (ADR-0008) |

### ログ層 (2)

実行時ログの蓄積とシート出力。

| ファイル | 役割 |
|---|---|
| [clsLogger.cls](logger/clslogger.md) | ログエントリの蓄積 / シート出力 / TS タイムスタンプ |
| [clsLogEntry.cls](logger/clslogentry.md) | 1 行分のログレコード |

### 特殊モジュール (1)

Excel が自動で持つドキュメントモジュール。

| ファイル | 役割 |
|---|---|
| [ThisWorkbook.cls](special/thisworkbook.md) | 本番版 ThisWorkbook (Workbook イベントハンドラ / 自動初期化なし) |

### Cowork デモ専用 (3)

職場 PC ではなく Cowork デモ環境専用の補助モジュール。**本番配布物には含めません。** Cowork デモ用 `.xlsm` を再現する場合だけインポートしてください。

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

ベースは `release_v2 + image_ext rev1`（2026-05-04 / 本番 24 モジュール）に Cowork デモ用 v6 ビルド（2026-05-06）の追加 3 モジュールを合わせた構成です。
