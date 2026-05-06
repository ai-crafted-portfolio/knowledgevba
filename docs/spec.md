---
title: 仕様
description: モジュール構成、検索スコアリング、画像解決規則、spec 駆動 UserForm DSL、ChromaDB 切替ポイント、テスト構成
tags:
  - excel
  - vba
  - spec
---

# 仕様

release v2 + image_ext rev1 時点（2026-05-04）の仕様を定義します。

## 1. モジュール構成（24 個）

VBA モジュールは「層」ごとにフォルダ分割した形で配布されます。インポート対象は 23 個、`ThisWorkbook.cls` は中身コピペで合計 24 個です。

| 層 | モジュール |
|---|---|
| インストーラ層 | `modSetup.bas` |
| エントリポイント層 | `modEntryMain.bas` / `modEntrySearch.bas` / `modEntryKnowledge.bas` / `modEntryFormat.bas` / `modEntrySettings.bas` / `modSpecExamples.bas` |
| ビジネスロジック層 | `clsSearchEngine.cls` / `clsKnowledgeManager.cls` / `clsFormatManager.cls` / `clsTaskController.cls` / `clsStorageResolver.cls` / `clsFieldMigrator.cls` / `clsLogger.cls` / `clsControlSpec.cls` / `clsFormSpec.cls` / `modFormBuilder.bas` |
| ユーティリティ層 | `modCommon.bas` / `modStringUtil.bas` / `modDateUtil.bas` / `modFileIO.bas` / `modImageRender.bas` / `clsLogEntry.cls` |
| 特殊モジュール | `ThisWorkbook.cls` |

image_ext rev1 で新規追加された 5 モジュール: `modImageRender` / `clsControlSpec` / `clsFormSpec` / `modFormBuilder` / `modSpecExamples`。

!!! info "層分離の方針"
    エントリポイント層は「ボタン配下から呼ばれる Public Sub」のみを置き、ビジネスロジック層のクラス群（`cls*`）に処理を委譲します。ユーティリティ層はビジネスロジック層から下方向にのみ依存し、上方向には依存しません。詳細は [アーキテクチャ](architecture.md) を参照。

## 2. シート構成（14 個）

`SetupSheetsAndButtons` マクロ（`modSetup.bas`）が以下 14 シートを存在チェック後に不在分のみ生成します。

| ID | シート名 | 役割 |
|---|---|---|
| M-01 | メイン | エントリポイント、各機能ボタン配置 |
| M-04 | 検索 | キーワード入力 + 結果一覧（H 列サムネ / I 列 Score） |
| M-08 | ナレッジ一覧 | メタデータ管理（KnwNo / FormatID / タイトル / 作成日 / 更新日 ほか） |
| M-09 | ナレッジ表示 | 単票表示、J4:N20 領域に詳細画像（最大 400×300px） |
| M-11 | 設定 | `dataFolder` パス、テストモードフラグ |
| ログ | ログ | 5 列構成（日時 / モジュール名 / 関数名 / メッセージ種別 / メッセージ内容） |
| その他 8 シート | フォーマット定義・タスク関連・各種台帳 | — |

セットアップマクロは同時に **フォームコントロールボタン 29 個** を全シートに配置・マクロ割り当てまで実施します。

## 3. 検索スコアリング

`clsSearchEngine.ScanAndMatch` が `<dataFolder>/*.txt` を走査して以下の式で Score を算出し、降順で結果行に書き込みます。

$$
\textit{Score} = 3 \times \textit{TitleHits} + 2 \times \textit{TargetFieldHits} + \textit{TotalOccurrences}
$$

具体例として、デモデータ `Macro_SeedDemoData` でキーワード `メモ` を AND モードで検索した場合の期待順位は次のとおり。

| 順位 | KnwNo | タイトル | 備考 |
|---|---|---|---|
| 1 | KN-2026-0420 | メモリ枯渇エラー対処メモ | タイトル「メモ」2 回 ×3 ブースト + 本文出現 |
| 2 | KN-2026-0421 | ChromaDB HNSW 再構築手順メモ | — |
| 3 | KN-2026-0424 | サムネ画像の自動配置メモ | — |
| 4 | KN-2026-0422 | VBA ADODB.Stream の代替メモ | — |
| 5 | KN-2026-0423 | RAG 検索のスコアリング設計 | タイトル「メモ」なし、本文ヒットのみ |

検索モードは `AND` / `OR` を切替可能。`TargetField` で「全フィールド」「タイトルのみ」「本文のみ」等のスコープを絞れます。

## 4. 画像表示機能（image_ext rev1）

検索結果一覧の H 列にサムネ（60×40px）、I 列に Score を表示し、ナレッジ表示画面の J4:N20 領域に最大 400×300px の詳細画像を `Shapes.AddPicture` で貼付します。

### 4.1 既定の画像パス解決規則

```
<dataFolder>/                       ← 設定画面で指定するナレッジファイル置き場
  ├─ KN-2026-0420.txt
  ├─ KN-2026-0421.txt
  └─ ...
<dataFolder の親>/
  └─ kb_images/                     ← 画像はここに置く（規約）
       ├─ KN-2026-0420.png
       ├─ KN-2026-0421.png
       └─ ...
```

例: `dataFolder = C:\knowledge\data\` なら画像は `C:\knowledge\kb_images\` から解決。

### 4.2 ImagePath スタンザ仕様

ナレッジ `.txt` 末尾（任意の位置）に以下を書くと既定パスを上書きできます。

```
ImagePath=KN-2026-0420.png
```

または絶対パス:

```
ImagePath=\\fileshare\kb_images\custom_image.png
```

スタンザが無い場合は **`<KnwNo>.png` を既定** として動作するため、既存ファイルとの互換性は保たれます。

## 5. spec 駆動 UserForm DSL（image_ext rev1）

`clsFormSpec` で宣言したコントロール仕様を、`modFormBuilder.BuildAndShow` が実行時に `VBProject.VBComponents.Add(3)` で UserForm を一時生成 → `designer.Controls.Add` でコントロール配置 → `CodeModule.AddFromString` でクリックハンドラ注入してから `.Show` する仕組みです。

### 5.1 サンプルコード

```vba
Sub MyForm()
    Dim spec As clsFormSpec
    Set spec = New clsFormSpec
    spec.FormTitle = "サンプル"
    spec.Width = 500 : spec.Height = 350

    Call spec.AddControl("Label", "lblTitle", 10, 10, 480, 20, "見出し")
    Call spec.AddControl("Image", "imgMain", 10, 40, 300, 200)
    Call spec.AddControl("TextBox", "txtNote", 320, 40, 170, 200, "メモ")
    Call spec.AddControl("Button", "btnClose", 200, 280, 80, 30, _
                          "閉じる", "frmCallback_searchResult_close")

    Call BuildAndShow(spec, True)  ' True = Modal
End Sub
```

### 5.2 サポートしているコントロール種別

`Label` / `TextBox` / `Image` / `Button` の 4 種別。`clsControlSpec` が name / left / top / width / height / caption / callbackName を保持し、`clsFormSpec` がコレクションでまとめて持ちます。

!!! warning "前提条件"
    `[ファイル] → [オプション] → [トラスト センター] → [トラスト センターの設定] → [マクロの設定] → [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する]` を **ON** にする必要があります。OFF の場合 `BuildAndShow` で「アクセスが拒否されました」エラーが出ます。検索のみの利用なら OFF のままで構いません。

## 6. フォーマット定義機構

`clsFormatManager` がフォーマット ID（例: `DEMO-MEMO`）ごとにフィールド定義を保持します。各ナレッジ `.txt` の先頭で `FormatID=` を宣言すると、検索画面側の TargetField ドロップダウンに該当フォーマットの宣言フィールドが選択肢として並びます。

`clsFieldMigrator` は旧フォーマットから新フォーマットへフィールド名を写像する役割で、フォーマット定義を変更しても既存ナレッジを壊さずに済むようになっています。

## 7. ChromaDB 連動切替ポイント

本実装は `<dataFolder>/*.txt` を ADODB で SJIS 読込する **モック実装** です。本番では `clsSearchEngine.cls` 冒頭のコメントブロックに記載の手順で:

1. 事前 ETL で chunks を `Sheet "Data"` に export しておく
2. `ScanAndMatch` の txt ループ箇所を Range 走査に置換する

この変更点は `clsSearchEngine` 1 ファイルに局所化されており、上位エントリポイントには影響しません。

!!! danger "VBA からの子プロセス起動禁止"
    `Shell` / `VBA.Shell` / `WScript.Shell.Run` / `WScript.Shell.Exec` 等の子プロセス起動は職場 PC では UAC 昇格や黙殺が発生して動かないため、本ツールでは **全面禁止** です（[ADR-0002](architecture.md#5-adr)）。ChromaDB 連携も外部プロセス起動ではなく、事前 ETL → Excel 内 Range 走査の方式を採ります。

## 8. テスト構成

| 区分 | 件数 | 実行方法 |
|---|---|---|
| 既存テスト（本番モジュール経由） | PASS 82 / SKIP 4 / FAIL 0 | `Application.Run "TestRunner_RunAll"` |
| image_ext rev1 拡張テスト | PASS 7 | `Application.Run "TestRunnerExt_RunAll"` |
| **合計** | **PASS 89 / SKIP 4 / FAIL 0** | 上記両方 |

### image_ext rev1 拡張テストの内訳

| ID | テスト | カテゴリ |
|---|---|---|
| T10-001 | ScoreMatch=0（FormatID 不一致） | T10 |
| T10-002 | ScoreMatch>0（キーワード一致） | T10 |
| T10-003 | ImagePath スタンザ明示指定 | T10 |
| T10-004 | ImagePath 既定値（`<KnwNo>.png`） | T10 |
| T10-005 | ResolveImageFolder = parent/kb_images | T10 |
| T11-001 | FormSpec.AddControl で 2 件追加 | T11 |
| T11-002 | BuildOnly: `designer.Controls.Count=2` | T11 |

拡張テスト本体（`modTestRunnerExt.bas`）は dev 専用で、配布物には含まれません。

## 9. v1 → v2 → v3 の差分

| 項目 | v1 | v2 | v3（最新） |
|---|---|---|---|
| `ThisWorkbook` の `Attribute VB_Base` | 欠落（Excel 実機でコンパイルエラー） | `0{00020819-0000-0000-C000-000000000046}` を注入 | 同左 |
| `clsSearchEngine` の `▶` 文字 | リテラル `"▶ 詳細"`（CP932 で `?` に化けていた） | `ChrW(&H25B6) & " 詳細"` で実行時構築 | 同左 |
| `Workbook_Open` 自動初期化 | なし | あり（シート未生成なら Setup、デモデータ未投入なら Seed） | 同左 |
| `modDemoSeeder` の `〜`（波ダッシュ） | あり | `-`（ASCII ハイフン）に置換 | 同左 |
| C-1〜C-3 修正 | — | — | 反映済 |

職場本番展開（`release_v2/`）には `clsSearchEngine` の修正のみが恒久反映されます。Workbook_Open 自動初期化と `modDemoSeeder` はデモ専用です。

## 10. バージョン

- **release v2** — rev21 ベース（M6-002/003/004 の検索結合セル対策、ボタン配置のセルアンカー対応を含む）
- **image_ext rev1（2026-05-04）** — 検索結果画像列追加 + UserForm spec 駆動生成基盤
- 配布物は **モジュールのみ** 構成。シート/ボタン生成は同梱の `modSetup.bas` の `SetupSheetsAndButtons` マクロで実施。
