---
title: 仕様
description: モジュール構成、検索スコアリング、画像解決規則、spec 駆動 UserForm DSL、テスト構成
tags:
  - excel
  - vba
  - spec
---

# 仕様

現バージョンの仕様を定義します。

!!! info "このページの読み方"
    本ページは「現状の本ツールが何をどう実現しているか」を網羅的に記載するリファレンスです。手順を知りたい場合は [セットアップ](setup.md) / [操作手順](operations.md) へ、層構造や配布パターンを知りたい場合は [アーキテクチャ](architecture.md) へジャンプしてください。

## 1. モジュール構成（48 個 + ThisWorkbook）

VBA モジュールは「層」ごとにフォルダ分割した形で配布されます。インポート対象は 48 個、`ThisWorkbook.cls` は中身コピペで合計 49 ファイルです。

| 層 | 個数 | 代表モジュール |
|---|---|---|
| インストーラ層 | 1 | `modSetup.bas` |
| エントリポイント層 | 6 | `modEntryMain.bas` / `modEntrySearch.bas` / `modEntryKnowledge.bas` / `modEntryFormat.bas` / `modEntrySettings.bas` / `modSpecExamples.bas` |
| ビジネスロジック層 | 21 | `clsSearchEngine.cls` / `clsKnowledgeManager.cls` / `clsFormatManager.cls` / `clsTaskController.cls` / `clsStorageResolver.cls` / `clsFieldMigrator.cls` / `clsLogger.cls` / `clsButtonSpec.cls` / `clsControlSpec.cls` / `clsFieldSpec.cls` / `clsFormSpec.cls` / `clsScreenSpec.cls` / `clsSectionSpec.cls` / `clsSetupOrchestrator.cls` / `clsSheetRenderer.cls` / `clsUserFormRenderer.cls` / `IScreenRenderer.cls` / `modFactory.bas` / `modFormBuilder.bas` / `modScreenRender.bas` / `modScreenSpecRegistry.bas` |
| ユーティリティ層 | 6 | `modCommon.bas` / `modStringUtil.bas` / `modDateUtil.bas` / `modFileIO.bas` / `modImageRender.bas` / `clsLogEntry.cls` |
| 画面層 | 14 | `clsMainScreen.cls`〜`clsLogScreen.cls`（M-01〜M-14 の各画面構築クラス） |
| 特殊モジュール | 1 | `ThisWorkbook.cls` |

全モジュールの詳細とソースコードは [ソースコード一覧](source/index.md) を参照してください。

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

セットアップマクロは同時に **フォームコントロールボタン 68 個** を全 14 シートに配置・マクロ割り当てまで実施します（メイン 12 個 + 業務 13 シート計 56 個）。

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

## 7. 検索バックエンド差替境界

本実装は `<dataFolder>/*.txt` を ADODB で SJIS 読込する方式です。`clsSearchEngine.cls` 冒頭のコメントブロックに、別のデータソース（事前 ETL で `Sheet "Data"` に export しておくなど）に差し替える場合の指針が記載されています。具体的には:

1. 事前 ETL で chunks を任意のシート（例: `Data`）に export しておく
2. `ScanAndMatch` の txt ループ箇所を Range 走査に置換する

この変更点は `clsSearchEngine` 1 ファイルに局所化されており、上位エントリポイントには影響しません。

!!! danger "VBA からの子プロセス起動は使わない"
    `Shell` / `VBA.Shell` / `WScript.Shell.Run` / `WScript.Shell.Exec` 等の子プロセス起動は使いません。外部データ連携が必要な場合も、外部プロセス起動ではなく事前 ETL → Excel 内 Range 走査の方式を採ります。

## 8. 既知の制約・トレードオフ

### 制約

- **検索アルゴリズムは単純スコアリング** — タイトル ×3 + 対象フィールド ×2 + 出現回数の加算式で、意味検索（ベクトル類似度）や同義語展開は対象外です
- **`clsFormSpec` DSL のコントロール種別は 4 種** — `Label` / `TextBox` / `Image` / `Button` のみ。`ListBox` / `ComboBox` 等の追加は将来検討（[§5.2](#52-サポートしているコントロール種別)）
- **クラスモジュール内 `Public Const` 禁止** — VBA 仕様の制約。列番号などの共通定数は `modCommon.bas` に集約しています
- **モバイル / Excel for Mac 不可** — `Forms.*` ProgID が無いため `Macro_ShowSearchResultPreview` は Windows 版 Excel でのみ動作します

### 設計上のトレードオフ

- **txt フラットファイル方式 vs DB 方式** — DB（SQLite 等）採用ならばインデックス検索ができる代わりに、ファイル単位の閲覧性・テキストエディタでの直接編集容易性が落ちます。閲覧性を取って `.txt` を採用しています
- **動的 UserForm 生成 vs 固定 UserForm** — 動的生成は Excel のオブジェクトモデルへの信頼設定が必要ですが、コントロール定義をコードで管理できる利点があります
