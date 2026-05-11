---
title: アーキテクチャ
description: 層分離図、依存関係、配布パターン、関連 ADR
tags:
  - excel
  - vba
  - architecture
---

# アーキテクチャ

## 1. 層分離

VBA モジュールは責務ごとに 5 層に分離されています。依存方向は上から下への一方向のみで、ユーティリティ層がビジネスロジック層を呼ぶことはありません。

!!! tip "層分離の狙い"
    上位の層で発生した UI イベントを、薄いアダプタ（エントリポイント層）でドメインロジック（ビジネスロジック層）に橋渡しし、副作用の無い純粋関数（ユーティリティ層）を共有資産として配置する構造です。これにより、検索バックエンドの差し替えや新フォーマット追加など、変更が必要なファイルを局所化できます。

```mermaid
flowchart TB
    subgraph Entry [エントリポイント層]
        EM[modEntryMain]
        ES[modEntrySearch]
        EK[modEntryKnowledge]
        EF[modEntryFormat]
        ESt[modEntrySettings]
        EX[modSpecExamples]
    end
    subgraph Biz [ビジネスロジック層]
        SE[clsSearchEngine]
        KM[clsKnowledgeManager]
        FM[clsFormatManager]
        TC[clsTaskController]
        SR[clsStorageResolver]
        FMig[clsFieldMigrator]
        LG[clsLogger]
        FS[clsFormSpec]
        CS[clsControlSpec]
        FB[modFormBuilder]
    end
    subgraph Util [ユーティリティ層]
        UC[modCommon]
        US[modStringUtil]
        UD[modDateUtil]
        UF[modFileIO]
        UI[modImageRender]
        ULE[clsLogEntry]
    end
    subgraph Special [特殊]
        TW[ThisWorkbook]
    end
    subgraph Inst [インストーラ層]
        SU[modSetup]
    end

    Entry --> Biz
    Biz --> Util
    TW --> Inst
    TW --> Entry
    SU --> Util
    FB --> CS
    FS --> CS
```

各層の責務は次のとおり。

| 層 | 役割 | 典型ファイル |
|---|---|---|
| エントリポイント層 | フォームコントロールボタンから直接呼ばれる Public Sub。引数なしまたは `ByVal` 単純型のみ受け取り、ビジネスロジック層に委譲する | `modEntry*.bas` |
| ビジネスロジック層 | クラス中心。検索、ナレッジ管理、フォーマット定義、タスク制御、ストレージ解決、フィールド移行、ログ、UserForm 動的生成 | `cls*.cls` + `modFormBuilder.bas` |
| ユーティリティ層 | 純粋関数とヘルパ。日付、文字列、ファイル I/O、画像描画。**上位層への依存ゼロ** | `mod*Util.bas` + `modCommon.bas` + `modImageRender.bas` |
| インストーラ層 | 1 回限りのセットアップ処理。シート 14 個と 68 ボタンを実行時生成 | `modSetup.bas` |
| 特殊モジュール | `ThisWorkbook`。`Workbook_Open` で初回起動時のみ自動初期化を起動 | `ThisWorkbook.cls` |

## 2. 依存方向の制約

- **エントリポイント → ビジネスロジック → ユーティリティ** の一方向のみ
- **ユーティリティ層から上位層を呼ぶことを禁止**（循環依存防止）
- **`ThisWorkbook` は例外的にインストーラ層を呼ぶ**（初回起動時のセットアップ起動のため）
- **`modFormBuilder` のみ `VBProject.VBComponents` を参照**（spec 駆動 UserForm 生成のため。他モジュールはオブジェクトモデル参照を持たない）

## 3. 配布パターン

### 3.1 「.xlsm 本体は配布しない、モジュールだけ配布する」アプローチ

完成した `.xlsm` をそのまま配布する案には次の問題があります。

1. ユーザがダウンロード時に Excel の保護ビューで開きマクロが既定で無効化される
2. 配布物の `.xlsm` を更新するたびにユーザが手動でファイル差し替える必要がある
3. 開発側が事前に「シート構造 + ボタン配置」を作り込んだ `.xlsm` を毎回保守する必要がある（フォームコントロールを Python 側から書き込むと Excel の修復ダイアログを誘発する）

そこで本ツールは **`.bas`/`.cls` モジュール一式を ZIP で配布し、ユーザは新規 `.xlsm` に VBE から手動 import → セットアップマクロを 1 回実行する** 配布アーキテクチャを採用しています。

### 3.2 セットアップフロー

```mermaid
sequenceDiagram
    participant User as 利用者
    participant Excel as 新規 .xlsm
    participant VBE as VBE
    participant Setup as modSetup.RunSetup

    User->>Excel: 新規 .xlsm 作成（マクロ有効）
    User->>VBE: Alt+F11 で開く
    User->>VBE: 48 モジュールを Import
    User->>VBE: ThisWorkbook の中身をコピペ
    User->>VBE: コンパイル確認
    User->>Excel: Alt+F8 → SetupSheetsAndButtons
    Excel->>Setup: 実行
    Setup->>Excel: 14 シート生成
    Setup->>Excel: 68 ボタン配置 + マクロ割当
    Setup->>Excel: 初期可視性設定（メインのみ表示）
    Setup-->>User: 「セットアップ完了。」MsgBox
    User->>Excel: Ctrl+S 保存 → 閉じる → 再オープン
    Excel->>Excel: Workbook_Open でメインシート表示
```

### 3.3 利点と欠点

**利点**

- 配布物に `.xlsm` が含まれないため、Excel 修復ダイアログ問題から完全に解放
- モジュール単位で diff が見えるため、変更レビューが容易
- ユーザが既存の自分の `.xlsm` にも import できる（複数の `.xlsm` への展開が容易）
- 開発側で「シート構造 + ボタン配置を作り込んだ template `.xlsm`」を保守せずに済む

**欠点**

- ユーザが初回セットアップで VBE を開いて手動 import する手間
- VBE を開く操作に不慣れなユーザ向けに丁寧な手順書（[操作手順](operations.md)）が必要
- セットアップマクロが何らかの理由で失敗した場合の状態が中途半端になり得る（途中までシートだけ生成された等）

## 4. データフロー

```mermaid
flowchart LR
    subgraph FS [ファイルシステム]
        Txt[(dataFolder /<br/>*.txt SJIS+CRLF)]
        Img[(kb_images /<br/>KnwNo.png)]
    end
    subgraph Excel [Excel プロセス内]
        Search[検索シート]
        SE[clsSearchEngine]
        KL[ナレッジ一覧]
        FM[clsFormatManager]
        Disp[ナレッジ表示]
        Form[UserForm<br/>spec 駆動]
        Log[ログシート]
    end

    Search -->|キーワード| SE
    SE -->|.txt 走査| Txt
    SE -->|FormatID 照合| FM
    SE -->|Score 降順| Search
    Search -->|サムネ取得| Img
    Search -->|行選択| Form
    Form -->|詳細画像| Img
    Disp -->|単票表示| KL
    Disp -->|詳細画像| Img
    SE -->|ログ出力| Log
```

## 5. 主な設計上の前提

本ツールが選択した設計上の主な前提は次のとおりです。

| 前提 | 内容 |
|---|---|
| VBA 子プロセス起動は使わない | `Shell` / `WScript.Shell.Run` 等は使わず、外部ツール連携が必要な機能はすべて Excel プロセス内で代替実装 |
| `.xlsm` テンプレート配布ではなく モジュール配布 + セルフセットアップ | `.xlsm` テンプレート配布は Excel 修復ダイアログ問題を誘発するため、モジュール配布 + 初回セットアップマクロ方式を採用 |
| 画面 spec 駆動 | 各業務画面（M-01〜M-14）は `clsScreenSpec` の宣言情報を `clsSheetRenderer` が物理配置することで、画面追加時の差分を 1 spec 定義に局所化 |

## 6. 拡張ポイント

将来の変更時に触る予定のファイルを限定するための境界:

- **検索バックエンド差替** — `clsSearchEngine.cls` 1 ファイルのみ差し替え（txt 走査 → Range 走査などへ）
- **新フォーマット追加** — `clsFormatManager` のフォーマット定義テーブルを追加、`clsFieldMigrator` で旧→新の写像を定義
- **新しい UserForm を spec 駆動で追加** — `modSpecExamples.bas` を雛形にして `clsFormSpec.AddControl` で宣言、`modFormBuilder.BuildAndShow` で起動
- **新シート追加** — `modScreenSpecRegistry.bas` に画面 spec を 1 件追加、`modFactory` 経由で対応する画面層クラスを生成

## 7. 既知の制約

- **VBA 子プロセス禁止** — 全モジュール共通の規約。`Shell` / `WScript.Shell.Run` / `CreateObject("WScript.Shell").Run` 等は使いません
- **層は一方向依存** — エントリポイント → ビジネスロジック → ユーティリティ。逆方向の参照は循環依存を防ぐため禁止しています
- **`ThisWorkbook` だけはインストーラ層を呼んでよい** — 初回起動時のセットアップ起動という性質上、特例として認めています
- **VBE Import 時の改行コード** — `.bas` / `.cls` ファイルが LF だと Excel VBE が正しく取り込めない場合があります。配布パッケージは CRLF で固定しています
