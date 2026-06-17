---
title: ユーザマニュアル｜外部設定ファイル
description: 動作設定ファイル・フォーマット定義ファイル・画面設定ファイルの置き場所と、設定項目・値をすべて説明します
---

# 外部設定ファイル

ナレッジ管理ツールは、保存先のフォルダや動作の設定を **外部のテキストファイル** で管理しています。
ファイルを書き換えるだけで、保存先や動作を変更できます。

このページでは、設定ファイルの種類・置き場所と、設定項目・値をすべて説明します。

## 設定ファイルの全体像

| 種類 | ファイル | 置き場所 | 役割 |
|---|---|---|---|
| 動作設定ファイル | `管理_config.txt` / `検索_config.txt` / `登録修正_config.txt` | `C:\KnowledgeMgr\` | 保存先フォルダや動作を決める |
| フォーマット定義ファイル | `（型ID）.txt`（例：`FAULT.txt`） | `C:\KnowledgeMgr\formats\` | 型ごとの項目（フィールド）を決める |
| 画面設定ファイル | 各画面のファイル | `C:\KnowledgeMgr\ui\（役）\` | 画面の見た目（列幅・色・項目名）を決める |

!!! warning "文字コードに注意"
    設定ファイルはすべて、メモ帳で **「ANSI」(Shift_JIS) 文字コード**・改行コード **CRLF** で保存してください。
    UTF-8 のまま保存すると、日本語の項目名や値が文字化けします。

## 動作設定ファイル（_config.txt）

保存先フォルダや動作を決めるファイルです。役（管理 / 検索 / 登録修正）ごとに 1 つずつあります。

### 置き場所

```text
C:\KnowledgeMgr\管理_config.txt
C:\KnowledgeMgr\検索_config.txt
C:\KnowledgeMgr\登録修正_config.txt
```

3 つのファイルは同じ項目を持ちます。通常はどれも同じ値にしておきます。
役ごとに動作を変えたいとき（たとえば管理だけ詳しい記録を残したいとき）だけ、値を変えます。

### 書き方

```text
[CONFIG]
debugLevel=INFO
logRotationRows=10000
data_dir=C:\KnowledgeMgr\data\
format_dir=C:\KnowledgeMgr\formats\
ui_dir=C:\KnowledgeMgr\ui\
backup_dir=C:\KnowledgeMgr\data\backup\
log_dir=C:\KnowledgeMgr\log\
config_dir=C:\KnowledgeMgr\
uiSchemaFailMode=safeDefault
```

- `設定項目=値` の形で 1 行に 1 項目を書きます。
- `[CONFIG]` のような `[ ]` で囲んだ見出し行は、区切りの目印です。
- `#` で始まる行は説明用のメモとして無視されます。
- ファイルに書かなかった項目は、下表の「既定値」が使われます。
- フォルダのパスは末尾に `\` を付けます。

### 設定項目の一覧

| 設定項目 | 意味 | 値の種類 | 既定値 | 例 |
|---|---|---|---|---|
| `data_dir` | ナレッジデータの保存先フォルダ | フォルダのパス | `C:\KnowledgeMgr\data\` | `C:\KnowledgeMgr\data\` |
| `format_dir` | フォーマット定義ファイルの置き場所 | フォルダのパス | `C:\KnowledgeMgr\formats\` | `C:\KnowledgeMgr\formats\` |
| `ui_dir` | 画面設定ファイルの置き場所 | フォルダのパス | `C:\KnowledgeMgr\ui\` | `C:\KnowledgeMgr\ui\` |
| `backup_dir` | バックアップの保存先フォルダ | フォルダのパス | `C:\KnowledgeMgr\backup\` | `C:\KnowledgeMgr\data\backup\` |
| `log_dir` | 操作ログの保存先フォルダ | フォルダのパス | `C:\KnowledgeMgr\log\` | `C:\KnowledgeMgr\log\` |
| `config_dir` | 設定ファイルそのものの置き場所 | フォルダのパス | `C:\KnowledgeMgr\` | `C:\KnowledgeMgr\` |
| `debugLevel` | 記録に残す詳しさ | 後述の 6 段階 | `ERROR` | `INFO` |
| `logRotationRows` | 操作ログ 1 ファイルに残す最大行数 | 数字 | `10000` | `10000` |
| `uiSchemaFailMode` | 画面設定ファイルに不備があったときの動き | `safeDefault` | `safeDefault` | `safeDefault` |

- **保存先フォルダ系**（`data_dir` / `format_dir` / `ui_dir` / `backup_dir` / `log_dir`）を変えると、その種類のファイルの読み書き先が変わります。先にフォルダを作ってから設定してください。
- **`config_dir`** は設定ファイル自身の場所です。通常は既定のままにします。
- **`logRotationRows`** で指定した行数を超えると、古い操作ログが切り替わります。
- **`uiSchemaFailMode`** の `safeDefault` は、画面設定ファイルに不備があっても、標準の見た目で画面を開く動きです。

### `debugLevel` の値

`debugLevel` は、記録（ログ）に残す詳しさを次の 6 段階から選びます。下にいくほど詳しく記録されます。

| 値 | 記録される内容 |
|---|---|
| `OFF` | 記録しない |
| `ERROR` | エラーだけ（既定値） |
| `WARN` | エラーと警告 |
| `INFO` | 上記＋おもな操作の記録 |
| `DEBUG` | 上記＋細かい動作の記録 |
| `TRACE` | 上記＋もっとも詳しい動作の記録 |

普段は `ERROR`（または `INFO`）のままで十分です。不具合を調べたいときだけ `DEBUG` や `TRACE` にします。
詳しく記録するほどログの量が増えるため、調査が終わったら元に戻してください。

### 古い設定項目について

ファイルによっては、次の項目が書かれていることがあります。これらは **現在のバージョンでは使われません**。
書かれていても害はありませんが、値を変えても動作は変わりません。

| 設定項目 | 補足 |
|---|---|
| `systemSheetVisibility` | 現在は使われません |
| `autoReloadOnStartup` | 現在は使われません |
| `migrateBackupEnabled` | 現在は使われません |

## フォーマット定義ファイル

型（フォーマット）ごとの項目を決めるファイルです。`format_dir`（既定では `C:\KnowledgeMgr\formats\`）に置きます。
ファイル名は型の ID に合わせます（例：`FAULT.txt`）。

通常は管理ブックの設計画面から作成・編集しますが、中身は次のようなテキストです。

### 書き方

```text
[FORMAT]
FormatId=FAULT
FormatName=障害対応
===
[FIELD]
FieldName=件名
FieldType=単一行
Required=TRUE
MaxLength=120
fieldPlaceholder=（例）経理システムへログインできない
searchTarget=TRUE
===
[FIELD]
FieldName=カテゴリ
FieldType=選択
Required=FALSE
DropdownOptions=アプリ障害|インフラ障害|ネットワーク障害|セキュリティ
fieldPlaceholder=（例）アプリ障害
searchTarget=FALSE
```

- 先頭の `[FORMAT]` で型そのものの情報を、続く `[FIELD]` で項目を 1 つずつ並べます。
- 項目どうしは `===` の行で区切ります。
- 上から書いた順に、登録画面・検索画面の項目の並び順になります。

### 型そのものの項目（[FORMAT]）

| 設定項目 | 意味 | 値の種類 | 例 |
|---|---|---|---|
| `FormatId` | 型の ID（ファイル名にもなる） | 英数字 | `FAULT` |
| `FormatName` | 画面に表示される型の名前 | 文字列 | `障害対応` |

### 各項目の属性（[FIELD]）

| 設定項目 | 意味 | 値の種類 | 既定の扱い | 例 |
|---|---|---|---|---|
| `FieldName` | 項目名（入力欄に表示される名前） | 文字列 | （必須） | `件名` |
| `FieldType` | 型 | `単一行` / `複数行` / `日付` / `選択` | （必須） | `単一行` |
| `Required` | 入力を必須にするか | `TRUE` / `FALSE` | `FALSE` | `TRUE` |
| `MaxLength` | 入力できる文字数の上限 | 数字（単一行で使用） | 指定なし＝上限なし | `120` |
| `Rows` | 複数行のときの入力欄の高さ | 数字（複数行で使用） | 指定なし＝標準の高さ | `4` |
| `Scroll` | 複数行で欄内スクロールを使うか | `TRUE` / `FALSE` | `FALSE` | `TRUE` |
| `DropdownOptions` | 選択のときの候補 | 候補を縦棒 `\|` で区切る | 指定なし | `低\|中\|高\|緊急` |
| `fieldPlaceholder` | 入力欄に薄く表示する見本 | 文字列 | 指定なし | `（例）山田 太郎` |
| `searchTarget` | 検索で本文として探すか | `TRUE` / `FALSE` | `FALSE` | `TRUE` |

各型の意味や選び方は、[フィールド型の説明](types.md)のページを参照してください。

## 画面設定ファイル

画面の見た目（シート名・列幅・色・項目名など）を決めるファイルです。
`ui_dir`（既定では `C:\KnowledgeMgr\ui\`）の、役ごとのフォルダに置きます。

```text
C:\KnowledgeMgr\ui\管理\
C:\KnowledgeMgr\ui\検索\
C:\KnowledgeMgr\ui\登録修正\
```

各ファイルの設定項目の詳しい一覧は、[画面設定ファイル一覧](../stanza/index.md)のページにまとめています。

## 関連ページ

- [フィールド型の説明](types.md)
- [フォーマット設計の流れ](format.md)
- [セットアップ](setup.md)
- [画面設定ファイル一覧](../stanza/index.md)
