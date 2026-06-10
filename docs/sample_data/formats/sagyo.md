---
title: 作業手順書 (SAGYO.txt)
description: サンプルデータ - 作業手順書 (SAGYO.txt)
---

# 作業手順書 (SAGYO.txt)

**配置先**: `C:\KnowledgeMgr\formats\SAGYO.txt`

メモ帳で **ANSI** (Shift-JIS) として保存してください。改行は CRLF。

```text
[FORMAT]
FormatId=SAGYO
FormatName=作業手順書
===
[FIELD]
FieldName=手順書番号
FieldType=単一行
Required=TRUE
MaxLength=20
fieldPlaceholder=（例）OPS-001
searchTarget=TRUE
===
[FIELD]
FieldName=作業名
FieldType=単一行
Required=TRUE
MaxLength=120
fieldPlaceholder=（例）サーバ室UPS月次点検
searchTarget=TRUE
===
[FIELD]
FieldName=対象設備
FieldType=単一行
Required=TRUE
MaxLength=120
fieldPlaceholder=（例）サーバ室1F UPS本体
searchTarget=TRUE
===
[FIELD]
FieldName=作業区分
FieldType=選択
Required=FALSE
DropdownOptions=月次点検|週次点検|日次点検|緊急対応
fieldPlaceholder=（例）月次点検
searchTarget=FALSE
===
[FIELD]
FieldName=危険度
FieldType=選択
Required=FALSE
DropdownOptions=低|中|高
fieldPlaceholder=（例）中
searchTarget=FALSE
===
[FIELD]
FieldName=作業予定日時
FieldType=日付
Required=FALSE
fieldPlaceholder=（例）2026-07-15 22:00
searchTarget=FALSE
===
[FIELD]
FieldName=必要な資材
FieldType=複数行
Required=FALSE
Rows=3
fieldPlaceholder=（例）電気絶縁手袋一式
searchTarget=FALSE
===
[FIELD]
FieldName=作業手順
FieldType=複数行
Required=TRUE
Rows=8
fieldPlaceholder=（例）1. 事前に通知 2. 自己診断LED確認 3. 電圧測定
searchTarget=FALSE
===
[FIELD]
FieldName=使用する工具・部品
FieldType=単一行
Required=FALSE
MaxLength=120
fieldPlaceholder=（例）デジタルテスタ、点検表
searchTarget=FALSE
===
[FIELD]
FieldName=安全上の注意
FieldType=複数行
Required=FALSE
Rows=4
fieldPlaceholder=（例）必ず2名以上で作業すること
searchTarget=FALSE
===
[FIELD]
FieldName=確認者
FieldType=単一行
Required=FALSE
MaxLength=40
fieldPlaceholder=（例）三谷 太郎
searchTarget=FALSE
===
```
