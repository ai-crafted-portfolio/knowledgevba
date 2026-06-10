---
title: 障害対応 (FAULT.txt)
description: サンプルデータ - 障害対応 (FAULT.txt)
---

# 障害対応 (FAULT.txt)

**配置先**: `C:\KnowledgeMgr\formats\FAULT.txt`

メモ帳で **ANSI** (Shift-JIS) として保存してください。改行は CRLF。

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
FieldName=発生日時
FieldType=日付
Required=TRUE
fieldPlaceholder=（例）2026-06-09 10:30
searchTarget=TRUE
===
[FIELD]
FieldName=担当者
FieldType=単一行
Required=FALSE
MaxLength=40
fieldPlaceholder=（例）山田 太郎
searchTarget=TRUE
===
[FIELD]
FieldName=カテゴリ
FieldType=選択
Required=FALSE
DropdownOptions=アプリ障害|インフラ障害|ネットワーク障害|セキュリティ
fieldPlaceholder=（例）アプリ障害
searchTarget=FALSE
===
[FIELD]
FieldName=優先度
FieldType=選択
Required=FALSE
DropdownOptions=低|中|高|緊急
fieldPlaceholder=（例）高
searchTarget=FALSE
===
[FIELD]
FieldName=概要
FieldType=複数行
Required=TRUE
Rows=4
fieldPlaceholder=（例）応答時間が通常の10倍以上に増加
searchTarget=FALSE
===
[FIELD]
FieldName=原因
FieldType=複数行
Required=FALSE
Rows=4
fieldPlaceholder=（例）長時間トランザクションの累積
searchTarget=FALSE
===
```
