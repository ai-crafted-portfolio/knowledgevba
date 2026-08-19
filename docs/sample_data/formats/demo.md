---
title: 新入力型デモ (DEMO.txt)
description: サンプルデータ - 新入力型デモ (DEMO.txt)
---

# 新入力型デモ (DEMO.txt)

**配置先**: `C:\KnowledgeMgr\formats\DEMO.txt`

メモ帳で **ANSI** (Shift-JIS) として保存してください。改行は CRLF。

新しく追加された入力型 (複数行（リッチ） / 複数行（コピー） / ファイル（リンク）) を
まとめて体験できるサンプルフォーマットです。

```text
[FORMAT]
FormatId=DEMO
FormatName=新入力型デモ
===
[FIELD]
FieldName=件名
FieldType=単一行
Required=TRUE
MaxLength=120
fieldPlaceholder=（例）サービス復旧手順
searchTarget=TRUE
===
[FIELD]
FieldName=手順書
FieldType=複数行リッチ
Required=FALSE
Rows=12
fieldPlaceholder=（例）# 概要 と書くと見出しになります
searchTarget=FALSE
===
[FIELD]
FieldName=定型コマンド
FieldType=複数行コピー
Required=FALSE
Rows=4
fieldPlaceholder=（例）よく使うコマンドを貼っておくと表示画面からまとめてコピーできます
searchTarget=FALSE
===
[FIELD]
FieldName=添付
FieldType=ファイルリンク
Required=FALSE
Rows=4
fieldPlaceholder=（例）［ファイル選択］ボタンで選ぶとパスが1行ずつ入ります
searchTarget=FALSE
===
[FIELD]
FieldName=確認者
FieldType=単一行
Required=FALSE
MaxLength=40
fieldPlaceholder=（例）山田 太郎
searchTarget=FALSE
```
