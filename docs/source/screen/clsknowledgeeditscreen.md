---
title: clsKnowledgeEditScreen.cls
---

# clsKnowledgeEditScreen.cls

| 項目 | 内容 |
|---|---|
| 層 | 画面層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 登録修正.xlsm |
| 役割 | M-06 ナレッジ修正画面の構築・再描画 |
| 行数 | 92 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsKnowledgeEditScreen.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKnowledgeEditScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsKnowledgeEditScreen (画面層 - M-06 ナレッジ修正)
' 概要:   M-06 画面の構築・再描画 + Btn handler 経路の logId emit。
'         spec を modScreenRender に委譲する標準パターン。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 配置:   S-登 (登録修正.xlsm)、placement_v2 §45
' 備考:   v2.1 (2026-05-20) で publish 投入。buildtest archive 旧版
'         (45L、LOG-M06-SCREENCLS-SETUP-*) を踏襲しつつ logging 規約
'         HandleButton メソッドは削除済 (2026-05-20)。ボタン処理の正規実装は modEntry* の Btn_*_v21