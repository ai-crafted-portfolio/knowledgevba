---
title: clsKnowledgeEditScreen.cls
---

# clsKnowledgeEditScreen.cls

| 項目 | 値 |
|---|---|
| 層 | 画面層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | ナレッジ修正 (M-06) の ScreenSpec 構築 |
| 行数 | 38 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsKnowledgeEditScreen` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

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
' クラス: clsKnowledgeEditScreen（画面層）
' 概要:   M-06 ナレッジ修正画面の構築・再描画。
'         spec を modScreenRender に委譲してタイトル/セクション帯/ボタン/
'         フィールドラベル/一覧ヘッダ/←メインに戻る を一括描画する。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)
    Exit Sub
ErrHandler:
    Debug.Print "[clsKnowledgeEditScreen.Setup] ERROR: " & Err.Description
End Sub

Public Sub Render()
    Call Setup
End Sub

```
