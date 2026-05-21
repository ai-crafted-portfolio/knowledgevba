---
title: clsButtonSpec.cls
---

# clsButtonSpec.cls

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | ボタン 1 個分の仕様を保持する値オブジェクト |
| 行数 | 96 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsButtonSpec.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsButtonSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsButtonSpec（画面層 ? データ）
' 概要:   1 個のボタンの仕様（位置・キャプション・色・マクロ）を保持する
'         ValueObject。コードから分離してデータ駆動で画面を構築する。
' 依存先: なし（ValueObject）
' ================================================================

Private m_btnName As String
Private m_caption As String
Private m_cellAddr As String
Private m_colorHex As String
Private m_groupName As String
Private m_hintAddr As String
Private m_hintText As String

' --- Property Get/Let ---
Public Property Get BtnName() As String
    BtnName = m_btnName
End Property
Public Property Let BtnName(ByVal value As String)
    m_btnName = value
End Property

Public Property Get Caption() As String
    Caption = m_caption
End Property
Public Property Let Caption(ByVal value As String)
    m_caption = value
End Property

Public Property Get CellAddr() As String
    CellAddr = m_cellAddr
End Property
Public Property Let CellAddr(ByVal value As String)
    m_cellAddr = value
End Property

Public Property Get ColorHex() As String
    ColorHex = m_colorHex
End Property
Public Property Let ColorHex(ByVal value As String)
    m_colorHex = value
End Property

Public Property Get GroupName() As String
    GroupName = m_groupName
End Property
Public Property Let GroupName(ByVal value As String)
    m_groupName = value
End Property

Public Property Get HintAddr() As String
    HintAddr = m_hintAddr
End Property
Public Property Let HintAddr(ByVal value As String)
    m_hintAddr = value
End Property

Public Property Get HintText() As String
    HintText = m_hintText
End Property
Public Property Let HintText(ByVal value As String)
    m_hintText = value
End Property

' ================================================================
' 関数名: Init
' 概要:   一括初期化（Builder 風）
' 引数:   btnName, caption, cellAddr, colorHex, groupName, hintAddr, hintText
' ================================================================
Public Sub Init(ByVal btnName As String, _
                ByVal caption As String, _
                ByVal cellAddr As String, _
                ByVal colorHex As String, _
                Optional ByVal groupName As String = "", _
                Optional ByVal hintAddr As String = "", _
                Optional ByVal hintText As String = "")
    m_btnName = btnName
    m_caption = caption
    m_cellAddr = cellAddr
    m_colorHex = colorHex
    m_groupName = groupName
    m_hintAddr = hintAddr
    m_hintText = hintText
End Sub
```
