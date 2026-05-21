---
title: clsSectionSpec.cls
---

# clsSectionSpec.cls

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | 画面内のセクション帯 1 個の仕様を保持する値オブジェクト |
| 行数 | 47 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsSectionSpec.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSectionSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsSectionSpec（画面層 ? データ）
' 概要:   1 個のセクション帯の仕様（位置・ラベル・色）を保持する ValueObject。
' 依存先: なし
' ================================================================

Private m_address As String
Private m_label As String
Private m_colorHex As String

Public Property Get Address() As String
    Address = m_address
End Property
Public Property Let Address(ByVal value As String)
    m_address = value
End Property

Public Property Get Label() As String
    Label = m_label
End Property
Public Property Let Label(ByVal value As String)
    m_label = value
End Property

Public Property Get ColorHex() As String
    ColorHex = m_colorHex
End Property
Public Property Let ColorHex(ByVal value As String)
    m_colorHex = value
End Property

Public Sub Init(ByVal address As String, ByVal label As String, ByVal colorHex As String)
    m_address = address
    m_label = label
    m_colorHex = colorHex
End Sub
```
