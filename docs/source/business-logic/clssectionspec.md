---
title: clsSectionSpec.cls
---

# clsSectionSpec.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | 1 画面内の 1 セクション (帯) の宣言情報を保持する値オブジェクト |
| 行数 | 47 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsSectionSpec` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

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
' クラス: clsSectionSpec（画面層 — データ）
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
