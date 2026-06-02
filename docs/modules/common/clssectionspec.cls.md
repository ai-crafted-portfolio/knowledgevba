---
title: clsSectionSpec.cls
description: clsSectionSpec.cls のソースコード（コピペ用）
---

# clsSectionSpec.cls

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsSectionSpec.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
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
