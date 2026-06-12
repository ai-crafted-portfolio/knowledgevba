---
title: clsSectionSpec.cls
description: clsSectionSpec.cls のソースコード（コピペ用）
---

# clsSectionSpec.cls

**配置先**: 共通モジュール（3 ブック共通）
**種類**: クラスモジュール
**更新日**: 2026-06-03 23:22

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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0680] clsSectionSpec.Address ENTER"  ' [ADR-0100]
    Address = m_address
End Property
Public Property Let Address(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0681] clsSectionSpec.Address ENTER"  ' [ADR-0100]
    m_address = value
End Property

Public Property Get Label() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0682] clsSectionSpec.Label ENTER"  ' [ADR-0100]
    Label = m_label
End Property
Public Property Let Label(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0683] clsSectionSpec.Label ENTER"  ' [ADR-0100]
    m_label = value
End Property

Public Property Get ColorHex() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0684] clsSectionSpec.ColorHex ENTER"  ' [ADR-0100]
    ColorHex = m_colorHex
End Property
Public Property Let ColorHex(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0685] clsSectionSpec.ColorHex ENTER"  ' [ADR-0100]
    m_colorHex = value
End Property

Public Sub Init(ByVal address As String, ByVal label As String, ByVal colorHex As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0686] clsSectionSpec.Init ENTER"  ' [ADR-0100]
    m_address = address
    m_label = label
    m_colorHex = colorHex
End Sub
```
