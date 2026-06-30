---
title: clsFieldSpec.cls
description: clsFieldSpec.cls のソースコード（コピペ用）
---

# clsFieldSpec.cls

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: クラスモジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsFieldSpec.cls`
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
Attribute VB_Name = "clsFieldSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsFieldSpec（画面層 ? データ）
' 概要:   1 個の入力フィールドの仕様（順序・ラベル・型・必須・行数）を
'         保持する ValueObject。データが空でもラベルは常時表示する根拠。
' 依存先: なし
' ================================================================

Private m_order As Long
Private m_label As String
Private m_typeText As String
Private m_required As Boolean
Private m_rowCount As Long
Private m_scrollEnabled As Boolean
Private m_hintText As String
Private m_orderAddr As String     ' 例: "A8"
Private m_reqMarkAddr As String   ' 例: "B8"
Private m_labelAddr As String     ' 例: "C8"
Private m_typeAddr As String      ' 例: "D8"
Private m_inputAddr As String     ' 例: "E8"

Public Property Get FieldOrder() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0438] clsFieldSpec.FieldOrder ENTER"  ' [ADR-0100]
    FieldOrder = m_order
End Property
Public Property Let FieldOrder(ByVal value As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0439] clsFieldSpec.FieldOrder ENTER"  ' [ADR-0100]
    m_order = value
End Property

Public Property Get Label() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0440] clsFieldSpec.Label ENTER"  ' [ADR-0100]
    Label = m_label
End Property
Public Property Let Label(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0441] clsFieldSpec.Label ENTER"  ' [ADR-0100]
    m_label = value
End Property

Public Property Get TypeText() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0442] clsFieldSpec.TypeText ENTER"  ' [ADR-0100]
    TypeText = m_typeText
End Property
Public Property Let TypeText(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0443] clsFieldSpec.TypeText ENTER"  ' [ADR-0100]
    m_typeText = value
End Property

Public Property Get Required() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0444] clsFieldSpec.Required ENTER"  ' [ADR-0100]
    Required = m_required
End Property
Public Property Let Required(ByVal value As Boolean)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0445] clsFieldSpec.Required ENTER"  ' [ADR-0100]
    m_required = value
End Property

Public Property Get RowCount() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0446] clsFieldSpec.RowCount ENTER"  ' [ADR-0100]
    RowCount = m_rowCount
End Property
Public Property Let RowCount(ByVal value As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0447] clsFieldSpec.RowCount ENTER"  ' [ADR-0100]
    m_rowCount = value
End Property

Public Property Get ScrollEnabled() As Boolean
    ScrollEnabled = m_scrollEnabled
End Property

Public Property Let ScrollEnabled(ByVal value As Boolean)
    m_scrollEnabled = value
End Property

Public Property Get HintText() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0448] clsFieldSpec.HintText ENTER"  ' [ADR-0100]
    HintText = m_hintText
End Property
Public Property Let HintText(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0449] clsFieldSpec.HintText ENTER"  ' [ADR-0100]
    m_hintText = value
End Property

Public Property Get OrderAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0450] clsFieldSpec.OrderAddr ENTER"  ' [ADR-0100]
    OrderAddr = m_orderAddr
End Property
Public Property Let OrderAddr(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0451] clsFieldSpec.OrderAddr ENTER"  ' [ADR-0100]
    m_orderAddr = value
End Property

Public Property Get ReqMarkAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0452] clsFieldSpec.ReqMarkAddr ENTER"  ' [ADR-0100]
    ReqMarkAddr = m_reqMarkAddr
End Property
Public Property Let ReqMarkAddr(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0453] clsFieldSpec.ReqMarkAddr ENTER"  ' [ADR-0100]
    m_reqMarkAddr = value
End Property

Public Property Get LabelAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0454] clsFieldSpec.LabelAddr ENTER"  ' [ADR-0100]
    LabelAddr = m_labelAddr
End Property
Public Property Let LabelAddr(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0455] clsFieldSpec.LabelAddr ENTER"  ' [ADR-0100]
    m_labelAddr = value
End Property

Public Property Get TypeAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0456] clsFieldSpec.TypeAddr ENTER"  ' [ADR-0100]
    TypeAddr = m_typeAddr
End Property
Public Property Let TypeAddr(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0457] clsFieldSpec.TypeAddr ENTER"  ' [ADR-0100]
    m_typeAddr = value
End Property

Public Property Get InputAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0458] clsFieldSpec.InputAddr ENTER"  ' [ADR-0100]
    InputAddr = m_inputAddr
End Property
Public Property Let InputAddr(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0459] clsFieldSpec.InputAddr ENTER"  ' [ADR-0100]
    m_inputAddr = value
End Property

Public Sub Init(ByVal fieldOrder As Long, _
                ByVal label As String, _
                ByVal typeText As String, _
                ByVal required As Boolean, _
                ByVal rowCount As Long, _
                ByVal hintText As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0460] clsFieldSpec.Init ENTER"  ' [ADR-0100]
    m_order = fieldOrder
    m_label = label
    m_typeText = typeText
    m_required = required
    m_rowCount = rowCount
    m_hintText = hintText
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0461] clsFieldSpec.Init EXIT-OK"  ' [ADR-0100]
End Sub

Public Sub SetCellAddrs(ByVal orderAddr As String, _
                        ByVal reqMarkAddr As String, _
                        ByVal labelAddr As String, _
                        ByVal typeAddr As String, _
                        ByVal inputAddr As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0462] clsFieldSpec.SetCellAddrs ENTER"  ' [ADR-0100]
    m_orderAddr = orderAddr
    m_reqMarkAddr = reqMarkAddr
    m_labelAddr = labelAddr
    m_typeAddr = typeAddr
    m_inputAddr = inputAddr
End Sub
```
