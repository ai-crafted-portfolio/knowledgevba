---
title: clsFieldSpec.cls
description: clsFieldSpec.cls のソースコード（コピペ用）
---

# clsFieldSpec.cls

**配置先**: `共通モジュール (3 ブック全て)` 用の VBA モジュール  
**種類**: クラス モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\`
- ファイル名: `clsFieldSpec.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

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
Private m_hintText As String
Private m_orderAddr As String     ' 例: "A8"
Private m_reqMarkAddr As String   ' 例: "B8"
Private m_labelAddr As String     ' 例: "C8"
Private m_typeAddr As String      ' 例: "D8"
Private m_inputAddr As String     ' 例: "E8"

Public Property Get FieldOrder() As Long
    FieldOrder = m_order
End Property
Public Property Let FieldOrder(ByVal value As Long)
    m_order = value
End Property

Public Property Get Label() As String
    Label = m_label
End Property
Public Property Let Label(ByVal value As String)
    m_label = value
End Property

Public Property Get TypeText() As String
    TypeText = m_typeText
End Property
Public Property Let TypeText(ByVal value As String)
    m_typeText = value
End Property

Public Property Get Required() As Boolean
    Required = m_required
End Property
Public Property Let Required(ByVal value As Boolean)
    m_required = value
End Property

Public Property Get RowCount() As Long
    RowCount = m_rowCount
End Property
Public Property Let RowCount(ByVal value As Long)
    m_rowCount = value
End Property

Public Property Get HintText() As String
    HintText = m_hintText
End Property
Public Property Let HintText(ByVal value As String)
    m_hintText = value
End Property

Public Property Get OrderAddr() As String
    OrderAddr = m_orderAddr
End Property
Public Property Let OrderAddr(ByVal value As String)
    m_orderAddr = value
End Property

Public Property Get ReqMarkAddr() As String
    ReqMarkAddr = m_reqMarkAddr
End Property
Public Property Let ReqMarkAddr(ByVal value As String)
    m_reqMarkAddr = value
End Property

Public Property Get LabelAddr() As String
    LabelAddr = m_labelAddr
End Property
Public Property Let LabelAddr(ByVal value As String)
    m_labelAddr = value
End Property

Public Property Get TypeAddr() As String
    TypeAddr = m_typeAddr
End Property
Public Property Let TypeAddr(ByVal value As String)
    m_typeAddr = value
End Property

Public Property Get InputAddr() As String
    InputAddr = m_inputAddr
End Property
Public Property Let InputAddr(ByVal value As String)
    m_inputAddr = value
End Property

Public Sub Init(ByVal fieldOrder As Long, _
                ByVal label As String, _
                ByVal typeText As String, _
                ByVal required As Boolean, _
                ByVal rowCount As Long, _
                ByVal hintText As String)
    m_order = fieldOrder
    m_label = label
    m_typeText = typeText
    m_required = required
    m_rowCount = rowCount
    m_hintText = hintText
End Sub

Public Sub SetCellAddrs(ByVal orderAddr As String, _
                        ByVal reqMarkAddr As String, _
                        ByVal labelAddr As String, _
                        ByVal typeAddr As String, _
                        ByVal inputAddr As String)
    m_orderAddr = orderAddr
    m_reqMarkAddr = reqMarkAddr
    m_labelAddr = labelAddr
    m_typeAddr = typeAddr
    m_inputAddr = inputAddr
End Sub
```
