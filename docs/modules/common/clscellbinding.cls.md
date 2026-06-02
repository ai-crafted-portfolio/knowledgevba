---
title: clsCellBinding.cls
description: clsCellBinding.cls のソースコード（コピペ用）
---

# clsCellBinding.cls

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsCellBinding.cls`
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
End
Attribute VB_Name = "clsCellBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' クラス: clsCellBinding (Phase 3 / Bridge layer)
' 概要:   Cell 1 つに対する read/write を抽象化する thin wrapper
'         target は Worksheet (本番) または Scripting.Dictionary (L2-UI mock) を受理
'         Dict mode の key = cellAddr 文字列、value = String
' Version: v2.1 (2026-05-18 初版、test_spec_v2.1_ui_io.md §1.2 / §7 準拠)
' 依存先: なし (modConfigHolder.GetValueOrDefault は使わない、純粋セル I/O)
' 関連:   ADR-0053 §2.9 (外部 I/O 独立化)、ADR-0062 P3b (Cell 書込み hang 回避)
' Q-TBD-UI-006 決定: thin class として実装、modEntry* 内は内部 helper 使用、
'                    clsCellBinding は public 直接利用向け facade
' ================================================================
Option Explicit

Private m_target As Object
Private m_cellAddr As String
Private m_isValid As Boolean

' ================================================================
' Init: target (Worksheet or Dictionary) + cellAddr で初期化
' ================================================================
Public Sub Init(ByVal target As Object, ByVal cellAddr As String)
    Set m_target = target
    m_cellAddr = cellAddr
    m_isValid = (Not (target Is Nothing)) And (Len(cellAddr) > 0)
End Sub

Public Property Get CellAddr() As String
    CellAddr = m_cellAddr
End Property

Public Property Get IsValid() As Boolean
    IsValid = m_isValid
End Property

' ================================================================
' ReadValue: target から cellAddr の値を String で取得
'   Dict mode: target.Exists(cellAddr) なら値、なければ vbNullString
'   Worksheet mode: Range(cellAddr).Cells(1,1).Value を CStr
' ================================================================
Public Function ReadValue() As String
    If Not m_isValid Then
        ReadValue = vbNullString
        Exit Function
    End If

    On Error GoTo ErrHandler

    If TypeName(m_target) = "Dictionary" Then
        If m_target.Exists(m_cellAddr) Then
            ReadValue = CStr(m_target(m_cellAddr))
        Else
            ReadValue = vbNullString
        End If
    Else
        Dim r As Range
        Set r = m_target.Range(m_cellAddr)
        Dim v As Variant
        v = r.Cells(1, 1).Value
        If IsError(v) Then
            ' #N/A / #REF! 等、エラー値の文字列化 (Q-TBD-UI-004 デフォルト挙動)
            ReadValue = CStr(CVErr(v))
        ElseIf IsNull(v) Then
            ReadValue = vbNullString
        ElseIf IsDate(v) Then
            ' Date は ISO 8601 (yyyy-mm-dd) に正規化
            ReadValue = Format(v, "yyyy-mm-dd")
        Else
            ReadValue = CStr(v)
        End If
    End If
    Exit Function

ErrHandler:
    ReadValue = vbNullString
End Function

' ================================================================
' WriteValue: target の cellAddr に value (String) を書込み
'   Dict mode: target(cellAddr) = value
'   Worksheet mode: Range(cellAddr).Cells(1,1).Value = value
' ================================================================
Public Sub WriteValue(ByVal value As String)
    If Not m_isValid Then Exit Sub

    On Error GoTo ErrHandler

    If TypeName(m_target) = "Dictionary" Then
        m_target(m_cellAddr) = value
    Else
        Dim r As Range
        Set r = m_target.Range(m_cellAddr)
        r.Cells(1, 1).Value = value
    End If
    Exit Sub

ErrHandler:
    ' 上位に伝播させない (warn 等価、log は modEntry* 側で取る)
End Sub
```
