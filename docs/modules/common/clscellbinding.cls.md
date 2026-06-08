---
title: clsCellBinding.cls
description: clsCellBinding.cls 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# clsCellBinding.cls

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 繧ｯ繝ｩ繧ｹ繝｢繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`clsCellBinding.cls`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0402] clsCellBinding.Init ENTER"  ' [ADR-0100]
    Set m_target = target
    m_cellAddr = cellAddr
    m_isValid = (Not (target Is Nothing)) And (Len(cellAddr) > 0)
End Sub

Public Property Get CellAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0403] clsCellBinding.CellAddr ENTER"  ' [ADR-0100]
    CellAddr = m_cellAddr
End Property

Public Property Get IsValid() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0404] clsCellBinding.IsValid ENTER"  ' [ADR-0100]
    IsValid = m_isValid
End Property

' ================================================================
' ReadValue: target から cellAddr の値を String で取得
'   Dict mode: target.Exists(cellAddr) なら値、なければ vbNullString
'   Worksheet mode: Range(cellAddr).Cells(1,1).Value を CStr
' ================================================================
Public Function ReadValue() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0405] clsCellBinding.ReadValue ENTER"  ' [ADR-0100]
    If Not m_isValid Then
        ReadValue = vbNullString
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0406] clsCellBinding.ReadValue EXIT-OK"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0407] clsCellBinding.ReadValue EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0408] clsCellBinding.ReadValue EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ReadValue = vbNullString
End Function

' ================================================================
' WriteValue: target の cellAddr に value (String) を書込み
'   Dict mode: target(cellAddr) = value
'   Worksheet mode: Range(cellAddr).Cells(1,1).Value = value
' ================================================================
Public Sub WriteValue(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0409] clsCellBinding.WriteValue ENTER"  ' [ADR-0100]
    If Not m_isValid Then Exit Sub

    On Error GoTo ErrHandler

    If TypeName(m_target) = "Dictionary" Then
        m_target(m_cellAddr) = value
    Else
        Dim r As Range
        Set r = m_target.Range(m_cellAddr)
        r.Cells(1, 1).Value = value
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0410] clsCellBinding.WriteValue EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0411] clsCellBinding.WriteValue EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' 上位に伝播させない (warn 等価、log は modEntry* 側で取る)
End Sub
```