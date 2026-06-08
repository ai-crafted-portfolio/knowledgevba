---
title: clsFormSpec.cls
description: clsFormSpec.cls 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# clsFormSpec.cls

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 繧ｯ繝ｩ繧ｹ繝｢繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`clsFormSpec.cls`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsFormSpec (ビジネスロジック層)
' 概要:   1 つの UserForm の宣言情報を保持。
'         FormTitle / Width / Height + Controls Collection。
'         AddControl で clsControlSpec を追加する DSL を提供。
' 依存先: clsControlSpec
' 備考:   modFormBuilder.BuildAndShow に渡して動的フォーム生成する。
' ================================================================

Private m_formTitle As String
Private m_width As Long
Private m_height As Long
Private m_controls As Collection

Private Sub Class_Initialize()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0463] clsFormSpec.Class_Initialize ENTER"  ' [ADR-0100]
    Set m_controls = New Collection
    m_width = 480
    m_height = 360
    m_formTitle = "Untitled"
End Sub

' --- FormTitle ---
Public Property Get FormTitle() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0464] clsFormSpec.FormTitle ENTER"  ' [ADR-0100]
    FormTitle = m_formTitle
End Property
Public Property Let FormTitle(ByVal v As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0465] clsFormSpec.FormTitle ENTER"  ' [ADR-0100]
    m_formTitle = v
End Property

' --- Width ---
Public Property Get Width() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0466] clsFormSpec.Width ENTER"  ' [ADR-0100]
    Width = m_width
End Property
Public Property Let Width(ByVal v As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0467] clsFormSpec.Width ENTER"  ' [ADR-0100]
    m_width = v
End Property

' --- Height ---
Public Property Get Height() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0468] clsFormSpec.Height ENTER"  ' [ADR-0100]
    Height = m_height
End Property
Public Property Let Height(ByVal v As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0469] clsFormSpec.Height ENTER"  ' [ADR-0100]
    m_height = v
End Property

' --- Controls (read-only Collection) ---
Public Property Get Controls() As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0470] clsFormSpec.Controls ENTER"  ' [ADR-0100]
    Set Controls = m_controls
End Property

' ================================================================
' 関数名: AddControl
' 概要:   clsControlSpec を生成して Controls Collection に追加する DSL。
'         同名 Control が既にあれば例外を投げる (Collection の挙動)。
' 引数:   cType / nm / l / t / w / h / cap / onClk - clsControlSpec.Init と同じ
' 戻り値: clsControlSpec - 追加された ControlSpec (チェイン用)
' ================================================================
Public Function AddControl(ByVal cType As String, ByVal nm As String, _
                            ByVal l As Long, ByVal t As Long, _
                            ByVal w As Long, ByVal h As Long, _
                            Optional ByVal cap As String = "", _
                            Optional ByVal onClk As String = "") As clsControlSpec
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0471] clsFormSpec.AddControl ENTER"  ' [ADR-0100]
    Dim cs As clsControlSpec
    Set cs = New clsControlSpec
    cs.Init cType, nm, l, t, w, h, cap, onClk
    m_controls.Add cs, nm
    Set AddControl = cs
End Function

' ================================================================
' 関数名: ControlByName
' 概要:   指定名の Control Spec を取得 (なければ Nothing)
' ================================================================
Public Function ControlByName(ByVal nm As String) As clsControlSpec
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0472] clsFormSpec.ControlByName ENTER"  ' [ADR-0100]
    On Error Resume Next
    Set ControlByName = m_controls(nm)
    On Error GoTo 0
End Function

' ================================================================
' 関数名: ControlCount
' 概要:   Controls Collection の件数
' ================================================================
Public Function ControlCount() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0473] clsFormSpec.ControlCount ENTER"  ' [ADR-0100]
    ControlCount = m_controls.Count
End Function
```