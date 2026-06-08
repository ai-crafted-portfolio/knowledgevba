---
title: clsButtonSpec.cls
description: clsButtonSpec.cls 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# clsButtonSpec.cls

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 繧ｯ繝ｩ繧ｹ繝｢繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`clsButtonSpec.cls`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
```vb
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0378] clsButtonSpec.BtnName ENTER"  ' [ADR-0100]
    BtnName = m_btnName
End Property
Public Property Let BtnName(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0379] clsButtonSpec.BtnName ENTER"  ' [ADR-0100]
    m_btnName = value
End Property

Public Property Get Caption() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0380] clsButtonSpec.Caption ENTER"  ' [ADR-0100]
    Caption = m_caption
End Property
Public Property Let Caption(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0381] clsButtonSpec.Caption ENTER"  ' [ADR-0100]
    m_caption = value
End Property

Public Property Get CellAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0382] clsButtonSpec.CellAddr ENTER"  ' [ADR-0100]
    CellAddr = m_cellAddr
End Property
Public Property Let CellAddr(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0383] clsButtonSpec.CellAddr ENTER"  ' [ADR-0100]
    m_cellAddr = value
End Property

Public Property Get ColorHex() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0384] clsButtonSpec.ColorHex ENTER"  ' [ADR-0100]
    ColorHex = m_colorHex
End Property
Public Property Let ColorHex(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0385] clsButtonSpec.ColorHex ENTER"  ' [ADR-0100]
    m_colorHex = value
End Property

Public Property Get GroupName() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0386] clsButtonSpec.GroupName ENTER"  ' [ADR-0100]
    GroupName = m_groupName
End Property
Public Property Let GroupName(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0387] clsButtonSpec.GroupName ENTER"  ' [ADR-0100]
    m_groupName = value
End Property

Public Property Get HintAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0388] clsButtonSpec.HintAddr ENTER"  ' [ADR-0100]
    HintAddr = m_hintAddr
End Property
Public Property Let HintAddr(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0389] clsButtonSpec.HintAddr ENTER"  ' [ADR-0100]
    m_hintAddr = value
End Property

Public Property Get HintText() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0390] clsButtonSpec.HintText ENTER"  ' [ADR-0100]
    HintText = m_hintText
End Property
Public Property Let HintText(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0391] clsButtonSpec.HintText ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0392] clsButtonSpec.Init ENTER"  ' [ADR-0100]
    m_btnName = btnName
    m_caption = caption
    m_cellAddr = cellAddr
    m_colorHex = colorHex
    m_groupName = groupName
    m_hintAddr = hintAddr
    m_hintText = hintText
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0393] clsButtonSpec.Init EXIT-OK"  ' [ADR-0100]
End Sub
```