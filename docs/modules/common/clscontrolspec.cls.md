---
title: clsControlSpec.cls
description: clsControlSpec.cls のソースコード（コピペ用）
---

# clsControlSpec.cls

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: クラスモジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsControlSpec.cls`
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
Attribute VB_Name = "clsControlSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsControlSpec (ビジネスロジック層 / 値オブジェクト)
' 概要:   1 つの UserForm 上 Control の宣言情報を保持する ValueObject。
'         ControlType / Name / Left / Top / Width / Height / Caption / OnClick
'         を Private 変数 + Property Get/Let で公開する (R11 準拠)。
' 依存先: なし
' 備考:   状態を持つので .cls (vba-coding-standards R10/R11)。
'         OnClick はマクロ名 (Application.Run で呼ぶ)。
' ================================================================

Private m_controlType As String
Private m_name As String
Private m_left As Long
Private m_top As Long
Private m_width As Long
Private m_height As Long
Private m_caption As String
Private m_onClick As String

' --- ControlType ---
Public Property Get ControlType() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0416] clsControlSpec.ControlType ENTER"  ' [ADR-0100]
    ControlType = m_controlType
End Property
Public Property Let ControlType(ByVal v As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0417] clsControlSpec.ControlType ENTER"  ' [ADR-0100]
    m_controlType = v
End Property

' --- Name ---
Public Property Get Name() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0418] clsControlSpec.Name ENTER"  ' [ADR-0100]
    Name = m_name
End Property
Public Property Let Name(ByVal v As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0419] clsControlSpec.Name ENTER"  ' [ADR-0100]
    m_name = v
End Property

' --- Left ---
Public Property Get Left() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0420] clsControlSpec.Left ENTER"  ' [ADR-0100]
    Left = m_left
End Property
Public Property Let Left(ByVal v As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0421] clsControlSpec.Left ENTER"  ' [ADR-0100]
    m_left = v
End Property

' --- Top ---
Public Property Get Top() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0422] clsControlSpec.Top ENTER"  ' [ADR-0100]
    Top = m_top
End Property
Public Property Let Top(ByVal v As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0423] clsControlSpec.Top ENTER"  ' [ADR-0100]
    m_top = v
End Property

' --- Width ---
Public Property Get Width() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0424] clsControlSpec.Width ENTER"  ' [ADR-0100]
    Width = m_width
End Property
Public Property Let Width(ByVal v As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0425] clsControlSpec.Width ENTER"  ' [ADR-0100]
    m_width = v
End Property

' --- Height ---
Public Property Get Height() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0426] clsControlSpec.Height ENTER"  ' [ADR-0100]
    Height = m_height
End Property
Public Property Let Height(ByVal v As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0427] clsControlSpec.Height ENTER"  ' [ADR-0100]
    m_height = v
End Property

' --- Caption ---
Public Property Get Caption() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0428] clsControlSpec.Caption ENTER"  ' [ADR-0100]
    Caption = m_caption
End Property
Public Property Let Caption(ByVal v As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0429] clsControlSpec.Caption ENTER"  ' [ADR-0100]
    m_caption = v
End Property

' --- OnClick (マクロ名) ---
Public Property Get OnClick() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0430] clsControlSpec.OnClick ENTER"  ' [ADR-0100]
    OnClick = m_onClick
End Property
Public Property Let OnClick(ByVal v As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0431] clsControlSpec.OnClick ENTER"  ' [ADR-0100]
    m_onClick = v
End Property

' ================================================================
' 関数名: Init
' 概要:   全プロパティを 1 回でセットする初期化メソッド
' 引数:   cType   - "Label"/"TextBox"/"Button"/"Image"/"ListBox"/"ComboBox"
'         nm      - Control 名 (Form 内で一意)
'         l/t/w/h - Left/Top/Width/Height (points)
'         cap     - Caption (Label/Button/CheckBox 用、省略可)
'         onClk   - OnClick マクロ名 (Button 用、省略可)
' ================================================================
Public Sub Init(ByVal cType As String, ByVal nm As String, _
                 ByVal l As Long, ByVal t As Long, _
                 ByVal w As Long, ByVal h As Long, _
                 Optional ByVal cap As String = "", _
                 Optional ByVal onClk As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0432] clsControlSpec.Init ENTER"  ' [ADR-0100]
    m_controlType = cType
    m_name = nm
    m_left = l
    m_top = t
    m_width = w
    m_height = h
    m_caption = cap
    m_onClick = onClk
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0433] clsControlSpec.Init EXIT-OK"  ' [ADR-0100]
End Sub
```
