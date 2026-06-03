---
title: clsControlSpec.cls
description: clsControlSpec.cls のソースコード（コピペ用）
---

# clsControlSpec.cls

**配置先**: 共通モジュール（3 ブック共通）
**種類**: クラスモジュール

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
    ControlType = m_controlType
End Property
Public Property Let ControlType(ByVal v As String)
    m_controlType = v
End Property

' --- Name ---
Public Property Get Name() As String
    Name = m_name
End Property
Public Property Let Name(ByVal v As String)
    m_name = v
End Property

' --- Left ---
Public Property Get Left() As Long
    Left = m_left
End Property
Public Property Let Left(ByVal v As Long)
    m_left = v
End Property

' --- Top ---
Public Property Get Top() As Long
    Top = m_top
End Property
Public Property Let Top(ByVal v As Long)
    m_top = v
End Property

' --- Width ---
Public Property Get Width() As Long
    Width = m_width
End Property
Public Property Let Width(ByVal v As Long)
    m_width = v
End Property

' --- Height ---
Public Property Get Height() As Long
    Height = m_height
End Property
Public Property Let Height(ByVal v As Long)
    m_height = v
End Property

' --- Caption ---
Public Property Get Caption() As String
    Caption = m_caption
End Property
Public Property Let Caption(ByVal v As String)
    m_caption = v
End Property

' --- OnClick (マクロ名) ---
Public Property Get OnClick() As String
    OnClick = m_onClick
End Property
Public Property Let OnClick(ByVal v As String)
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
    m_controlType = cType
    m_name = nm
    m_left = l
    m_top = t
    m_width = w
    m_height = h
    m_caption = cap
    m_onClick = onClk
End Sub
```
