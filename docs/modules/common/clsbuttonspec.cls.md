---
title: clsButtonSpec.cls
description: clsButtonSpec.cls のソースコード（コピペ用）
---

# clsButtonSpec.cls

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: クラスモジュール
**更新日**: 2026-06-03 23:22 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsButtonSpec.cls`
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
