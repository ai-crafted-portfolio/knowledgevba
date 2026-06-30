---
title: clsKnowledgeRegisterScreen.cls
description: clsKnowledgeRegisterScreen.cls のソースコード（コピペ用）
---

# clsKnowledgeRegisterScreen.cls

**配置先**: `検索.xlsm` 用の VBA モジュール
**種類**: クラスモジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\search\`
- ファイル名: `clsKnowledgeRegisterScreen.cls`
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
Attribute VB_Name = "clsKnowledgeRegisterScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsKnowledgeRegisterScreen (画面層 - M-05)
' 概要:   M-05 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1629] clsKnowledgeRegisterScreen.Init ENTER"  ' [ADR-0100]
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1630] clsKnowledgeRegisterScreen.Setup ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsKnowledgeRegisterScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "LOG-M05-SCREENCLS-SETUP-ENTRY")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsKnowledgeRegisterScreen", "Setup", "EXIT ok", "LOG-M05-SCREENCLS-SETUP-EXIT-OK")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1631] clsKnowledgeRegisterScreen.Setup EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1632] clsKnowledgeRegisterScreen.Setup EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Call modScreenRender.LogScreenError("clsKnowledgeRegisterScreen", "Setup", stepName, Err.Number, Err.Description, "LOG-M05-SCREENCLS-SETUP-ERR")
End Sub

Public Sub Render()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1633] clsKnowledgeRegisterScreen.Render ENTER"  ' [ADR-0100]
    Setup
End Sub
```
