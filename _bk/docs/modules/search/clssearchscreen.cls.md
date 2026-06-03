---
title: clsSearchScreen.cls
description: clsSearchScreen.cls のソースコード（コピペ用）
---

# clsSearchScreen.cls

**配置先**: `検索.xlsm` 用の VBA モジュール  
**種類**: クラス モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\search\`
- ファイル名: `clsSearchScreen.cls`
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
Attribute VB_Name = "clsSearchScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' 繧ｯ繝ｩ繧ｹ: clsSearchScreen (逕ｻ髱｢螻､ - M-08)
' 讎りｦ�:   M-08 逕ｻ髱｢縺ｮ讒狗ｯ峨�ｻ蜀肴緒逕ｻ縲Ｔpec 繧� modScreenRender 縺ｫ蟋碑ｭｲ縲�
' 萓晏ｭ伜��: IScreenRenderer, clsScreenSpec, modScreenRender
' 蛯呵�:   E2E rerun (2026-05-12) 縺ｧ truncated 迥ｶ諷九ｒ蜈ｨ繧ｯ繝ｩ繧ｹ蜷御ｸ繝�繝ｳ繝励Ξ縺ｧ蠕ｩ譌ｧ縲�
'         ENTER/EXIT 繝医Ξ繝ｼ繧ｹ + ErrHandler 縺ｧ縲後←縺薙〒螟ｱ謨励阪′蛻�縺九ｋ讒矩縺ｫ邨ｱ荳縲�
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsSearchScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "LOG-M08-SCREENCLS-SETUP-ENTRY")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsSearchScreen", "Setup", "EXIT ok", "LOG-M08-SCREENCLS-SETUP-EXIT-OK")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsSearchScreen", "Setup", stepName, Err.Number, Err.Description, "LOG-M08-SCREENCLS-SETUP-ERR")
End Sub

Public Sub Render()
    Setup
End Sub
```
