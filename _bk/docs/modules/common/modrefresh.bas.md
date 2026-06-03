---
title: modRefresh.bas
description: modRefresh.bas のソースコード（コピペ用）
---

# modRefresh.bas

**配置先**: `共通モジュール (3 ブック全て)` 用の VBA モジュール  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\`
- ファイル名: `modRefresh.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

---

## ソースコード

```vb
Attribute VB_Name = "modRefresh"
' ============================================================
' modRefresh (Phase R-3-psi-Refresh, 2026-05-29)
'   ui_seed/<role>/M-NN.txt 繧� edit 縺励◆縺ゅ→縲∝�� install 縺帙★縺ｫ sheet 繧�
'   蜀肴緒逕ｻ縺吶ｋ縺溘ａ縺ｮ entry point縲Ｃutton OnClick / Alt+F8 macro 縺九ｉ蜻ｼ縺ｶ縲�
'   螳滉ｽ薙�ｯ clsSetupOrchestrator.Reapply* (RunFullSetup step6 ApplyUiStanzas
'   縺ｨ蜷御ｸ pipeline: BindSheet -> ClearScreen -> ApplyFromStanza -> ProtectSheet)縲�
'   譛ｬ繝｢繧ｸ繝･繝ｼ繝ｫ縺ｯ ScreenUpdating/EnableEvents 謚第ｭ｢ + ActiveSheet 蠕ｩ蜈�縺ｮ
'   阮�縺�繝ｩ繝�繝代�ｮ縺ｿ縲�ASCII-only (CP932/UTF-8 round-trip 螳牙�ｨ, ADR-0006)縲�
' ============================================================
Option Explicit

' 蜈ｨ sheet (LOG 莉･螟�) 繧� ui_seed 縺九ｉ蜀肴緒逕ｻ縲Ｃutton / Alt+F8 逕ｨ縲�
Public Sub Btn_RefreshAllSheets()
    Dim prevSU As Boolean, prevEv As Boolean
    Dim activeWs As Object
    On Error Resume Next
    Set activeWs = ThisWorkbook.ActiveSheet
    prevSU = Application.ScreenUpdating
    prevEv = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo 0

    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    orch.ReapplyAllSheets

    On Error Resume Next
    If Not activeWs Is Nothing Then activeWs.Activate
    Application.EnableEvents = prevEv
    Application.ScreenUpdating = prevSU
    Application.StatusBar = RefreshedMsg()
    On Error GoTo 0
End Sub

' ActiveSheet 1 譫壹□縺大�肴緒逕ｻ縲�Alt+F8 / button 逕ｨ縲�
Public Sub Btn_RefreshSheet_Active()
    Dim prevSU As Boolean, prevEv As Boolean
    Dim activeWs As Object
    On Error Resume Next
    Set activeWs = ThisWorkbook.ActiveSheet
    prevSU = Application.ScreenUpdating
    prevEv = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo 0

    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    orch.ReapplyActiveSheet

    On Error Resume Next
    If Not activeWs Is Nothing Then activeWs.Activate
    Application.EnableEvents = prevEv
    Application.ScreenUpdating = prevSU
    Application.StatusBar = RefreshedMsg()
    On Error GoTo 0
End Sub

' 謖�螳� screenId (M-NN) 1 譫壹ｒ蜀肴緒逕ｻ縲Ｑrogrammatic 逕ｨ (button OnAction 髱槫ｯｾ雎｡)縲�
Public Sub Btn_RefreshSheet(ByVal screenId As String)
    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    orch.ReapplySheet screenId
End Sub

' "譖ｴ譁ｰ螳御ｺ�" 繧� ChrW 縺ｧ讒狗ｯ� (譛ｬ .bas 繧� ASCII-only 縺ｫ菫昴▽)縲�
Private Function RefreshedMsg() As String
    RefreshedMsg = ChrW(&H66F4) & ChrW(&H65B0) & ChrW(&H5B8C) & ChrW(&H4E86)
End Function
```
