---
title: modRefresh.bas
description: modRefresh.bas のソースコード（コピペ用）
---

# modRefresh.bas

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modRefresh.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modRefresh"
' ============================================================
' modRefresh (Phase R-3-psi-Refresh, 2026-05-29)
'   ui_seed/<role>/M-NN.txt を edit したあと、再 install せずに sheet を
'   再描画するための entry point。button OnClick / Alt+F8 macro から呼ぶ。
'   実体は clsSetupOrchestrator.Reapply* (RunFullSetup step6 ApplyUiStanzas
'   と同一 pipeline: BindSheet -> ClearScreen -> ApplyFromStanza -> ProtectSheet)。
'   本モジュールは ScreenUpdating/EnableEvents 抑止 + ActiveSheet 復元の
'   薄いラッパのみ。ASCII-only (CP932/UTF-8 round-trip 安全, ADR-0006)。
' ============================================================
Option Explicit

' 全 sheet (LOG 以外) を ui_seed から再描画。button / Alt+F8 用。
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

' ActiveSheet 1 枚だけ再描画。Alt+F8 / button 用。
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

' 指定 screenId (M-NN) 1 枚を再描画。programmatic 用 (button OnAction 非対象)。
Public Sub Btn_RefreshSheet(ByVal screenId As String)
    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    orch.ReapplySheet screenId
End Sub

' "更新完了" を ChrW で構築 (本 .bas を ASCII-only に保つ)。
Private Function RefreshedMsg() As String
    RefreshedMsg = ChrW(&H66F4) & ChrW(&H65B0) & ChrW(&H5B8C) & ChrW(&H4E86)
End Function
```
