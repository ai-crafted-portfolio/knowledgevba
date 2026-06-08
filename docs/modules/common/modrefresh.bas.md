---
title: modRefresh.bas
description: modRefresh.bas 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# modRefresh.bas

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`modRefresh.bas`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1318] modRefresh.Btn_RefreshAllSheets ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_RefreshAllSheets"
    Dim XLSM As String
    XLSM = Replace$(ThisWorkbook.Name, ".xlsm", "")
    modBtnGuard.LogEnter BTN, XLSM
    modCommon.AppendProgressLog "[BTN-REFRESH-ENTER] " & Now
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1319] modRefresh.Btn_RefreshAllSheets EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    ' [P5d fix 2026-06-04] ScreenUpdating=False 削除 (shape 描画 silent fail 回避)
    Dim prevEv As Boolean
    Dim activeWs As Object
    On Error Resume Next
    Set activeWs = ThisWorkbook.ActiveSheet
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo 0

    modCommon.AppendProgressLog "[BTN-REFRESH-PASS-CHECKPREREQ]"
    ' [P5e fix 2026-06-04] install と同じ RunFullSetup を invoke (ReapplyAllSheets による画面破壊回避)
    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    Dim baseName As String
    baseName = Replace(ThisWorkbook.Name, ".xlsm", "")
    modCommon.AppendProgressLog "[BTN-REFRESH-DISPATCH] baseName=" & baseName
    Select Case baseName
        Case ChrW(&H7BA1) & ChrW(&H7406): orch.RunForAdmin
        Case ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63): orch.RunForRegister
        Case ChrW(&H691C) & ChrW(&H7D22): orch.RunForSearch
        Case Else: orch.RunForAdmin
    End Select

    On Error Resume Next
    If Not activeWs Is Nothing Then activeWs.Activate
    Application.EnableEvents = prevEv
    Application.StatusBar = RefreshedMsg()
    On Error GoTo 0
    ' 2026-06-07: Re-seed M-03 C4 + reload M-10 storage paths after refresh
    If XLSM = ChrW(&H7BA1) & ChrW(&H7406) Then
        On Error Resume Next
        modEntryFormat.SeedM03FormatIdIfEmpty
        Dim wsStor As Worksheet
        Set wsStor = ThisWorkbook.Worksheets(ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & ChrW(&H8A2D) & ChrW(&H5B9A))
        modEntrySettings.LoadStorageToSheet wsStor
        On Error GoTo 0
    End If
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1320] modRefresh.Btn_RefreshAllSheets EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1321] modRefresh.Btn_RefreshAllSheets EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    ' [P2 fix] ScreenUpdating/EnableEvents restore before HandleButtonError
    On Error Resume Next
    If Not activeWs Is Nothing Then activeWs.Activate
    Application.EnableEvents = prevEv
    On Error GoTo 0
    modBtnGuard.HandleButtonError BTN, Err, XLSM
    modBtnGuard.LogExit BTN, XLSM, False
End Sub

' ActiveSheet 1 枚だけ再描画。Alt+F8 / button 用。
Public Sub Btn_RefreshSheet_Active()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1322] modRefresh.Btn_RefreshSheet_Active ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_RefreshSheet_Active"
    Dim XLSM As String
    XLSM = Replace$(ThisWorkbook.Name, ".xlsm", "")
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1323] modRefresh.Btn_RefreshSheet_Active EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim prevSU As Boolean, prevEv As Boolean
    Dim activeWs As Object
    On Error Resume Next
    Set activeWs = ThisWorkbook.ActiveSheet
    prevSU = Application.ScreenUpdating
    prevEv = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo 0

    ' [P5e horizontal expand 2026-06-04] install と同じ RunFullSetup 経路 (画面破壊回避)
    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    Dim baseName As String
    baseName = Replace(ThisWorkbook.Name, ".xlsm", "")
    modCommon.AppendProgressLog "[BTN-REFRESH-ACTIVE-DISPATCH] baseName=" & baseName
    Select Case baseName
        Case ChrW(&H7BA1) & ChrW(&H7406): orch.RunForAdmin
        Case ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63): orch.RunForRegister
        Case ChrW(&H691C) & ChrW(&H7D22): orch.RunForSearch
        Case Else: orch.RunForAdmin
    End Select

    On Error Resume Next
    If Not activeWs Is Nothing Then activeWs.Activate
    Application.EnableEvents = prevEv
    Application.ScreenUpdating = prevSU
    Application.StatusBar = RefreshedMsg()
    On Error GoTo 0
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1324] modRefresh.Btn_RefreshSheet_Active EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1325] modRefresh.Btn_RefreshSheet_Active EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.HandleButtonError BTN, Err, XLSM
    modBtnGuard.LogExit BTN, XLSM, False
End Sub

' 指定 screenId (M-NN) 1 枚を再描画。programmatic 用 (button OnAction 非対象)。
Public Sub Btn_RefreshSheet(ByVal screenId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1326] modRefresh.Btn_RefreshSheet ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_RefreshSheet"
    Dim XLSM As String
    XLSM = Replace$(ThisWorkbook.Name, ".xlsm", "")
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1327] modRefresh.Btn_RefreshSheet EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    orch.ReapplySheet screenId
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1328] modRefresh.Btn_RefreshSheet EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1329] modRefresh.Btn_RefreshSheet EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.HandleButtonError BTN, Err, XLSM
    modBtnGuard.LogExit BTN, XLSM, False
End Sub

' "更新完了" を ChrW で構築 (本 .bas を ASCII-only に保つ)。
Private Function RefreshedMsg() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1330] modRefresh.RefreshedMsg ENTER"  ' [ADR-0100]
    RefreshedMsg = ChrW(&H66F4) & ChrW(&H65B0) & ChrW(&H5B8C) & ChrW(&H4E86)
End Function
```