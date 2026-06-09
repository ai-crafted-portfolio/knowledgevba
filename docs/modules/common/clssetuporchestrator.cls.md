---
title: clsSetupOrchestrator.cls
description: clsSetupOrchestrator.cls のソースコード（コピペ用）
---

# clsSetupOrchestrator.cls

**配置先**: 共通モジュール（3 ブック共通）
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsSetupOrchestrator.cls`
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
End
Attribute VB_Name = "clsSetupOrchestrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' ?N???X: clsSetupOrchestrator?iv2.1?APhase 5-7?A3 xlsm ?? setup?j
' ?T?v:   xlsm ?N?????? setup orchestration?iconfig ??? ?? ???O?????? ?? ?V?[?g?\?z ?? ??? ?? ActiveSheet?j
' Version: v2.1?i2026-05-16 EOD?AQ1-Q57 ?????f?j
' ??A:   ADR-0053 ??2.1 / ??2.5 / ??2.8 #4, Q44 / Q39-Q44 / Q34
' SSOT:   ?{?N???X??{?? runtime setup ?? SSOT?iADR-0053 ??2 / SSOT-Q17?j?B
'         3 xlsm ??V?[?g?\?????E?N???V?[?g???? ADR-0053 ??2.1 ?\?????? SSOT?B
' v2.1 ??v?X?V:
'   - ???\???iADR-0053 ??2.1.1 ?T?u?O???[?v???j: ?o?^?C??=M-05/M-06?A????=M-08?A???=M-02/M-03/M-04/M-12/M-13/M-10/M-11/M-14?iLOG ????j
'   - ?N???? ActiveSheet?iADR-0053 ??2.1 / ??9 M-2?j: ?o?^?C??=M-05?A????=M-08?A???=M-02
'   - Q39: ???.xlsm ?? SheetChange ?n???h???s?v?iM-11 ??? ID ??n??wM-11 ???x?L?q??????A?V ID ??n??? M-11=????????Barchitecture ??10-3 L37 ???L??????j
'   - Q34: ???.xlsm ?N?????? CleanupOldBackups ????o
'   - ?H???L?q?p?~?iQ30?j
' ================================================================
Option Explicit



Private m_logger As clsLogger

' xlsm ?? sheet ?\???iADR-0053 ??2.1 / modSetup.bas SHEETS_* ????l?AR6-01?j
' SSOT: ADR-0053 ??2.1 ?\?i??2.1.1 ?T?u?O???[?v???j?B?o?^?C??=M-05/M-06?A????=M-08?A???=M-02/M-03/M-04/M-12/M-13/M-10/M-11/M-14?iLOG ??S xlsm ????j
Private Const SHEETS_TOUROKU As String = "M-05|M-06|LOG"
' v2.3 fix (2026-05-27): M-07 ?i???b?W?? ?p?~ (ADR-0072 ??2.1?AM-08 ?????????)
' v2.3 fix (2026-05-31): M-09 ?i???b?W?\?? ?p?~ (M-08 grid DoubleClick ?\???A?I??OpenViewWithId ?????????o?H?)
Private Const SHEETS_KENSAKU As String = "M-08|M-05|LOG"
' v2.3 fix (2026-05-27): M-13 ?t?@?C???`????? ?p?~ (ADR-0072 ??2.1)
Private Const SHEETS_KANRI As String = "M-02|M-03|M-12|M-10|M-11|LOG"   ' 2026-06-07: M-14 retired

' ?^?u?F?iADR-0053 ??2.1.1?j
Private Const TAB_COLOR_TOUROKU As Long = 16711892  ' #FFB6C1 pink (BGR: B6C1FF -> FFC1B6 reversed)
Private Const TAB_COLOR_KENSAKU As Long = 15128749  ' #ADD8E6 lightblue
Private Const TAB_COLOR_KANRI As Long = 9498256     ' #90EE90 lightgreen

' ?N???? ActiveSheet?iADR-0053 ??2.1 ?\ / ??9 M-2?AR6-01?j
' SSOT: ?o?^?C??=M-05?A????=M-08?A???=M-02

' ----------------------------------------------------------------
' Public API: spec ??6.1 convenience wrappers
' ----------------------------------------------------------------

' v2.3 (2026-05-27 fix): JP xlsm names constructed via ChrW() so the .cls
' source stays ASCII-only and survives CP932/UTF-8 round trips.
Private Function NAME_KANRI() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0687] clsSetupOrchestrator.NAME_KANRI ENTER"  ' [ADR-0100]
    NAME_KANRI = ChrW(&H7BA1) & ChrW(&H7406)
End Function
Private Function NAME_TOUROKU() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0688] clsSetupOrchestrator.NAME_TOUROKU ENTER"  ' [ADR-0100]
    NAME_TOUROKU = ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63)
End Function
Private Function NAME_KENSAKU() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0689] clsSetupOrchestrator.NAME_KENSAKU ENTER"  ' [ADR-0100]
    NAME_KENSAKU = ChrW(&H691C) & ChrW(&H7D22)
End Function

Public Sub RunForRegister()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0690] clsSetupOrchestrator.RunForRegister ENTER"  ' [ADR-0100]
    RunFullSetup NAME_TOUROKU
End Sub

Public Sub RunForSearch()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0691] clsSetupOrchestrator.RunForSearch ENTER"  ' [ADR-0100]
    RunFullSetup NAME_KENSAKU
End Sub

Public Sub RunForAdmin()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0692] clsSetupOrchestrator.RunForAdmin ENTER"  ' [ADR-0100]
    RunFullSetup NAME_KANRI
End Sub

' ----------------------------------------------------------------
' Public API: Refresh Screen (Phase R-3-psi-Refresh, 2026-05-29)
'   ui_seed/<role>/<screen-id>.txt を編??したあと、??? install せずに sheet ??
'   再描画する。RunFullSetup step 6 (ApplyUiStanzas) と同じ pipeline ??
'   そ???まま再利用する (BindSheet -> ClearScreen -> ApplyFromStanza ->
'   ProtectSheet)。role は ThisWorkbook 名から判定する???で 3 book 共通??
' ----------------------------------------------------------------

' 現 role の全 sheet (LOG 以??) ?? ui_seed から再適用する??
Public Sub ReapplyAllSheets()
    modCommon.AppendProgressLog "[RAS-ENTER] " & Now
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0693] clsSetupOrchestrator.ReapplyAllSheets ENTER"  ' [ADR-0100]
    On Error GoTo EH
    Dim xlsmName As String
    xlsmName = CurrentXlsmName()
    EnsureConfigLoaded xlsmName

    Dim sheetList As String
    sheetList = GetSheetsCsv(xlsmName)
    modCommon.AppendProgressLog "[RAS-CSV] xlsm=" & xlsmName & " list=" & sheetList
    If Len(sheetList) = 0 Then Exit Sub

    Dim names() As String
    names = Split(sheetList, "|")

    Dim renderer As IScreenRenderer
    Set renderer = New clsSheetRenderer

    ' [P5c fix 2026-06-04] ApplyTabColors 削除 (install 時のみ設定、refresh 時の Protected sheet では err=1004、ADR-0093)
    modCommon.AppendProgressLog "[RAS-TABCOLOR-SKIPPED] reason=already-set-at-install xlsm=" & xlsmName

    Dim i As Long
    For i = LBound(names) To UBound(names)
        If names(i) <> "LOG" Then
            On Error Resume Next
            modCommon.AppendProgressLog "[RAS-BIND] sheet=" & names(i) & " xlsm=" & xlsmName
            renderer.BindSheet names(i)
            If Err.Number <> 0 Then modCommon.AppendProgressLog "[RAS-BIND-ERR] " & names(i) & " err=" & Err.Number & " desc=" & Err.Description
            Err.Clear
            modCommon.AppendProgressLog "[RAS-CLEAR] sheet=" & names(i)
            renderer.ClearScreen
            If Err.Number <> 0 Then modCommon.AppendProgressLog "[RAS-CLEAR-ERR] " & names(i) & " err=" & Err.Number & " desc=" & Err.Description
            Err.Clear
            modCommon.AppendProgressLog "[RAS-APPLY] sheet=" & names(i) & " xlsm=" & xlsmName
            renderer.ApplyFromStanza xlsmName, names(i)
            If Err.Number <> 0 Then modCommon.AppendProgressLog "[RAS-APPLY-ERR] " & names(i) & " err=" & Err.Number & " desc=" & Err.Description
            Err.Clear
            modCommon.AppendProgressLog "[RAS-PROT] sheet=" & names(i)
            renderer.ProtectSheet "light"
            If Err.Number <> 0 Then modCommon.AppendProgressLog "[RAS-PROT-ERR] " & names(i) & " err=" & Err.Number & " desc=" & Err.Description
            Err.Clear
            On Error GoTo EH
        End If
    Next i

    ' BUG-NEW1 fix (2026-05-31): ClearScreen wipes ALL shapes including the v23
    ' form-control buttons (新規登録 / 修正フォームを開?? / 表示フォームを開??)
    ' that are placed by modSheetButtons.PlaceV23SheetButtons during install.
    ' Re-place them here so ReapplyAllSheets keeps the UI workable.
    On Error Resume Next
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0695] clsSetupOrchestrator.ReapplyAllSheets STEP-1 pre modSheetButtons.PlaceV23SheetButtons"  ' [ADR-0100]
    Call modSheetButtons.PlaceV23SheetButtons
    On Error GoTo EH
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0694] clsSetupOrchestrator.ReapplyAllSheets EXIT-OK"  ' [ADR-0100]
    Exit Sub
EH:
    Debug.Print "[clsSetupOrchestrator.ReapplyAllSheets ERROR] " & Err.Description
End Sub

' ???? screenId (M-prefixed screen identifier) の sheet 1 枚だけ???適用する??
Public Sub ReapplySheet(ByVal screenId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0696] clsSetupOrchestrator.ReapplySheet ENTER"  ' [ADR-0100]
    On Error GoTo EH
    Dim xlsmName As String
    xlsmName = CurrentXlsmName()
    EnsureConfigLoaded xlsmName

    Dim renderer As IScreenRenderer
    Set renderer = New clsSheetRenderer
    On Error Resume Next
    renderer.BindSheet screenId
    renderer.ClearScreen
    renderer.ApplyFromStanza xlsmName, screenId
    renderer.ProtectSheet "light"
    On Error GoTo EH

    ' BUG-NEW1 fix (2026-05-31): same rationale as ReapplyAllSheets - the v23
    ' form-control buttons live outside the ui_seed stanza and must be replaced
    ' after ClearScreen. PlaceV23SheetButtons is idempotent (it clears its own
    ' "v23_btn_*" shapes first), so calling it for any single-sheet refresh is
    ' safe and cheap.
    On Error Resume Next
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0698] clsSetupOrchestrator.ReapplySheet STEP-1 pre modSheetButtons.PlaceV23SheetButtons"  ' [ADR-0100]
    Call modSheetButtons.PlaceV23SheetButtons
    On Error GoTo EH
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0697] clsSetupOrchestrator.ReapplySheet EXIT-OK"  ' [ADR-0100]
    Exit Sub
EH:
    Debug.Print "[clsSetupOrchestrator.ReapplySheet ERROR] " & Err.Description
End Sub

' ActiveSheet (JP 表示??) ?? screenId に??引きして 1 枚???適用する??
Public Sub ReapplyActiveSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0699] clsSetupOrchestrator.ReapplyActiveSheet ENTER"  ' [ADR-0100]
    On Error GoTo EH
    Dim xlsmName As String
    xlsmName = CurrentXlsmName()
    EnsureConfigLoaded xlsmName

    Dim sid As String
    sid = ScreenIdForSheetName(ThisWorkbook.ActiveSheet.Name, xlsmName)
    If Len(sid) > 0 Then ReapplySheet sid
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0700] clsSetupOrchestrator.ReapplyActiveSheet EXIT-OK"  ' [ADR-0100]
    Exit Sub
EH:
    Debug.Print "[clsSetupOrchestrator.ReapplyActiveSheet ERROR] " & Err.Description
End Sub

' ----------------------------------------------------------------
' Public API
' ----------------------------------------------------------------

' ?T?v: xlsm ?N??????S setup ?????s
' ????: xlsmName = "?o?^?C??" / "????" / "???"
Public Sub RunFullSetup(ByVal xlsmName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0701] clsSetupOrchestrator.RunFullSetup ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    ' v2.3 install hang fix (2026-05-26): ?e step ??i?????O??????
    ' install_admin.bat ?o?H???????i?????? c:\temp\setup_admin_progress.txt
    ' ???????????????? (Debug.Print ?? VBE Immediate ??????o???????A
    ' modCommon.AppendProgressLog ??t?@?C?????p)?B
    ' clsLogger ???????O (step 1) ??t?@?C???o????A???????? (step 2 ??~) ??
    ' LOG ?V?[?g??? m_logger.LogInfo ??????B
    Dim t As String

    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 0 : enter xlsmName=" & xlsmName
    modCommon.AppendProgressLog t & "RunFullSetup step 0 : enter xlsmName=" & xlsmName

    ' v2.3 fix (2026-05-31): the workbook may still be Protected from a previous
    ' install. EnsureSheets / DropRetiredScreens cannot delete sheets while
    ' Workbook.Protect Structure=True, so unprotect at the top of every setup
    ' run. ApplyProtection at step 7 re-protects after all sheet ops finish.
    On Error Resume Next
    ThisWorkbook.Unprotect
    Dim wsi As Long
    For wsi = 1 To ThisWorkbook.Worksheets.Count
        ThisWorkbook.Worksheets(wsi).Unprotect
    Next wsi
    Err.Clear
    On Error GoTo ErrHandler

    ' (1) modConfigLoader ?? xlsm ?? config.txt ?? read ?? modConfigHolder ??Z?b?g?iQ8?j
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 1 pre : LoadConfig(" & xlsmName & ")"
    modCommon.AppendProgressLog t & "RunFullSetup step 1 pre : LoadConfig(" & xlsmName & ")"
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0704] clsSetupOrchestrator.RunFullSetup STEP-1 pre modConfigLoader.LoadConfig"  ' [ADR-0100]
    Call modConfigLoader.LoadConfig(xlsmName)
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 1 post : LoadConfig done"
    modCommon.AppendProgressLog t & "RunFullSetup step 1 post : LoadConfig done"

    ' (2) clsLogger ???????iQ7 ???? ERROR?j
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 2 pre : EnsureLogSheet + clsLogger.Init"
    modCommon.AppendProgressLog t & "RunFullSetup step 2 pre : EnsureLogSheet + clsLogger.Init"
    EnsureLogSheet xlsmName
    Set m_logger = New clsLogger
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0705] clsSetupOrchestrator.RunFullSetup STEP-2 pre m_logger.Init"  ' [ADR-0100]
    Call m_logger.Init(ThisWorkbook.Worksheets("LOG"))
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "?N??: " & xlsmName, "", "LOG-SETUP-START"
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 2 post : clsLogger ready"
    modCommon.AppendProgressLog t & "RunFullSetup step 2 post : clsLogger ready"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 2 post : clsLogger ready", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler

    ' (3) ???.xlsm ???: ?N?????? 90 ?????o?b?N?A?b?v???????iQ34?j
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 3 pre : CleanupOldBackups (kanri-only)"
    modCommon.AppendProgressLog t & "RunFullSetup step 3 pre : CleanupOldBackups (kanri-only)"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 3 pre : CleanupOldBackups", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler
    If xlsmName = NAME_KANRI Then
        Dim deleted As Long
        deleted = modKnowledgeFileIO.CleanupOldBackups()
        If deleted > 0 Then
            m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", _
                "90 ?????o?b?N?A?b?v??: " & deleted & " ??", "", "LOG-BCK-CLEAN"
        End If
    End If
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 3 post : CleanupOldBackups done"
    modCommon.AppendProgressLog t & "RunFullSetup step 3 post : CleanupOldBackups done"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 3 post : CleanupOldBackups done", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler

    ' (4) ?V?[?g?m??iv2.1 ?\???A???p?~?? M-01 ???j
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 4 pre : EnsureSheets"
    modCommon.AppendProgressLog t & "RunFullSetup step 4 pre : EnsureSheets"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 4 pre : EnsureSheets", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler
    EnsureSheets xlsmName
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 4 post : EnsureSheets done"
    modCommon.AppendProgressLog t & "RunFullSetup step 4 post : EnsureSheets done"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 4 post : EnsureSheets done", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler

    ' (5) ?^?u?F???iADR-0053 ??2.1.1?j
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 5 pre : ApplyTabColors"
    modCommon.AppendProgressLog t & "RunFullSetup step 5 pre : ApplyTabColors"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 5 pre : ApplyTabColors", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler
    ApplyTabColors xlsmName
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 5 post : ApplyTabColors done"
    modCommon.AppendProgressLog t & "RunFullSetup step 5 post : ApplyTabColors done"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 5 post : ApplyTabColors done", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler

    ' (6) UI ?X?^???U?K?p?iQ20 modUILoader.ApplyUiToSheet?j
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 6 pre : ApplyUiStanzas"
    modCommon.AppendProgressLog t & "RunFullSetup step 6 pre : ApplyUiStanzas"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 6 pre : ApplyUiStanzas", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler
    ApplyUiStanzas xlsmName
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 6 post : ApplyUiStanzas done"
    modCommon.AppendProgressLog t & "RunFullSetup step 6 post : ApplyUiStanzas done"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 6 post : ApplyUiStanzas done", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler

    ' (7) Workbook.Protect Structure + ?V?[?g???
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 7 pre : ApplyProtection"
    modCommon.AppendProgressLog t & "RunFullSetup step 7 pre : ApplyProtection"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 7 pre : ApplyProtection", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler
    ApplyProtection
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 7 post : ApplyProtection done"
    modCommon.AppendProgressLog t & "RunFullSetup step 7 post : ApplyProtection done"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 7 post : ApplyProtection done", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler

    ' (8) ActiveSheet = ?N???????iQ44?j
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 8 pre : ActivateStartupSheet"
    modCommon.AppendProgressLog t & "RunFullSetup step 8 pre : ActivateStartupSheet"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 8 pre : ActivateStartupSheet", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler
    ActivateStartupSheet xlsmName
    ' R-2-Fix2b (2026-05-28): a content sheet is now active; hide LOG VeryHidden
    ' (cannot hide the active sheet, so this runs after ActivateStartupSheet).
    On Error Resume Next
    ThisWorkbook.Worksheets("LOG").Visible = xlSheetHidden   ' 2026-06-07: was VeryHidden -- switched so user can right-click Unhide
    On Error GoTo ErrHandler
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 8 post : ActivateStartupSheet done"
    modCommon.AppendProgressLog t & "RunFullSetup step 8 post : ActivateStartupSheet done"
    On Error Resume Next
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 8 post : ActivateStartupSheet done", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler

    ' Phase Q-2 (2026-05-27): place v2.3 form-control buttons on M-05/06/08/09
    ' so the user can drive the UserForm flow without Alt+F8.
    On Error Resume Next
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0706] clsSetupOrchestrator.RunFullSetup STEP-3 pre modSheetButtons.PlaceV23SheetButtons"  ' [ADR-0100]
    Call modSheetButtons.PlaceV23SheetButtons
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "[progress] step 8.5 : PlaceV23SheetButtons done", xlsmName, "LOG-SETUP-PROG"
    On Error GoTo ErrHandler

    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "????: " & xlsmName, "", "LOG-SETUP-OK"
    t = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print t & "RunFullSetup step 9 : COMPLETE xlsmName=" & xlsmName
    modCommon.AppendProgressLog t & "RunFullSetup step 9 : COMPLETE xlsmName=" & xlsmName
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0702] clsSetupOrchestrator.RunFullSetup EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0703] clsSetupOrchestrator.RunFullSetup EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Dim eT As String
    eT = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print eT & "RunFullSetup ErrHandler : xlsmName=" & xlsmName & " errNum=" & Err.Number & " errDesc=" & Err.Description
    modCommon.AppendProgressLog eT & "RunFullSetup ErrHandler : xlsmName=" & xlsmName & " errNum=" & Err.Number & " errDesc=" & Err.Description
    ' Q7 ?K?? X: error handler ????K?? LogError
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsSetupOrchestrator", "RunFullSetup", Err.Description, "", "LOG-SETUP-ERR"
    Else
        Debug.Print "[clsSetupOrchestrator.RunFullSetup ERROR] " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------
' Private: setup ?X?e?b?v
' ----------------------------------------------------------------

Private Sub EnsureLogSheet(ByVal xlsmName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0707] clsSetupOrchestrator.EnsureLogSheet ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("LOG")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "LOG"
    End If
    ' Q7 ?K?? Y
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0708] clsSetupOrchestrator.EnsureLogSheet EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub EnsureSheets(ByVal xlsmName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0709] clsSetupOrchestrator.EnsureSheets ENTER"  ' [ADR-0100]
    Dim sheetList As String
    sheetList = GetSheetsCsv(xlsmName)
    If Len(sheetList) = 0 Then Exit Sub

    Dim names() As String
    names = Split(sheetList, "|")

    Dim i As Long
    For i = LBound(names) To UBound(names)
        ' v2.3 fix (2026-05-30): pass xlsmName so EnsureSheetExists can resolve
        ' ui_seed [SHEET].SheetName (display name) and avoid recreating literal
        ' "M-NN" sheets that the sheet audit / rename step already migrated to
        ' the JP display name (e.g. "格納???設??"). See split-state bug fix.
        EnsureSheetExists xlsmName, names(i)
    Next i

    ' v2.3 fix (2026-05-27): delete default "Sheet1" / "Sheet2" / "Sheet3" left
    ' over from a fresh empty xlsm if they are not in our target set and the
    ' workbook now has at least one of our managed sheets present.
    DeleteDefaultLeftoverSheets sheetList

    ' v2.3 fix (2026-05-30): build "managed set" for both literal screenIds AND
    ' their display names, then drop any stray literal "M-NN" sibling that now
    ' has a renamed display-name twin in the same workbook. This finishes the
    ' split-state cleanup so the next Workbook_Open cannot re-trigger it.
    DropOrphanLiteralScreenIdSheets xlsmName, sheetList

    ' SPEC_DRIFT-NEW1/NEW2 fix (2026-05-31): drop sheets for screens that this
    ' role explicitly retired (M-07 ナレ??ジ一覧 in search, M-13 ファイル形??
    ' 設?? in admin). Older installs left these sheets behind because they
    ' were never in the SHEETS_* SSOT, but also never explicitly listed as
    ' "delete-me". Now they are.
    DropRetiredScreens xlsmName
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0710] clsSetupOrchestrator.EnsureSheets EXIT-OK"  ' [ADR-0100]
End Sub

' SPEC_DRIFT-NEW1/NEW2 (2026-05-31): per-role list of screenIds that the v2.3
' design explicitly retired. EnsureSheets removes any sheet whose Name (or
' display-name resolved via ui_seed [SHEET].SheetName backup) matches one of
' these. Safe: we never touch a screen that is in the current SHEETS_* SSOT.
Private Sub DropRetiredScreens(ByVal xlsmName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0711] clsSetupOrchestrator.DropRetiredScreens ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim retired As String
    Select Case xlsmName
        Case NAME_KENSAKU
            ' M-07 = ナレ??ジ一覧 (search-role, retired by ADR-0072 §2.1)
            ' M-09 = ナレ??ジ表示 (search-role, retired 2026-05-31; view UserForm
            '        is launched only via M-08 grid DoubleClick -> OpenViewWithId)
            retired = "M-07|M-09"
        Case NAME_KANRI
            ' M-13 = ファイル形式設?? (admin-role, retired by ADR-0072 §2.1)
            retired = "M-13|M-04|M-14"   ' 2026-06-07: M-14 ｴ鋓顥ｸ retired (user dropped operation-log screen)
        Case Else
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0712] clsSetupOrchestrator.DropRetiredScreens EXIT-OK"  ' [ADR-0100]
            Exit Sub
    End Select

    Dim ids() As String
    ids = Split(retired, "|")

    Application.DisplayAlerts = False
    ' v2.3 fix (2026-05-31): build complete name set (literal screenId +
    ' display name via ui_seed lookup, with hard-coded fallback) then
    ' iterate the Worksheets collection and delete every sheet whose Name
    ' matches. This is more robust than Worksheets(name) lookup because it
    ' tolerates any subtle name-encoding mismatch the placeholder <screen-id>.txt
    ' round-trip might introduce.
    Dim dropSet As Object
    Set dropSet = CreateObject("Scripting.Dictionary")
    dropSet.CompareMode = 1
    Dim k As Long
    For k = LBound(ids) To UBound(ids)
        Dim sid As String
        sid = ids(k)
        If Not dropSet.Exists(sid) Then dropSet.Add sid, True
        Dim disp As String
        disp = ReadDisplayNameFromUiSeed(xlsmName, sid)
        If Len(disp) = 0 Then disp = RetiredDisplayNameFallback(sid)
        If Len(disp) > 0 And disp <> sid Then
            If Not dropSet.Exists(disp) Then dropSet.Add disp, True
        End If
    Next k

    ' Iterate Worksheets in reverse so Delete during enumeration is safe.
    Dim w As Long
    For w = ThisWorkbook.Worksheets.Count To 1 Step -1
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(w)
        If dropSet.Exists(ws.Name) Then
            ws.Delete
        End If
    Next w
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

' v2.3 fix (2026-05-31): hard-coded display-name fallback for retired screens
' whose ui_seed <screen-id>.txt has been replaced by a placeholder. JP literals use
' ChrW() so this .cls stays CP932-clean.
Private Function RetiredDisplayNameFallback(ByVal screenId As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0713] clsSetupOrchestrator.RetiredDisplayNameFallback ENTER"  ' [ADR-0100]
    Select Case screenId
        Case "M-07"
            ' ナレ??ジ一覧 = U+30CA U+30EC U+30C3 U+30B8 U+4E00 U+89A7
            RetiredDisplayNameFallback = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H4E00) & ChrW(&H89A7)
        Case "M-09"
            ' ナレ??ジ表示 = U+30CA U+30EC U+30C3 U+30B8 U+8868 U+793A
            RetiredDisplayNameFallback = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H8868) & ChrW(&H793A)
        Case "M-13"
            ' ファイル形式設?? = U+30D5 U+30A1 U+30A4 U+30EB U+5F62 U+5F0F U+8A2D U+5B9A
            RetiredDisplayNameFallback = ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H5F62) & ChrW(&H5F0F) & ChrW(&H8A2D) & ChrW(&H5B9A)
        Case "M-14"
            ' M-14 display name (ASCII-safe comment to avoid CP932 issues)
            RetiredDisplayNameFallback = ChrW(&H64CD) & ChrW(&H4F5C) & ChrW(&H30ED) & ChrW(&H30B0)
        Case Else
            RetiredDisplayNameFallback = ""
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0714] clsSetupOrchestrator.RetiredDisplayNameFallback EXIT-OK"  ' [ADR-0100]
End Function

' v2.3 fix (2026-05-30): if both the literal screenId (e.g. "M-10") and its
' display-name twin (e.g. "格納???設??") exist after EnsureSheets, the literal
' one is the leftover (created accidentally by an older EnsureSheetExists run
' or by a partially-completed install). Remove the literal so the next open
' starts clean. We only delete when the display-name sheet is confirmed
' present, so we never lose user data.
Private Sub DropOrphanLiteralScreenIdSheets(ByVal xlsmName As String, ByVal sheetList As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0715] clsSetupOrchestrator.DropOrphanLiteralScreenIdSheets ENTER"  ' [ADR-0100]
    On Error Resume Next
    If Len(sheetList) = 0 Then Exit Sub
    Dim ids() As String
    ids = Split(sheetList, "|")
    Application.DisplayAlerts = False
    Dim k As Long
    For k = LBound(ids) To UBound(ids)
        Dim sid As String
        sid = ids(k)
        If sid <> "LOG" Then
            Dim disp As String
            disp = ReadDisplayNameFromUiSeed(xlsmName, sid)
            If Len(disp) > 0 And disp <> sid Then
                Dim wsDisp As Worksheet
                Dim wsLit As Worksheet
                Set wsDisp = Nothing
                Set wsLit = Nothing
                Set wsDisp = ThisWorkbook.Worksheets(disp)
                Set wsLit = ThisWorkbook.Worksheets(sid)
                If Not wsDisp Is Nothing And Not wsLit Is Nothing Then
                    wsLit.Delete
                End If
            End If
        End If
    Next k
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0716] clsSetupOrchestrator.DropOrphanLiteralScreenIdSheets EXIT-OK"  ' [ADR-0100]
End Sub

' v2.3 (2026-05-27): remove default "SheetN" tabs from a fresh xlsm.
' Safe: only deletes if the name is "Sheet1".."Sheet9" and not in our managed
' set, and the workbook has at least one managed sheet (so we never empty it).
Private Sub DeleteDefaultLeftoverSheets(ByVal sheetList As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0717] clsSetupOrchestrator.DeleteDefaultLeftoverSheets ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim managed As String
    managed = "|" & sheetList & "|"
    Dim toDelete As Collection
    Set toDelete = New Collection
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim nm As String
        nm = ws.Name
        If Len(nm) >= 6 And Len(nm) <= 7 Then
            If Left(nm, 5) = "Sheet" Then
                Dim suf As String
                suf = Mid(nm, 6)
                If IsNumeric(suf) Then
                    If InStr(managed, "|" & nm & "|") = 0 Then
                        toDelete.Add ws
                    End If
                End If
            End If
        End If
    Next ws
    If toDelete.Count > 0 And ThisWorkbook.Worksheets.Count - toDelete.Count >= 1 Then
        Application.DisplayAlerts = False
        Dim j As Long
        For j = 1 To toDelete.Count
            toDelete.Item(j).Delete
        Next j
        Application.DisplayAlerts = True
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0718] clsSetupOrchestrator.DeleteDefaultLeftoverSheets EXIT-OK"  ' [ADR-0100]
End Sub

' v2.3 fix (2026-05-30): rewrite to consult ui_seed [SHEET].SheetName first
' so we never recreate the literal "M-NN" tab when the display-name twin
' (e.g. "格納???設??") already exists. Old behaviour resurrected literal
' tabs on every Workbook_Open and reproduced the split-state bug after
' the sheet audit had deleted them. Resolution order:
'   1. resolve display name from ui_seed/<role>/<screenId>.txt
'   2. if display-name sheet exists -> done
'   3. if literal screenId sheet exists -> rename to display name
'   4. otherwise add a new sheet and name it display-name (or screenId
'      when no display name is known, e.g. LOG)
Private Sub EnsureSheetExists(ByVal xlsmName As String, ByVal screenId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0719] clsSetupOrchestrator.EnsureSheetExists ENTER"  ' [ADR-0100]
    On Error Resume Next

    ' (1) resolve display name from ui_seed
    Dim displayName As String
    displayName = ReadDisplayNameFromUiSeed(xlsmName, screenId)
    If Len(displayName) = 0 Then displayName = screenId   ' LOG / unseeded fallback

    ' (2) sheet with display name already exists -> nothing to do
    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets(displayName)
    If Not ws Is Nothing Then
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0720] clsSetupOrchestrator.EnsureSheetExists EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    Err.Clear

    ' (3) literal screenId sheet exists -> rename to display name (idempotent
    ' when displayName == screenId, e.g. LOG)
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets(screenId)
    If Not ws Is Nothing Then
        If ws.Name <> displayName Then
            ws.Name = displayName
            If Err.Number <> 0 Then Err.Clear   ' rename conflict -> keep literal
        End If
        On Error GoTo 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0721] clsSetupOrchestrator.EnsureSheetExists EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    Err.Clear

    ' (4) neither exists -> create and name with display name
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = displayName
    If Err.Number <> 0 Then
        Err.Clear
        ' name collision (e.g. another sheet already owns displayName but
        ' Worksheets() lookup missed it due to case/whitespace) -> fall back
        ws.Name = screenId
        If Err.Number <> 0 Then Err.Clear
    End If
    On Error GoTo 0
End Sub

' v2.3 fix (2026-05-30): read [SHEET].SheetName from
' <ui_dir>\<xlsmName>\<screenId>.txt. Returns "" when the file is missing,
' unparseable, or has no SheetName key. xlsmName uses the JP folder names
' (管?? / 登録修正 / 検索) populated by _auto_install.ps1 Step 1.5.
Private Function ReadDisplayNameFromUiSeed(ByVal xlsmName As String, ByVal screenId As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0722] clsSetupOrchestrator.ReadDisplayNameFromUiSeed ENTER"  ' [ADR-0100]
    On Error GoTo EH
    ReadDisplayNameFromUiSeed = ""
    If screenId = "LOG" Then Exit Function

    Dim uiDir As String
    uiDir = modConfigHolder.GetUiDir()
    If Len(uiDir) = 0 Then Exit Function

    Dim filePath As String
    filePath = uiDir & xlsmName & "\" & screenId & ".txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function

    Dim secs As Collection
    Set secs = modStanzaIO.ParseStanzaFile(filePath)
    If secs Is Nothing Then Exit Function

    Dim j As Long
    Dim sec As ClsStanzaSection
    For j = 1 To secs.Count
        Set sec = secs.Item(j)
        If sec.SectionName = "SHEET" Then
            ReadDisplayNameFromUiSeed = Trim(sec.GetValue("SheetName"))
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0723] clsSetupOrchestrator.ReadDisplayNameFromUiSeed EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next j
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0724] clsSetupOrchestrator.ReadDisplayNameFromUiSeed EXIT-OK"  ' [ADR-0100]
    Exit Function
EH:
    ReadDisplayNameFromUiSeed = ""
End Function

Private Sub ApplyTabColors(ByVal xlsmName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0725] clsSetupOrchestrator.ApplyTabColors ENTER"  ' [ADR-0100]
    Dim sheetList As String
    sheetList = GetSheetsCsv(xlsmName)
    If Len(sheetList) = 0 Then
        modCommon.AppendProgressLog modCommon.ProgressTs() & "ApplyTabColors EXIT-EMPTY xlsm=" & xlsmName
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0726] clsSetupOrchestrator.ApplyTabColors EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    Dim tabColor As Long
    On Error Resume Next
    tabColor = GetTabColor(xlsmName)
    If Err.Number <> 0 Then
        modCommon.AppendProgressLog modCommon.ProgressTs() & "ApplyTabColors GetTabColor-FAIL err=" & Err.Number & " desc=" & Err.Description
        Err.Clear
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0727] clsSetupOrchestrator.ApplyTabColors EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    On Error GoTo 0

    Dim names() As String
    names = Split(sheetList, "|")

    modCommon.AppendProgressLog modCommon.ProgressTs() & "ApplyTabColors LOOP-START count=" & (UBound(names) - LBound(names) + 1)

    Dim i As Long
    ' [P5b fix 2026-06-04] resolve internal ID -> JP display name via ui_seed (ADR-0080)
    Dim resolvedName As String
    For i = LBound(names) To UBound(names)
        If names(i) <> "LOG" Then
            resolvedName = ReadDisplayNameFromUiSeed(xlsmName, names(i))
            If Len(resolvedName) = 0 Then resolvedName = names(i)
            modCommon.AppendProgressLog modCommon.ProgressTs() & "ApplyTabColors PRE sheet=" & names(i) & " resolved=" & resolvedName
            On Error Resume Next
            ThisWorkbook.Worksheets(resolvedName).Tab.Color = tabColor
            If Err.Number <> 0 Then
                modCommon.AppendProgressLog modCommon.ProgressTs() & "ApplyTabColors ERR sheet=" & names(i) & " resolved=" & resolvedName & " err=" & Err.Number & " desc=" & Err.Description
                Err.Clear
            Else
                modCommon.AppendProgressLog modCommon.ProgressTs() & "ApplyTabColors OK sheet=" & names(i) & " resolved=" & resolvedName
            End If
            On Error GoTo 0
        End If
    Next i

    modCommon.AppendProgressLog modCommon.ProgressTs() & "ApplyTabColors LOOP-END"
End Sub

Private Sub ApplyUiStanzas(ByVal xlsmName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0728] clsSetupOrchestrator.ApplyUiStanzas ENTER"  ' [ADR-0100]
    Dim sheetList As String
    sheetList = GetSheetsCsv(xlsmName)
    If Len(sheetList) = 0 Then
        modCommon.AppendProgressLog modCommon.ProgressTs() & "ApplyUiStanzas EXIT-EMPTY xlsm=" & xlsmName
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0729] clsSetupOrchestrator.ApplyUiStanzas EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    Dim names() As String
    names = Split(sheetList, "|")

    Dim renderer As IScreenRenderer
    Set renderer = New clsSheetRenderer

    ' Phase FT (2026-06-03 ADR-0097): per-sheet Debug.Print breadcrumb so install.log
    ' shows the exact sheet name in flight when a hang surfaces.
    modCommon.AppendProgressLog modCommon.ProgressTs() & "ApplyUiStanzas LOOP-START count=" & (UBound(names) - LBound(names) + 1)
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If names(i) <> "LOG" Then
            Debug.Print "[FT-APPLYUI-ENTER] xlsm=" & xlsmName & " sheet=" & names(i) & _
                        " ts=" & Format$(Now, "hh:nn:ss")
            modCommon.AppendProgressLog "[FT-APPLYUI-ENTER] xlsm=" & xlsmName & _
                        " sheet=" & names(i)
            On Error Resume Next
            renderer.BindSheet names(i)
            renderer.ClearScreen
            renderer.ApplyFromStanza xlsmName, names(i)
            renderer.ProtectSheet "light"
            Debug.Print "[FT-APPLYUI-EXIT] xlsm=" & xlsmName & " sheet=" & names(i) & _
                        " err=" & Err.Number
            modCommon.AppendProgressLog "[FT-APPLYUI-EXIT] xlsm=" & xlsmName & _
                        " sheet=" & names(i) & " err=" & Err.Number
            If Err.Number <> 0 Then
                If Not m_logger Is Nothing Then
                    m_logger.LogWarn "clsSetupOrchestrator", "ApplyUiStanzas", _
                        "UI ?K?p?G???[ " & names(i) & ": " & Err.Description, "", "LOG-SETUP-UI-WARN"
                End If
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next i
    modCommon.AppendProgressLog modCommon.ProgressTs() & "ApplyUiStanzas LOOP-END"
End Sub

Private Sub ApplyProtection()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0730] clsSetupOrchestrator.ApplyProtection ENTER"  ' [ADR-0100]
    On Error Resume Next
    ' 2026-06-07: removed Workbook.Protect Structure:=True so users can
    ' right-click sheet tab -> Unhide to reveal LOG sheet. Workbook structure
    ' protection blocks both hide/unhide and insert/delete/rename sheets;
    ' the user explicitly chose access to LOG over accidental-sheet-edit
    ' guardrails.
    ' ThisWorkbook.Protect Structure:=True, Windows:=False  ' DISABLED
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub ActivateStartupSheet(ByVal xlsmName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0731] clsSetupOrchestrator.ActivateStartupSheet ENTER"  ' [ADR-0100]
    Dim startupSheet As String
    Select Case xlsmName
        Case NAME_TOUROKU: startupSheet = STARTUP_SHEET_TOUROKU
        Case NAME_KENSAKU: startupSheet = STARTUP_SHEET_KENSAKU
        Case NAME_KANRI:   startupSheet = STARTUP_SHEET_KANRI
        Case Else: Exit Sub
    End Select

    On Error Resume Next
    Dim activated As Boolean
    activated = False
    ThisWorkbook.Worksheets(startupSheet).Activate
    If Err.Number = 0 Then activated = True
    Err.Clear
    On Error GoTo 0
    ' R-2-Fix2 (2026-05-28): startup id (M-prefixed screen identifier) may not match the actual JP sheet
    ' name, so Activate fails and the book opens on the trailing near-empty sheet.
    ' Fallback: activate the first non-LOG visible sheet.
    If Not activated Then
        Dim wsAct As Object
        For Each wsAct In ThisWorkbook.Worksheets
            If wsAct.Name <> "LOG" And wsAct.Visible = xlSheetVisible Then
                On Error Resume Next
                wsAct.Activate
                On Error GoTo 0
                Exit For
            End If
        Next wsAct
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0732] clsSetupOrchestrator.ActivateStartupSheet EXIT-OK"  ' [ADR-0100]
End Sub

' ----------------------------------------------------------------
' Private: ?????
' ----------------------------------------------------------------

Private Function GetSheetsCsv(ByVal xlsmName As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0733] clsSetupOrchestrator.GetSheetsCsv ENTER"  ' [ADR-0100]
    Select Case xlsmName
        Case NAME_TOUROKU: GetSheetsCsv = SHEETS_TOUROKU
        Case NAME_KENSAKU: GetSheetsCsv = SHEETS_KENSAKU
        Case NAME_KANRI:   GetSheetsCsv = SHEETS_KANRI
        Case Else:       ' [B3 fix] fallback to KANRI sheets so UI not vanish on unknown xlsmName
            GetSheetsCsv = "LOG"
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0734] clsSetupOrchestrator.GetSheetsCsv EXIT-OK"  ' [ADR-0100]
End Function

' Phase R-3-psi-Refresh: ThisWorkbook ?? (kanri.xlsm ??) から role 名を得る??
Private Function CurrentXlsmName() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0735] clsSetupOrchestrator.CurrentXlsmName ENTER"  ' [ADR-0100]
    Dim n As String
    n = ThisWorkbook.Name
    Dim p As Long
    p = InStrRev(n, ".")
    If p > 0 Then n = Left(n, p - 1)
    ' [ADR-0107 + B3 fix] StrComp(vbTextCompare) for safe normalization
    Dim k As String, t As String, s As String
    k = NAME_KANRI: t = NAME_TOUROKU: s = NAME_KENSAKU
    If StrComp(n, k, vbTextCompare) = 0 Then n = k
    If StrComp(n, t, vbTextCompare) = 0 Then n = t
    If StrComp(n, s, vbTextCompare) = 0 Then n = s
    CurrentXlsmName = n
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0736] clsSetupOrchestrator.CurrentXlsmName EXIT-OK"  ' [ADR-0100]
End Function

' Refresh ?? install 後に単独で呼ばれた場合、config が未ロードなら読み込む??
Private Sub EnsureConfigLoaded(ByVal xlsmName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0737] clsSetupOrchestrator.EnsureConfigLoaded ENTER"  ' [ADR-0100]
    On Error Resume Next
    If Len(modConfigHolder.GetUiDir()) = 0 Then
        modConfigLoader.LoadConfig xlsmName
    End If
    On Error GoTo 0
End Sub

' JP 表示?? -> screenId (M-prefixed screen identifier)。role の sheet 一覧の?? seed の [SHEET].SheetName と突合??
' ADR-0090 fix (2026-06-01): function body was a half-finished copy of
' ReadDisplayNameFromUiSeed - it referenced an undeclared filePath
' variable and never iterated through the screen list. Rewritten to
' scan each candidate screenId .txt for [SHEET].SheetName == sheetName.
Private Function ScreenIdForSheetName(ByVal sheetName As String, ByVal xlsmName As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0738] clsSetupOrchestrator.ScreenIdForSheetName ENTER"  ' [ADR-0100]
    On Error GoTo EH
    ScreenIdForSheetName = ""
    Dim sheetList As String
    sheetList = GetSheetsCsv(xlsmName)
    If Len(sheetList) = 0 Then Exit Function
    Dim names() As String
    names = Split(sheetList, "|")
    Dim uiDir As String
    uiDir = modConfigHolder.GetUiDir()
    If Len(uiDir) = 0 Then Exit Function
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim screenId As String
        screenId = Trim(names(i))
        If Len(screenId) > 0 Then
            Dim filePath As String
            filePath = uiDir & xlsmName & "\" & screenId & ".txt"
            If fso.FileExists(filePath) Then
                Dim secs As Collection
                Set secs = modStanzaIO.ParseStanzaFile(filePath)
                If Not secs Is Nothing Then
                    Dim j As Long
                    Dim sec As ClsStanzaSection
                    For j = 1 To secs.Count
                        Set sec = secs.Item(j)
                        If sec.SectionName = "SHEET" Then
                            If Trim(sec.GetValue("SheetName")) = sheetName Then
                                ScreenIdForSheetName = screenId
                                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0739] clsSetupOrchestrator.ScreenIdForSheetName EXIT-OK"  ' [ADR-0100]
                                Exit Function
                            End If
                        End If
                    Next j
                End If
            End If
        End If
    Next i
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0740] clsSetupOrchestrator.ScreenIdForSheetName EXIT-OK"  ' [ADR-0100]
    Exit Function
EH:
    ScreenIdForSheetName = ""
End Function
Private Function GetTabColor(ByVal xlsmName As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0741] clsSetupOrchestrator.GetTabColor ENTER"  ' [ADR-0100]
    Select Case xlsmName
        Case NAME_TOUROKU: GetTabColor = TAB_COLOR_TOUROKU
        Case NAME_KENSAKU: GetTabColor = TAB_COLOR_KENSAKU
        Case NAME_KANRI:   GetTabColor = TAB_COLOR_KANRI
        Case Else:       GetTabColor = 0
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0742] clsSetupOrchestrator.GetTabColor EXIT-OK"  ' [ADR-0100]
End Function

' ADR-0006/0090/0094 JP literal removal (Const -> Private Property Get).
Private Property Get STARTUP_SHEET_TOUROKU() As String
    STARTUP_SHEET_TOUROKU = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H767B) & ChrW(&H9332)
End Property

Private Property Get STARTUP_SHEET_KENSAKU() As String
    STARTUP_SHEET_KENSAKU = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H691C) & ChrW(&H7D22)
End Property

Private Property Get STARTUP_SHEET_KANRI() As String
    STARTUP_SHEET_KANRI = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H4E00) & ChrW(&H89A7)
End Property
```
