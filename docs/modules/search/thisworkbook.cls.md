---
title: ThisWorkbook.cls
description: ThisWorkbook.cls のソースコード（コピペ用）
---

# ThisWorkbook.cls

**配置先**: `検索.xlsm` 用の VBA モジュール  
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\search\`
- ファイル名: `ThisWorkbook.cls`
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
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ================================================================
' ThisWorkbook ?????W???[?? (???.xlsm ??p)
' ?z?u:   ThisWorkbook.cls ???????.xlsm ?? VBE ????? Import (VBE ?d?l?? Import ?s??A?R?s?y???)
' ??A:   clsSetupOrchestrator.bas v2, modConfigLoader.bas, modConfigHolder.bas
' Version: v2.1 (2026-05-16 EOD?AQ1-Q57 ???f)
' Phase: 7
' ADR:   ADR-0053 ??2.1
' v2.1:  ???.xlsm ???\?? = M-08 (1 ???) + LOG  (v2.3 retired M-07/M-09)
'        Q44: ?N???? ActiveSheet = M-08 (?????????, STARTUP_SHEET)
'        Q19: format_dir ??? write/delete ????.xlsm ???? (modFormatLoader ???? ThisWorkbook.Name enforce)
'        Q39: config.txt ??e?L?X?g?G?f?B?^?????W?AWorkbook_SheetChange ?n???h???s?v?iM-11 ??????????EdebugLevel ??W GUI?ASSOT-Q22?j
' ================================================================
Option Explicit

Private Const XLSM_NAME As String = "検索"
Private Const STARTUP_SHEET As String = "ナレッジ検索"  ' v2.1 Q44 ?m?? (M-01 ???j???[???A?N???? = ????????s)

' ================================================================
' Workbook_Open
' ?T?v:   xlsm ?N?????? setup ?????s (???.xlsm ?p)
' ??y:   1. modConfigLoader ?? xlsm ??????? config.txt ?? read ?? modConfigHolder ??Z?b?g (Q8)
'         2. clsLogger.Init (???O?V?[?g + debugLevel ERROR ????AQ7)
'         3. modKnowledgeFileIO.CleanupOldBackups ?? 90 ???? backup ?????? (Q34)
'         4. clsSetupOrchestrator.RunFullSetup("???")
'            - ?V?[?g?m?? (M-08/LOG)
'            - UI ?X?^???U?K?p (modUILoader.ApplyUiToSheet?AQ20)
'            - ?^?u?F (ADR-0053 ??2.1.1)
'            - format_dir ?????? (???? .txt ??????????t?H???_ seed)
'            - Workbook.Protect Structure + ?V?[?g???
'            - ActiveSheet = M-08 (?????????, STARTUP_SHEET) (Q44)
' ================================================================
Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' v2.1 Q34: ?N?????? 90 ?????o?b?N?A?b?v???????? (???.xlsm ?????{)
    On Error Resume Next
    Call modKnowledgeFileIO.CleanupOldBackups
    On Error GoTo ErrHandler

    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    Call orch.RunFullSetup(XLSM_NAME)

    ' S5-LOG-02: SAVE-EXIT-OK-II-003 (Workbook_Open success exit, screen ???)
    On Error Resume Next
    Dim oLogger003 As clsLogger
    Set oLogger003 = New clsLogger
    oLogger003.Init ThisWorkbook.Worksheets("LOG")
    oLogger003.LogInfo "ThisWorkbook_kensaku", "Workbook_Open", "Workbook_Open ??????: " & XLSM_NAME, "", "SAVE-EXIT-OK-II-003"
    On Error GoTo 0

    ' Q44: startup ActiveSheet = STARTUP_SHEET (production spec mirror)
    On Error Resume Next
    ThisWorkbook.Worksheets(STARTUP_SHEET).Activate
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Dim msg As String
    msg = "????.xlsm ?N???G???[: " & Err.Description & vbCrLf & _
          "config.txt ????? / debugLevel ?l / ?V?[?g?\?? / format_dir ???????? ???m?F???????????"
    Debug.Print msg
    ' S5-LOG: BACKTOMAIN-ERR-EE-031 (Workbook_Open failure, ????????B?s?? = back-to-main ?????J??G???[)
    On Error Resume Next
    Dim oLogger031 As clsLogger
    Set oLogger031 = New clsLogger
    oLogger031.Init ThisWorkbook.Worksheets("LOG")
    oLogger031.LogError "ThisWorkbook_kensaku", "Workbook_Open", "Workbook_Open ???s: " & Err.Description, "", "BACKTOMAIN-ERR-EE-031"
    On Error GoTo 0
    ' v2.3 install_admin.bat ?? headless ???@?e?X?g?? Setup_admin ??
    ' ???? MsgBox ???o????n???O????????P?v???B
    ' modCommon.IsHeadless() ?? COM ???????s?????o???A???????
    ' MsgBox ???o???? clsLogger / Debug.Print ?????m????B
    ' Application.Run "Setup_admin" ?o?H??? Workbook_Open ??N??
    ' ?????????????A???????L?????o?H??N????????????S??B
    If Not modCommon.IsHeadless() Then
        MsgBox msg, vbCritical, "Workbook_Open"
    Else
        Debug.Print "[HEADLESS] suppressed MsgBox: " & msg
        modCommon.AppendProgressLog modCommon.ProgressTs() & "ThisWorkbook(???).Workbook_Open ErrHandler suppressed MsgBox: " & Err.Description
    End If
End Sub

' ================================================================
' Workbook_SheetBeforeDoubleClick (v2.3 Phase O-1, 2026-05-27)
' M-08 ????????O???b?h??i???b?W????Z?? (column 1 = A) ???_?u???N???b?N
' ????? modEntryUserForm.OpenViewWithId ?????? M-09 ?i???b?W?\??
' ?t?H?[?????J???B
' ?d?l: ??v?? v2.3 Accepted "M-08 ?i???b?W????_???l" B53 (dblClickKnowledgeNo)?B
' ??? sheet: ?\???? "?i???b?W????" ????? ID "M-08"?B
' ?????: ?? A?A?f?[?^?s (14 ?s???~?B13 ?s???? header)?B
' ================================================================
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrHandler
    Dim nm As String
    nm = Sh.Name
    Dim isSearchSheet As Boolean
    isSearchSheet = (nm = "M-08" Or nm = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H691C) & ChrW(&H7D22))
    If Not isSearchSheet Then Exit Sub

    ' Phase R-3-Omega: 7-col grid per SSOT. knowledgeNo column is now B (col 2);
    ' col A = No (sequential). Data rows = 14+ (header up to row 13).
    ' 2026-05-31 UX: relax column gate to A..G so DoubleClick on any cell
    ' inside a result row opens the same kid view (B column value is still
    ' the authoritative kid source).
    If Target.Row < 14 Then Exit Sub
    If Target.Column < 1 Or Target.Column > 7 Then Exit Sub
    Dim kid As String
    kid = Trim$(CStr(Sh.Cells(Target.Row, 2).Value))
    If Len(kid) = 0 Then Exit Sub

    Cancel = True   ' ?Z????W???[?h??}?~
    modEntryUserForm.OpenViewWithId kid
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Workbook_SheetBeforeDoubleClick: " & Err.Number & " " & Err.Description
End Sub

' ================================================================
' Workbook_BeforeClose
' ================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("LOG")
    If Not ws Is Nothing Then
        Dim r As Long
        r = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
        If r < 9 Then r = 9
        ws.Cells(r, 1).Value = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
        ws.Cells(r, 2).Value = "ThisWorkbook"
        ws.Cells(r, 3).Value = "BeforeClose"
        ws.Cells(r, 4).Value = "INFO"
        ws.Cells(r, 5).Value = "xlsm ?I??: " & XLSM_NAME
    End If
    On Error GoTo 0
End Sub
```
