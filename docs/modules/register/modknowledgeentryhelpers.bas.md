---
title: modKnowledgeEntryHelpers.bas
description: modKnowledgeEntryHelpers.bas のソースコード（コピペ用）
---

# modKnowledgeEntryHelpers.bas

**配置先**: `登録修正.xlsm` 用の VBA モジュール
**種類**: 標準モジュール
**更新日**: 2026-06-03 23:22

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\register\`
- ファイル名: `modKnowledgeEntryHelpers.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modKnowledgeEntryHelpers"
Option Explicit

' ================================================================
' Module: modKnowledgeEntryHelpers (register layer helper)
' Purpose: 7 helper functions required by modEntryKnowledge (BUG-3 fix)
'   - BuildLogger()       : new clsLogger + Init(LOG sheet)
'   - GetDataFolder()     : modConfigHolder.GetDataDir()
'   - IsTestMode()        : modCommon.IsHeadless() proxy
'   - ShowError(op, det, act)
'   - ShowWarning(op, det, act)
'   - ShowInfo(op, det)
'   - ConfirmAction(op, det) As Boolean
' Notes:
'   - JP literals encoded via ChrW() to keep this file ASCII-source safe
'     (per memory rule feedback_edit_tool_cp932_conversion).
'   - LOG sheet name is "LOG" (modCommon.SHEET_LOG constant exists).
'   - In headless/COM mode, dialogs are suppressed and routed to logger.
' ================================================================

Private Const SHEET_LOG_NAME As String = "LOG"

' ----------------------------------------------------------------
' Title strings (JP) built via ChrW so file stays ASCII source.
' ----------------------------------------------------------------
Private Function TitleError() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1747] modKnowledgeEntryHelpers.TitleError ENTER"  ' [ADR-0100]
    ' JP: title 'Error'
    TitleError = ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
End Function

Private Function TitleWarning() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1748] modKnowledgeEntryHelpers.TitleWarning ENTER"  ' [ADR-0100]
    ' JP: title 'Warning'
    TitleWarning = ChrW(&H8B66) & ChrW(&H544A)
End Function

Private Function TitleInfo() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1749] modKnowledgeEntryHelpers.TitleInfo ENTER"  ' [ADR-0100]
    ' JP: title 'Info'
    TitleInfo = ChrW(&H60C5) & ChrW(&H5831)
End Function

Private Function TitleConfirm() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1750] modKnowledgeEntryHelpers.TitleConfirm ENTER"  ' [ADR-0100]
    ' JP: title 'Confirm'
    TitleConfirm = ChrW(&H78BA) & ChrW(&H8A8D)
End Function

Private Function LblOperation() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1751] modKnowledgeEntryHelpers.LblOperation ENTER"  ' [ADR-0100]
    ' JP: label 'Operation: '
    LblOperation = ChrW(&H64CD) & ChrW(&H4F5C) & ChrW(&H540D) & ":" & " "
End Function

Private Function LblDetail() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1752] modKnowledgeEntryHelpers.LblDetail ENTER"  ' [ADR-0100]
    ' JP: label 'Detail: '
    LblDetail = ChrW(&H5185) & ChrW(&H5BB9) & ":" & " "
End Function

Private Function LblAction() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1753] modKnowledgeEntryHelpers.LblAction ENTER"  ' [ADR-0100]
    ' JP: label 'Action: '
    LblAction = ChrW(&H5BFE) & ChrW(&H51E6) & ":" & " "
End Function

Private Function LblConfirmPrompt() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1754] modKnowledgeEntryHelpers.LblConfirmPrompt ENTER"  ' [ADR-0100]
    ' JP: prompt 'Are you sure to proceed?'
    LblConfirmPrompt = ChrW(&H5B9F) & ChrW(&H884C) & ChrW(&H3057) & ChrW(&H3066) & _
                      ChrW(&H3082) & ChrW(&H3088) & ChrW(&H308D) & ChrW(&H3057) & _
                      ChrW(&H3044) & ChrW(&H3067) & ChrW(&H3059) & ChrW(&H304B) & "?"
End Function

' ================================================================
' Public Function: BuildLogger
' Return: initialized clsLogger bound to ThisWorkbook "LOG" sheet.
'         If LOG sheet missing or Init fails, returns a logger object
'         whose calls become no-ops (On Error Resume Next inside caller).
' ================================================================
Public Function BuildLogger() As clsLogger
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1755] modKnowledgeEntryHelpers.BuildLogger ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim lg As clsLogger
    Set lg = New clsLogger
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG_NAME)
    On Error GoTo ErrHandler
    If Not ws Is Nothing Then
        ' iter20 fix: in headless E2E COM mode the modConfigHolder is not
        ' initialized via Workbook_Open (EnableEvents=false), so debug
        ' level defaults to ERROR and LogInfo entries are silently dropped.
        ' Override to INFO when running headless so LOG-effect test
        ' assertions can observe entry-point markers.
        If modCommon.IsHeadless() Then
            lg.Init ws, 3  ' DEBUG_LEVEL_INFO
        Else
            lg.Init ws
        End If
    End If
    Set BuildLogger = lg
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1756] modKnowledgeEntryHelpers.BuildLogger EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1757] modKnowledgeEntryHelpers.BuildLogger EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    On Error Resume Next
    Set BuildLogger = New clsLogger
End Function

' ================================================================
' Public Function: GetDataFolder
' Return: data_dir resolved via modConfigHolder (config.txt SSOT).
'         Falls back to "" on error.
' ================================================================
Public Function GetDataFolder() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1758] modKnowledgeEntryHelpers.GetDataFolder ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    GetDataFolder = modConfigHolder.GetDataDir()
    Exit Function
ErrHandler:
    GetDataFolder = ""
End Function

' ================================================================
' Public Function: IsTestMode
' Return: True when running headless (E2E COM mode) so dialog
'         suppression behaves like the legacy modEntryMain.IsTestMode().
' ================================================================
Public Function IsTestMode() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1759] modKnowledgeEntryHelpers.IsTestMode ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    IsTestMode = modCommon.IsHeadless()
    Exit Function
ErrHandler:
    IsTestMode = False
End Function

' ================================================================
' Public Sub: ShowError
' ================================================================
Public Sub ShowError(ByVal operation As String, ByVal detail As String, ByVal action As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1760] modKnowledgeEntryHelpers.ShowError ENTER"  ' [ADR-0100]
    On Error Resume Next
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowError", operation, detail, action)
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1761] modKnowledgeEntryHelpers.ShowError EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    Dim body As String
    body = LblOperation() & operation & vbCrLf & _
           LblDetail() & detail & vbCrLf & _
           LblAction() & action
    MsgBox body, vbCritical, TitleError()
End Sub

' ================================================================
' Public Sub: ShowWarning
' ================================================================
Public Sub ShowWarning(ByVal operation As String, ByVal detail As String, ByVal action As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1762] modKnowledgeEntryHelpers.ShowWarning ENTER"  ' [ADR-0100]
    On Error Resume Next
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowWarning", operation, detail, action)
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1763] modKnowledgeEntryHelpers.ShowWarning EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    Dim body As String
    body = LblOperation() & operation & vbCrLf & _
           LblDetail() & detail & vbCrLf & _
           LblAction() & action
    MsgBox body, vbExclamation, TitleWarning()
End Sub

' ================================================================
' Public Sub: ShowInfo
' ================================================================
Public Sub ShowInfo(ByVal operation As String, ByVal detail As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1764] modKnowledgeEntryHelpers.ShowInfo ENTER"  ' [ADR-0100]
    On Error Resume Next
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowInfo", operation, detail, "")
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1765] modKnowledgeEntryHelpers.ShowInfo EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    Dim body As String
    body = LblOperation() & operation & vbCrLf & _
           LblDetail() & detail
    MsgBox body, vbInformation, TitleInfo()
End Sub

' ================================================================
' Public Function: ConfirmAction
' Return: True if user (or auto in test mode) confirms.
' ================================================================
Public Function ConfirmAction(ByVal operation As String, ByVal detail As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1766] modKnowledgeEntryHelpers.ConfirmAction ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If IsTestMode() Then
        Call LogDialogSuppressed("ConfirmAction", operation, detail, "auto-Yes")
        ConfirmAction = True
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1767] modKnowledgeEntryHelpers.ConfirmAction EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim body As String
    body = LblOperation() & operation & vbCrLf & _
           LblDetail() & detail & vbCrLf & _
           LblConfirmPrompt()
    Dim r As VbMsgBoxResult
    r = MsgBox(body, vbQuestion + vbYesNo, TitleConfirm())
    ConfirmAction = (r = vbYes)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1768] modKnowledgeEntryHelpers.ConfirmAction EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1769] modKnowledgeEntryHelpers.ConfirmAction EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ConfirmAction = False
End Function

' ================================================================
' Private Sub: LogDialogSuppressed
' Notes:   Mirrors legacy modEntryMain behavior; logs that a dialog
'          would have appeared while headless.
' ================================================================
Private Sub LogDialogSuppressed(ByVal dialogType As String, _
                                ByVal operation As String, _
                                ByVal detail As String, _
                                ByVal action As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1770] modKnowledgeEntryHelpers.LogDialogSuppressed ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = BuildLogger()
    Dim msg As String
    msg = "[" & dialogType & " suppressed] op=" & operation & _
          ", detail=" & detail
    If Len(action) > 0 Then msg = msg & ", action=" & action
    lg.LogInfo "modKnowledgeEntryHelpers", "LogDialogSuppressed", msg
    Err.Clear
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1771] modKnowledgeEntryHelpers.LogDialogSuppressed EXIT-OK"  ' [ADR-0100]
End Sub
```
