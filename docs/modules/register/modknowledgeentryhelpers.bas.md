---
title: modKnowledgeEntryHelpers.bas
description: modKnowledgeEntryHelpers.bas のソースコード（コピペ用）
---

# modKnowledgeEntryHelpers.bas

**配置先**: `登録修正.xlsm` 用の VBA モジュール
**種類**: 標準モジュール

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
    ' JP: title 'Error'
    TitleError = ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
End Function

Private Function TitleWarning() As String
    ' JP: title 'Warning'
    TitleWarning = ChrW(&H8B66) & ChrW(&H544A)
End Function

Private Function TitleInfo() As String
    ' JP: title 'Info'
    TitleInfo = ChrW(&H60C5) & ChrW(&H5831)
End Function

Private Function TitleConfirm() As String
    ' JP: title 'Confirm'
    TitleConfirm = ChrW(&H78BA) & ChrW(&H8A8D)
End Function

Private Function LblOperation() As String
    ' JP: label 'Operation: '
    LblOperation = ChrW(&H64CD) & ChrW(&H4F5C) & ChrW(&H540D) & ":" & " "
End Function

Private Function LblDetail() As String
    ' JP: label 'Detail: '
    LblDetail = ChrW(&H5185) & ChrW(&H5BB9) & ":" & " "
End Function

Private Function LblAction() As String
    ' JP: label 'Action: '
    LblAction = ChrW(&H5BFE) & ChrW(&H51E6) & ":" & " "
End Function

Private Function LblConfirmPrompt() As String
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
    Exit Function
ErrHandler:
    On Error Resume Next
    Set BuildLogger = New clsLogger
End Function

' ================================================================
' Public Function: GetDataFolder
' Return: data_dir resolved via modConfigHolder (config.txt SSOT).
'         Falls back to "" on error.
' ================================================================
Public Function GetDataFolder() As String
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
    On Error Resume Next
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowError", operation, detail, action)
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
    On Error Resume Next
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowWarning", operation, detail, action)
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
    On Error Resume Next
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowInfo", operation, detail, "")
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
    On Error GoTo ErrHandler
    If IsTestMode() Then
        Call LogDialogSuppressed("ConfirmAction", operation, detail, "auto-Yes")
        ConfirmAction = True
        Exit Function
    End If
    Dim body As String
    body = LblOperation() & operation & vbCrLf & _
           LblDetail() & detail & vbCrLf & _
           LblConfirmPrompt()
    Dim r As VbMsgBoxResult
    r = MsgBox(body, vbQuestion + vbYesNo, TitleConfirm())
    ConfirmAction = (r = vbYes)
    Exit Function
ErrHandler:
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
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = BuildLogger()
    Dim msg As String
    msg = "[" & dialogType & " suppressed] op=" & operation & _
          ", detail=" & detail
    If Len(action) > 0 Then msg = msg & ", action=" & action
    lg.LogInfo "modKnowledgeEntryHelpers", "LogDialogSuppressed", msg
    Err.Clear
End Sub
```
