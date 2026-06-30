---
title: modBtnGuard.bas
description: modBtnGuard.bas のソースコード（コピペ用）
---

# modBtnGuard.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modBtnGuard.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modBtnGuard"
Option Explicit

' ============================================================
' modBtnGuard
' Role: 全 Btn_*** ハンドラの共通ガード層。
'       前提条件 (config / sheet / format / 入力) を一括 check、
'       失敗時は自然言語メッセージで MsgBox + clsLogger.LogError、
'       ボタン進捗を Debug.Print + clsLogger に並走出力。
' Coding rule: ADR-0094 D1 (CP932 strict + Python script for write)
' ============================================================

Private Const MOD_NAME As String = "modBtnGuard"

' ------------------------------------------------------------
' Public Function CheckPrereq
' Args:
'   btnName      - 呼出ボタン名 (例: "Btn_RegisterNew")
'   requirements - 必要な前提条件を ; 区切りで列挙
'                  "config" / "sheet:<name>" / "format:<id>"
'                  "input:<cellAddr>" / "data_dir" / "format_dir"
'   xlsmName     - 呼出ブック名 ("管理"/"登録修正"/"検索")
' Returns: True = OK, False = NG (内部で MsgBox + LogError 済)
' ------------------------------------------------------------
Public Function CheckPrereq(ByVal btnName As String, _
                             ByVal requirements As String, _
                             ByVal xlsmName As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0871] modBtnGuard.CheckPrereq ENTER"  ' [ADR-0100]
    Dim reqArr() As String
    Dim i As Long
    Dim req As String
    Dim okAll As Boolean
    okAll = True

    If Len(Trim(requirements)) = 0 Then
        CheckPrereq = True
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0872] modBtnGuard.CheckPrereq EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    reqArr = Split(requirements, ";")
    For i = LBound(reqArr) To UBound(reqArr)
        req = Trim(reqArr(i))
        If Len(req) > 0 Then
            If Not CheckSingleReq(btnName, req, xlsmName) Then
                okAll = False
                Exit For
            End If
        End If
    Next i

    CheckPrereq = okAll
End Function

' ------------------------------------------------------------
' Private CheckSingleReq
' 単一前提条件の評価。NG 時は MsgBox + LogError + Debug.Print して False。
' ------------------------------------------------------------
Private Function CheckSingleReq(ByVal btnName As String, _
                                 ByVal req As String, _
                                 ByVal xlsmName As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0873] modBtnGuard.CheckSingleReq ENTER"  ' [ADR-0100]
    Dim msg As String
    Dim colonPos As Long
    Dim kind As String
    Dim param As String
    Dim ok As Boolean
    Dim dataDir As String
    Dim fmtDir As String
    ok = True

    colonPos = InStr(1, req, ":")
    If colonPos > 0 Then
        kind = LCase(Trim(Left(req, colonPos - 1)))
        param = Trim(Mid(req, colonPos + 1))
    Else
        kind = LCase(Trim(req))
        param = ""
    End If

    Select Case kind
        Case "config"
            If Not IsConfigLoaded() Then
                ' 2026-06-07: lazy-load config when the holder is empty
                ' (UserForm Unload resets module-level state in some setups).
                On Error Resume Next
                modConfigLoader.LoadConfig xlsmName
                On Error GoTo 0
            End If
            If Not IsConfigLoaded() Then
                msg = modBtnMessages.GetMessage("MSG-BTN-PRE-001")
                ReportPrereqFail btnName, req, msg, xlsmName
                ok = False
            End If

        Case "sheet"
            If Not SheetExists(param) Then
                msg = modBtnMessages.GetMessage("MSG-BTN-PRE-002", param)
                ReportPrereqFail btnName, req, msg, xlsmName
                ok = False
            End If

        Case "format"
            If Not FormatIdExists(param) Then
                msg = modBtnMessages.GetMessage("MSG-BTN-PRE-003", param)
                ReportPrereqFail btnName, req, msg, xlsmName
                ok = False
            End If

        Case "input"
            If Not InputCellHasValue(param) Then
                msg = modBtnMessages.GetMessage("MSG-BTN-PRE-004", param)
                ReportPrereqFail btnName, req, msg, xlsmName
                ok = False
            End If

        Case "data_dir"
            dataDir = SafeGetDataDir()
            If Len(dataDir) = 0 Or Not FolderExists(dataDir) Then
                msg = modBtnMessages.GetMessage("MSG-BTN-PRE-005", dataDir)
                ReportPrereqFail btnName, req, msg, xlsmName
                ok = False
            End If

        Case "format_dir"
            fmtDir = SafeGetFormatDir()
            If Len(fmtDir) = 0 Or Not FolderExists(fmtDir) Then
                msg = modBtnMessages.GetMessage("MSG-BTN-PRE-006", fmtDir)
                ReportPrereqFail btnName, req, msg, xlsmName
                ok = False
            End If

        Case Else
            ' 未知の req は安全側 OK (warn 出力)
            Debug.Print "[BTN-PRE-WARN] " & btnName & " unknown_req=" & req
    End Select

    CheckSingleReq = ok
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0874] modBtnGuard.CheckSingleReq EXIT-OK"  ' [ADR-0100]
End Function

' ------------------------------------------------------------
' Private ReportPrereqFail
' MsgBox (Headless 抑止) + Debug.Print + LogError 三点出力
' ------------------------------------------------------------
Private Sub ReportPrereqFail(ByVal btnName As String, _
                              ByVal req As String, _
                              ByVal msg As String, _
                              ByVal xlsmName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0875] modBtnGuard.ReportPrereqFail ENTER"  ' [ADR-0100]
    Dim title As String
    title = ChrW(&H524D) & ChrW(&H63D0) & ChrW(&H6761) & ChrW(&H4EF6) & ChrW(&H672A) & ChrW(&H5145) & ChrW(&H8DB3) & ": " & btnName
    Debug.Print "[BTN-PRE-NG] " & btnName & " req=" & req & " msg=" & msg
    SafeLogError MOD_NAME, btnName, "PRE-NG req=" & req & " : " & msg
    SafeAppendProgress "[BTN-PRE-NG] " & xlsmName & "/" & btnName & " req=" & req
    If Not modCommon.IsHeadless() Then
        MsgBox msg, vbExclamation, title
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0876] modBtnGuard.ReportPrereqFail EXIT-OK"  ' [ADR-0100]
End Sub

' ------------------------------------------------------------
' Public Sub HandleButtonError
' Btn_*** の ErrHandler から呼ぶ。raw err を握って自然言語化。
' ------------------------------------------------------------
Public Sub HandleButtonError(ByVal btnName As String, _
                              ByRef errInfo As ErrObject, _
                              ByVal xlsmName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0877] modBtnGuard.HandleButtonError ENTER"  ' [ADR-0100]
    Dim friendly As String
    Dim shown As String
    Dim errNum As Long
    Dim errDesc As String
    errNum = errInfo.Number
    errDesc = errInfo.Description
    friendly = modBtnMessages.GetErrorMessage(errNum)
    shown = friendly & vbCrLf & vbCrLf & _
            "(" & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC) & ChrW(&H8A73) & ChrW(&H7D30) & ": " & errDesc & ")"
    Debug.Print "[BTN-ERR] " & btnName & " errNum=" & errNum & " " & errDesc
    SafeLogError MOD_NAME, btnName, "BTN-ERR errNum=" & errNum & " desc=" & errDesc
    SafeAppendProgress "[BTN-ERR] " & xlsmName & "/" & btnName & " errNum=" & errNum
    If Not modCommon.IsHeadless() Then
        MsgBox shown, vbCritical, ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC) & ": " & btnName
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0878] modBtnGuard.HandleButtonError EXIT-OK"  ' [ADR-0100]
End Sub

' ------------------------------------------------------------
' Public Sub LogStep / LogEnter / LogExit
' ボタン処理進捗の三点並走 (Debug.Print + clsLogger + ProgressLog)
' ------------------------------------------------------------
Public Sub LogStep(ByVal btnName As String, ByVal stepName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0879] modBtnGuard.LogStep ENTER"  ' [ADR-0100]
    Dim ts As String
    ts = Format$(Now, "hh:nn:ss")
    Debug.Print "[STEP] " & btnName & "." & stepName & " ts=" & ts
    SafeLogInfo MOD_NAME, btnName, "STEP " & stepName
    SafeAppendProgress "[STEP] " & btnName & "." & stepName
End Sub

Public Sub LogEnter(ByVal btnName As String, ByVal xlsmName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0880] modBtnGuard.LogEnter ENTER"  ' [ADR-0100]
    Dim ts As String
    ts = Format$(Now, "hh:nn:ss")
    Debug.Print "[ENTER] " & xlsmName & "/" & btnName & " ts=" & ts
    SafeLogInfo MOD_NAME, btnName, "ENTER xlsm=" & xlsmName
    SafeAppendProgress "[ENTER] " & xlsmName & "/" & btnName
End Sub

Public Sub LogExit(ByVal btnName As String, ByVal xlsmName As String, ByVal okFlag As Boolean)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0881] modBtnGuard.LogExit ENTER"  ' [ADR-0100]
    Dim ts As String
    Dim verdict As String
    ts = Format$(Now, "hh:nn:ss")
    If okFlag Then
        verdict = "OK"
    Else
        verdict = "NG"
    End If
    Debug.Print "[EXIT] " & xlsmName & "/" & btnName & " verdict=" & verdict & " ts=" & ts
    SafeLogInfo MOD_NAME, btnName, "EXIT verdict=" & verdict & " xlsm=" & xlsmName
    SafeAppendProgress "[EXIT] " & xlsmName & "/" & btnName & " " & verdict
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0882] modBtnGuard.LogExit EXIT-OK"  ' [ADR-0100]
End Sub

' ============================================================
' Private helpers
' ============================================================

Private Function IsConfigLoaded() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0883] modBtnGuard.IsConfigLoaded ENTER"  ' [ADR-0100]
    On Error Resume Next
    IsConfigLoaded = modConfigHolder.IsInitialized()
    If Err.Number <> 0 Then
        IsConfigLoaded = False
        Err.Clear
    End If
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0884] modBtnGuard.IsConfigLoaded EXIT-OK"  ' [ADR-0100]
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0885] modBtnGuard.SheetExists ENTER"  ' [ADR-0100]
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function FormatIdExists(ByVal formatId As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0886] modBtnGuard.FormatIdExists ENTER"  ' [ADR-0100]
    On Error Resume Next
    FormatIdExists = modFormatLoader.FormatExists(formatId)
    If Err.Number <> 0 Then
        FormatIdExists = False
        Err.Clear
    End If
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0887] modBtnGuard.FormatIdExists EXIT-OK"  ' [ADR-0100]
End Function

Private Function InputCellHasValue(ByVal cellAddr As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0888] modBtnGuard.InputCellHasValue ENTER"  ' [ADR-0100]
    Dim v As Variant
    On Error Resume Next
    v = ActiveSheet.Range(cellAddr).Value
    If Err.Number <> 0 Then
        InputCellHasValue = False
        Err.Clear
        On Error GoTo 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0889] modBtnGuard.InputCellHasValue EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    On Error GoTo 0
    If IsNull(v) Or IsEmpty(v) Then
        InputCellHasValue = False
    ElseIf Len(Trim(CStr(v))) = 0 Then
        InputCellHasValue = False
    Else
        InputCellHasValue = True
    End If
End Function

Private Function FolderExists(ByVal path As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0890] modBtnGuard.FolderExists ENTER"  ' [ADR-0100]
    On Error Resume Next
    FolderExists = (Len(Dir(path, vbDirectory)) > 0)
    If Err.Number <> 0 Then
        FolderExists = False
        Err.Clear
    End If
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0891] modBtnGuard.FolderExists EXIT-OK"  ' [ADR-0100]
End Function

Private Function SafeGetDataDir() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0892] modBtnGuard.SafeGetDataDir ENTER"  ' [ADR-0100]
    On Error Resume Next
    SafeGetDataDir = modConfigHolder.GetDataDir()
    If Err.Number <> 0 Then
        SafeGetDataDir = ""
        Err.Clear
    End If
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0893] modBtnGuard.SafeGetDataDir EXIT-OK"  ' [ADR-0100]
End Function

Private Function SafeGetFormatDir() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0894] modBtnGuard.SafeGetFormatDir ENTER"  ' [ADR-0100]
    On Error Resume Next
    SafeGetFormatDir = modConfigHolder.GetFormatDir()
    If Err.Number <> 0 Then
        SafeGetFormatDir = ""
        Err.Clear
    End If
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0895] modBtnGuard.SafeGetFormatDir EXIT-OK"  ' [ADR-0100]
End Function

Private Sub SafeLogError(ByVal modName As String, ByVal funcName As String, ByVal msg As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0896] modBtnGuard.SafeLogError ENTER"  ' [ADR-0100]
    On Error Resume Next
    If Not gLogger Is Nothing Then
        gLogger.LogError modName, funcName, msg
    End If
    On Error GoTo 0
End Sub

Private Sub SafeLogInfo(ByVal modName As String, ByVal funcName As String, ByVal msg As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0897] modBtnGuard.SafeLogInfo ENTER"  ' [ADR-0100]
    On Error Resume Next
    If Not gLogger Is Nothing Then
        gLogger.LogInfo modName, funcName, msg
    End If
    On Error GoTo 0
End Sub

Private Sub SafeAppendProgress(ByVal msg As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0898] modBtnGuard.SafeAppendProgress ENTER"  ' [ADR-0100]
    On Error Resume Next
    modCommon.AppendProgressLog msg
    Err.Clear
    On Error GoTo 0
End Sub
```
