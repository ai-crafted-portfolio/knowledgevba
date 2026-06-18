---
title: modE2E_Environment.bas
description: modE2E_Environment.bas のソースコード（コピペ用）
---

# modE2E_Environment.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-08 12:53

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modE2E_Environment.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modE2E_Environment"
Option Explicit

Public Function Run_Environment_Cases(ByVal role As String) As String
    Dim sb As String, n As Long, sep As String
    sb = "{""role"":""env-" & role & """,""cases"":["
    sep = ""
    sb = sb & sep & EnvCase("env-01", "excel_version", CStr(Application.Version)): sep = ","
    sb = sb & sep & EnvCase("env-02", "excel_build", CStr(Application.Build))
    sb = sb & sep & EnvCase("env-03", "os_name", CStr(Application.OperatingSystem))
    sb = sb & sep & EnvCase("env-04", "user_name", Environ$("USERNAME"))
    sb = sb & sep & EnvCase("env-05", "computer_name", Environ$("COMPUTERNAME"))
    sb = sb & sep & EnvCase("env-06", "system_root", Environ$("SystemRoot"))
    sb = sb & sep & EnvCase("env-07", "temp_dir", Environ$("TEMP"))
    sb = sb & sep & EnvCase("env-08", "user_profile", Environ$("USERPROFILE"))
    sb = sb & sep & EnvCase("env-09", "calculation_mode", CStr(Application.Calculation))
    sb = sb & sep & EnvCase("env-10", "screen_updating", CStr(Application.ScreenUpdating))
    sb = sb & sep & EnvCase("env-11", "enable_events", CStr(Application.EnableEvents))
    sb = sb & sep & EnvCase("env-12", "workbook_count", CStr(Application.Workbooks.Count))
    sb = sb & sep & EnvCase("env-13", "active_printer", CStr(Application.ActivePrinter))
    sb = sb & sep & EnvCase("env-14", "cursor_movement", CStr(Application.CursorMovement))
    sb = sb & sep & EnvCase("env-15", "this_workbook", ThisWorkbook.Name)
    sb = sb & "]}"
    Run_Environment_Cases = sb
End Function

Private Function EnvCase(ByVal id As String, ByVal name As String, ByVal observed As String) As String
    Dim p As String
    If Len(observed) > 0 Then p = "true" Else p = "false"
    EnvCase = "{""id"":""" & id & """,""name"":""" & name & """,""pass"":" & p & ",""note"":""" & Replace(observed, """", "'") & """}"
End Function

Public Sub Run_Environment_Cases_Out(ByVal outPath As String)
    Dim role As String, s As String, fh As Integer
    role = Replace(ThisWorkbook.Name, ".xlsm", "")
    s = Run_Environment_Cases(role)
    fh = FreeFile
    Open outPath For Output As #fh
    Print #fh, s
    Close #fh
End Sub
```
