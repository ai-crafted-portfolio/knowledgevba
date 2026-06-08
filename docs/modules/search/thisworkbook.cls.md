---
title: ThisWorkbook.cls
description: ThisWorkbook.cls のソースコード（コピペ用）
---

# ThisWorkbook.cls

**配置先**: `検索.xlsm` 用の VBA モジュール
**種類**: クラスモジュール

---

## 取り込み方法（ThisWorkbook は特別）

**【重要】`ThisWorkbook` は Excel のブック（.xlsm）に最初から存在する document module です。ファイルとして Import できません。**

1. VBE（`Alt + F11`）を開き、プロジェクトツリーの `ThisWorkbook` を **ダブルクリック**してコードペインを開きます。
2. コードペインの**既存コードをすべて削除**します。
3. 下のコード全文をコピーして**貼り付け**ます。

> VB クラスモジュールヘッダ（`VERSION 1.0 CLASS` / `BEGIN` / `END` / `Attribute` 行）は不要です。本体は `Option Explicit` から始まります。
> 文字コード・改行はコードペインに直接貼り付けるため、ファイル保存時の ANSI/CRLF 指定は不要です。

---

## ソースコード


```vb
' ================================================================
' ThisWorkbook 専用モジュール (検索.xlsm 専用)
' 配置:   ThisWorkbook.cls として検索.xlsm の VBE に直接 Import (VBE 仕様で Import 不可、コピペのみ)
' 関連:   clsSetupOrchestrator.bas v2, modConfigLoader.bas, modConfigHolder.bas
' Version: v2.3 (2026-06-02 mojibake 修復後再生成)
' Phase: 7 / R3-Omega
' ADR:   ADR-0053 §2.1
' 画面構成: M-08 (1 画面) + LOG  (v2.3 M-07/M-09 廃止)
'        Q44: 起動時 ActiveSheet = ナレッジ検索 (M-08)
'        Q19: format_dir への write/delete は管理.xlsm 専用 (検索.xlsm からは read のみ)
'        Q39: config.txt はテキストエディタ直接編集、Workbook_SheetChange ハンドラ不要
' ================================================================
Option Explicit

Private Const XLSM_NAME As String = "検索"
Private Const STARTUP_SHEET As String = "ナレッジ検索"  ' v2.1 Q44 確定 (M-01 メニュー削除、起動時 = 業務画面直行)

' ================================================================
' Workbook_Open
' 概要:   xlsm 起動時に setup を実行 (検索.xlsm 用)
' 手順:   1. modConfigLoader が xlsm 名対応の config.txt を read → modConfigHolder にセット (Q8)
'         2. clsLogger.Init (ログシート + debugLevel ERROR 既定、Q7)
'         3. modKnowledgeFileIO.CleanupOldBackups で 90 日超 backup 自動削除 (Q34)
'         4. clsSetupOrchestrator.RunFullSetup("検索")
'            - シート確認 (M-08/LOG)
'            - UI スタンザ適用 (modUILoader.ApplyUiToSheet、Q20)
'            - タブ色 (ADR-0053 §2.1.1)
'            - format_dir 検査 (もし .txt 無ければ空フォルダ seed)
'            - Workbook.Protect Structure + シート保護
'            - ActiveSheet = ナレッジ検索 (M-08) (Q44)
' ================================================================
Private Sub Workbook_Open()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1772] ThisWorkbook.Workbook_Open ENTER"  ' [ADR-0100]
    ' [ADR-0100][gDebugLevel] init from config with safe fallback
    On Error Resume Next
    modCommon.gDebugLevel = modConfigHolder.GetDebugLevel()
    If Err.Number <> 0 Or modCommon.gDebugLevel < 0 Or modCommon.gDebugLevel > 5 Then
        modCommon.gDebugLevel = DEFAULT_DEBUG_LEVEL_FALLBACK
    End If
    Err.Clear
    On Error GoTo 0
    Debug.Print "[D-INIT] " & XLSM_NAME & " gDebugLevel=" & modCommon.gDebugLevel & " ts=" & Format$(Now, "hh:nn:ss")
    Debug.Print "[WBOPEN-ENTER] " & XLSM_NAME & " ts=" & Format$(Now, "hh:nn:ss")  ' [WBOPEN-DEBUG-PRINT-INJECTED]
    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' v2.1 Q34: 起動時に 90 日超バックアップ自動削除 (管理.xlsm のみ実施)
    On Error Resume Next
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1775] ThisWorkbook.Workbook_Open STEP-1 pre modKnowledgeFileIO.CleanupOldBackups"  ' [ADR-0100]
    Call modKnowledgeFileIO.CleanupOldBackups
    On Error GoTo ErrHandler

    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1776] ThisWorkbook.Workbook_Open STEP-2 pre orch.RunFullSetup"  ' [ADR-0100]
    Call orch.RunFullSetup(XLSM_NAME)
    ' 2026-06-06: delayed re-render (skip install COM)
    On Error Resume Next
    If Application.UserControl Then
        Application.OnTime Now + TimeValue("0:00:01"), "modRefresh.Btn_RefreshAllSheets"
    End If
    Err.Clear
    On Error GoTo ErrHandler

    ' S5-LOG-02: SAVE-EXIT-OK-II-003 (Workbook_Open success exit, screen 検索)
    On Error Resume Next
    Dim oLogger003 As clsLogger
    Set oLogger003 = New clsLogger
    oLogger003.Init ThisWorkbook.Worksheets("LOG")
    oLogger003.LogInfo "ThisWorkbook_kensaku", "Workbook_Open", "Workbook_Open 正常完了: " & XLSM_NAME, "", "SAVE-EXIT-OK-II-003"
    On Error GoTo 0

    ' Q44: startup ActiveSheet = STARTUP_SHEET (production spec mirror)
    On Error Resume Next
    ThisWorkbook.Worksheets(STARTUP_SHEET).Activate
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Debug.Print "[WBOPEN-EXIT] " & XLSM_NAME & " ts=" & Format$(Now, "hh:nn:ss")  ' [WBOPEN-DEBUG-PRINT-INJECTED]
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1773] ThisWorkbook.Workbook_Open EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1774] ThisWorkbook.Workbook_Open EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[WBOPEN-ERR] " & XLSM_NAME & " err=" & Err.Number & " " & Err.Description  ' [WBOPEN-DEBUG-PRINT-INJECTED]
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Dim msg As String
    msg = "検索.xlsm 起動エラー: " & Err.Description & vbCrLf & _
          "config.txt の存在 / debugLevel 値 / シート構造 / format_dir 書き込み権限 を確認してください"
    Debug.Print msg
    ' S5-LOG: BACKTOMAIN-ERR-EE-031 (Workbook_Open failure)
    On Error Resume Next
    Dim oLogger031 As clsLogger
    Set oLogger031 = New clsLogger
    oLogger031.Init ThisWorkbook.Worksheets("LOG")
    oLogger031.LogError "ThisWorkbook_kensaku", "Workbook_Open", "Workbook_Open 失敗: " & Err.Description, "", "BACKTOMAIN-ERR-EE-031"
    On Error GoTo 0
    ' v2.3 install_search.bat / headless 自動テストで Setup_search を
    ' 直接 MsgBox 出してハングさせる事への警戒対策。
    ' modCommon.IsHeadless() が COM 自動実行を検出し、その場合は
    ' MsgBox を出さず clsLogger / Debug.Print のみで通知する。
    If Not modCommon.IsHeadless() Then
        MsgBox msg, vbCritical, ChrW(&H8D77) & ChrW(&H52D5) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
    Else
        Debug.Print "[HEADLESS] suppressed MsgBox: " & msg
        modCommon.AppendProgressLog modCommon.ProgressTs() & "ThisWorkbook(検索).Workbook_Open ErrHandler suppressed MsgBox: " & Err.Description
    End If
End Sub

' ================================================================
' Workbook_SheetBeforeDoubleClick (v2.3 Phase O-1, 2026-05-27)
' M-08 ナレッジ検索一覧の knowledgeNo セル (column B) をダブルクリック
' すると modEntryUserForm.OpenViewWithId で M-09 ナレッジ表示
' フォームを開く。
' 仕様: 設計書 v2.3 Accepted "M-08 ナレッジ検索_数値" B53 (dblClickKnowledgeNo)。
' 対象 sheet: 表示名 "ナレッジ検索" もしくは ID "M-08"。
' 対象列: 列 A～G、データ行 (14 行以降。13 行以下は header)。
' ================================================================
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1777] ThisWorkbook.Workbook_SheetBeforeDoubleClick ENTER"  ' [ADR-0100]
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
    kid = Trim(CStr(Sh.Cells(Target.Row, 2).Value))
    If Len(kid) = 0 Then Exit Sub

    Cancel = True   ' セル編集モード抑止
    modEntryUserForm.OpenViewWithId kid
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1778] ThisWorkbook.Workbook_SheetBeforeDoubleClick EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1779] ThisWorkbook.Workbook_SheetBeforeDoubleClick EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] Workbook_SheetBeforeDoubleClick: " & Err.Number & " " & Err.Description
End Sub

' ================================================================
' Workbook_BeforeClose
' ================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1780] ThisWorkbook.Workbook_BeforeClose ENTER"  ' [ADR-0100]
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
        ws.Cells(r, 5).Value = "xlsm 終了: " & XLSM_NAME
    End If
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1781] ThisWorkbook.Workbook_BeforeClose EXIT-OK"  ' [ADR-0100]
End Sub
```
