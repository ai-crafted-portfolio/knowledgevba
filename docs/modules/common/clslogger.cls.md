---
title: clsLogger.cls
description: clsLogger.cls のソースコード（コピペ用）
---

# clsLogger.cls

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: クラスモジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsLogger.cls`
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
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' クラス: clsLogger（v2.1、Phase 2 task 2.1～2.2、ログシート専用）
' 概要:   debugLevel 6 値フィルタ + FIFO ローテーション + シート専用出力
' Version: v2.1（2026-05-16 EOD、Q1-Q57 解消反映）
' 関連:   ADR-0053 §2.4 / §5.2 / §7.3 N6
'         Q7（debugLevel 既定 ERROR + VBA 規約 3 件追加）
'         Q17（GetDebugLevel As Long、enum 値返却）
'         Q2（LOG sheet xlSheetHidden 既定）
' v2.1 変更点（v1 からの差分）:
'   - 外部 file 出力 (C:\kvba\runtime.log) 完全削除
'   - debugLevel 6 値（OFF=0 ～ TRACE=5、modConfigHolder の Public Const 参照）
'   - debugLevel 読込元 = modConfigHolder.GetDebugLevel()（NM-2 解消）
'   - FIFO ローテーション（既定 10000 行、modConfigHolder.GetValueOrDefault("logRotationRows", "10000")）
'   - LOG sheet 既定 Visible = xlSheetHidden（VeryHidden ではない、Q2）
' ================================================================
Option Explicit

' v2.1 debugLevel enum 6 値（modConfigHolder にも同等の定数あり）
Private Const DEBUG_OFF As Long = 0
Private Const DEBUG_ERROR As Long = 1
Private Const DEBUG_WARN As Long = 2
Private Const DEBUG_INFO As Long = 3
Private Const DEBUG_DEBUG As Long = 4
Private Const DEBUG_TRACE As Long = 5

Private m_logSheet As Worksheet
Private m_debugLevel As Long
Private m_nextRow As Long
Private m_rotationRows As Long  ' FIFO 上限行数（-1 = 無効）

' v2.3 task AD-LOG (2026-06-03): Init-not-yet buffer + Debug.Print mirror
Private m_buffer As Collection      ' pending entries before Init succeeds
Private m_initialized As Boolean    ' True after a successful Init()

' ----------------------------------------------------------------
' 初期化
' ----------------------------------------------------------------

Public Sub Init(ByVal logSheet As Worksheet, Optional ByVal debugLevelOverride As Long = -99)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0548] clsLogger.Init ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Set m_logSheet = logSheet
    If debugLevelOverride = -99 Then
        ' [ADR-0100][gDebugLevel-sync] prefer Workbook_Open populated
        ' modCommon.gDebugLevel; fall back to direct config read for
        ' COM headless paths that bypass Workbook_Open.
        If modCommon.gDebugLevel > 0 Then
            m_debugLevel = modCommon.gDebugLevel
        Else
            m_debugLevel = modConfigHolder.GetDebugLevel()
        End If
        ' iter20 fix: headless E2E COM mode (PS EnableEvents=false)
        ' bypasses Workbook_Open so the holder may not be initialized
        ' yet -- bump from ERROR default to INFO so LOG-effect test
        ' assertions can observe entry-point markers.
        On Error Resume Next
        If modCommon.IsHeadless() And m_debugLevel < DEBUG_INFO Then
            m_debugLevel = DEBUG_INFO
        End If
        On Error GoTo ErrHandler
        ' Q7/Q17 default fallback marker:  ' Q7/Q17 既定 ERROR
    Else
        m_debugLevel = debugLevelOverride
    End If
    m_nextRow = 0
    m_rotationRows = CLng(modConfigHolder.GetValueOrDefault("logRotationRows", "10000"))

    ' Q2: LOG sheet 既定 Visible = xlSheetHidden（VeryHidden ではない、ユーザ右クリック再表示可）
    ' [N8 2026-06-12] do NOT force-hide on every Init: setup hides LOG initially;
    ' the user may unhide it deliberately and Logger kept re-hiding it.
    ' v2.3 AD-LOG: mark ready and flush any pre-Init buffered entries
    m_initialized = True
    FlushBuffer
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0549] clsLogger.Init EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0550] clsLogger.Init EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[clsLogger.Init ERROR] " & Err.Description
End Sub

Private Sub Class_Initialize()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0551] clsLogger.Class_Initialize ENTER"  ' [ADR-0100]
    m_debugLevel = DEBUG_OFF
    m_nextRow = 0
    m_rotationRows = -1
    Set m_buffer = New Collection  ' AD-LOG: stash pre-Init entries
    m_initialized = False
End Sub

' ----------------------------------------------------------------
' Public API（v1 互換、内部実装のみ v2.1 化）
' ----------------------------------------------------------------

Public Sub LogError(ByVal moduleName As String, ByVal funcName As String, ByVal msg As String, _
                    Optional ByVal target As String = "", Optional ByVal logId As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0552] clsLogger.LogError ENTER"  ' [ADR-0100]
    ' v2.3 AD-LOG: ERROR is ALWAYS recorded (no debugLevel gate)
    DebugPrintMirror "ERROR", moduleName, funcName, msg, target, logId
    Dispatch "ERROR", moduleName, funcName, msg, target, logId
End Sub

Public Sub LogWarn(ByVal moduleName As String, ByVal funcName As String, ByVal msg As String, _
                   Optional ByVal target As String = "", Optional ByVal logId As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0553] clsLogger.LogWarn ENTER"  ' [ADR-0100]
    ' v2.3 AD-LOG: WARN is ALWAYS recorded (no debugLevel gate)
    DebugPrintMirror "WARN", moduleName, funcName, msg, target, logId
    Dispatch "WARN", moduleName, funcName, msg, target, logId
End Sub

Public Sub LogInfo(ByVal moduleName As String, ByVal funcName As String, _
                   Optional ByVal message As String = "", _
                   Optional ByVal target As String = "", Optional ByVal logId As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0554] clsLogger.LogInfo ENTER"  ' [ADR-0100]
    ' v2.3 AD-LOG: INFO honors debugLevel; Debug.Print mirror always on
    DebugPrintMirror "INFO", moduleName, funcName, message, target, logId
    If m_debugLevel < DEBUG_INFO Then Exit Sub
    Dispatch "INFO", moduleName, funcName, message, target, logId
End Sub

Public Sub LogDebug(ByVal moduleName As String, ByVal funcName As String, ByVal msg As String, _
                    Optional ByVal target As String = "", Optional ByVal logId As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0555] clsLogger.LogDebug ENTER"  ' [ADR-0100]
    ' v2.3 AD-LOG: DEBUG honors debugLevel; Debug.Print mirror always on
    DebugPrintMirror "DEBUG", moduleName, funcName, msg, target, logId
    If m_debugLevel < DEBUG_DEBUG Then Exit Sub
    Dispatch "DEBUG", moduleName, funcName, msg, target, logId
End Sub

Public Sub LogTrace(ByVal moduleName As String, ByVal funcName As String, ByVal msg As String, _
                    Optional ByVal target As String = "", Optional ByVal logId As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0556] clsLogger.LogTrace ENTER"  ' [ADR-0100]
    DebugPrintMirror "TRACE", moduleName, funcName, msg, target, logId
    If m_debugLevel < DEBUG_TRACE Then Exit Sub
    Dispatch "TRACE", moduleName, funcName, msg, target, logId
End Sub

Public Sub LogEntry(ByVal moduleName As String, ByVal funcName As String, Optional ByVal logId As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0557] clsLogger.LogEntry ENTER"  ' [ADR-0100]
    DebugPrintMirror "TRACE", moduleName, funcName, "[ENTRY]", "", logId
    If m_debugLevel < DEBUG_TRACE Then Exit Sub
    Dispatch "TRACE", moduleName, funcName, "[ENTRY]", "", logId
End Sub

Public Sub LogExit(ByVal moduleName As String, ByVal funcName As String, Optional ByVal logId As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0558] clsLogger.LogExit ENTER"  ' [ADR-0100]
    DebugPrintMirror "TRACE", moduleName, funcName, "[EXIT]", "", logId
    If m_debugLevel < DEBUG_TRACE Then Exit Sub
    Dispatch "TRACE", moduleName, funcName, "[EXIT]", "", logId
End Sub

Public Sub LogBranch(ByVal moduleName As String, ByVal funcName As String, ByVal branchInfo As String, Optional ByVal logId As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0559] clsLogger.LogBranch ENTER"  ' [ADR-0100]
    DebugPrintMirror "TRACE", moduleName, funcName, "[BRANCH] " & branchInfo, "", logId
    If m_debugLevel < DEBUG_TRACE Then Exit Sub
    Dispatch "TRACE", moduleName, funcName, "[BRANCH] " & branchInfo, "", logId
End Sub

Public Sub LogLoopProgress(ByVal moduleName As String, ByVal funcName As String, ByVal current As Long, ByVal total As Long, Optional ByVal logId As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0560] clsLogger.LogLoopProgress ENTER"  ' [ADR-0100]
    DebugPrintMirror "TRACE", moduleName, funcName, "[LOOP " & current & "/" & total & "]", "", logId
    If m_debugLevel < DEBUG_TRACE Then Exit Sub
    Dispatch "TRACE", moduleName, funcName, "[LOOP " & current & "/" & total & "]", "", logId
End Sub

Public Sub LogErr(ByVal moduleName As String, ByVal funcName As String, ByRef errObj As ErrObject, _
                  Optional ByVal target As String = "", Optional ByVal logId As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0561] clsLogger.LogErr ENTER"  ' [ADR-0100]
    Dim composed As String
    composed = "Err#" & errObj.Number & " " & errObj.Description
    DebugPrintMirror "ERROR", moduleName, funcName, composed, target, logId
    Dispatch "ERROR", moduleName, funcName, composed, target, logId
End Sub

' Q17: GetDebugLevel As Long、enum 値返却（modConfigHolder 経由で 5 秒 polling 想定）
Public Function GetDebugLevel() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0562] clsLogger.GetDebugLevel ENTER"  ' [ADR-0100]
    GetDebugLevel = m_debugLevel
End Function

Public Function IsTraceEnabled() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0563] clsLogger.IsTraceEnabled ENTER"  ' [ADR-0100]
    IsTraceEnabled = (m_debugLevel >= DEBUG_TRACE)
End Function

' debugLevel 動的更新（5 秒 polling 用、Application.OnTime で起動）
Public Sub RefreshDebugLevel()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0564] clsLogger.RefreshDebugLevel ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim newLevel As Long
    newLevel = modConfigHolder.GetDebugLevel()
    If newLevel <> m_debugLevel Then
        WriteLine "INFO", "clsLogger", "RefreshDebugLevel", _
                  "debugLevel " & ChrW(&H5909) & ChrW(&H66F4) & ": " & m_debugLevel & " -> " & newLevel, "", "LOG-LV-006"
        m_debugLevel = newLevel
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0565] clsLogger.RefreshDebugLevel EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0566] clsLogger.RefreshDebugLevel EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[clsLogger.RefreshDebugLevel ERROR] " & Err.Description
End Sub

Public Sub ClearLog()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0567] clsLogger.ClearLog ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_logSheet Is Nothing Then Exit Sub
    Dim dataStart As Long
    dataStart = 9  ' LOG_DATA_START_ROW（modCommon 経由）
    Dim lastRow As Long
    lastRow = m_logSheet.Cells(m_logSheet.Rows.Count, 1).End(xlUp).Row
    If lastRow >= dataStart Then
        m_logSheet.Rows(dataStart & ":" & lastRow).Clear
    End If
    m_nextRow = 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0568] clsLogger.ClearLog EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0569] clsLogger.ClearLog EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[clsLogger.ClearLog ERROR] " & Err.Description
End Sub

' ----------------------------------------------------------------
' Private: シート書き込み + FIFO ローテーション
' ----------------------------------------------------------------

Private Sub WriteLine(ByVal level As String, ByVal moduleName As String, ByVal funcName As String, _
                       ByVal msg As String, ByVal target As String, ByVal logId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0570] clsLogger.WriteLine ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_logSheet Is Nothing Then Exit Sub

    Dim r As Long
    r = NextLogRow()

    ' FIFO ローテーション（rotation 上限超過時、先頭行を削除）
    If m_rotationRows > 0 Then
        Dim dataStart As Long
        dataStart = 9  ' LOG_DATA_START_ROW
        If r - dataStart >= m_rotationRows Then
            m_logSheet.Rows(dataStart).Delete
            r = r - 1
        End If
    End If

    m_logSheet.Cells(r, 1).Value = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
    m_logSheet.Cells(r, 2).Value = moduleName
    m_logSheet.Cells(r, 3).Value = funcName
    m_logSheet.Cells(r, 4).Value = level
    m_logSheet.Cells(r, 5).Value = msg
    m_logSheet.Cells(r, 6).Value = Environ("USERNAME")  ' Q1 actor 自動記録
    m_logSheet.Cells(r, 7).Value = target
    If Len(logId) > 0 Then
        m_logSheet.Cells(r, 8).Value = logId
    End If

    m_nextRow = r + 1
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0571] clsLogger.WriteLine EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0572] clsLogger.WriteLine EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' Q7 規約 X：error handler 内で LogError 呼ぶ → 自己再帰で無限ループ防ぐため Debug.Print のみ
    Debug.Print "[clsLogger.WriteLine ERROR] " & Err.Description
End Sub

Private Function NextLogRow() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0573] clsLogger.NextLogRow ENTER"  ' [ADR-0100]
    If m_nextRow > 0 Then
        NextLogRow = m_nextRow
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0574] clsLogger.NextLogRow EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim dataStart As Long
    dataStart = 9  ' LOG_DATA_START_ROW
    Dim lastRow As Long
    lastRow = m_logSheet.Cells(m_logSheet.Rows.Count, 1).End(xlUp).Row
    If lastRow < dataStart Then
        NextLogRow = dataStart
    Else
        NextLogRow = lastRow + 1
    End If
End Function

' ----------------------------------------------------------------
' v2.3 AD-LOG helpers: Debug.Print mirror, Dispatch, buffer flush
' ----------------------------------------------------------------

Private Sub DebugPrintMirror(ByVal level As String, ByVal moduleName As String, _
                              ByVal funcName As String, ByVal msg As String, _
                              ByVal target As String, ByVal logId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0575] clsLogger.DebugPrintMirror ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim line As String
    line = "[" & level & "] " & moduleName & "." & funcName & " - " & msg
    If Len(target) > 0 Then line = line & " | target=" & target
    If Len(logId) > 0 Then line = line & " | logId=" & logId
    Debug.Print line
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0576] clsLogger.DebugPrintMirror EXIT-OK"  ' [ADR-0100]
End Sub

' Dispatch: write to sheet if Init succeeded; otherwise stash in buffer
Private Sub Dispatch(ByVal level As String, ByVal moduleName As String, _
                     ByVal funcName As String, ByVal msg As String, _
                     ByVal target As String, ByVal logId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0577] clsLogger.Dispatch ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_initialized And Not (m_logSheet Is Nothing) Then
        WriteLine level, moduleName, funcName, msg, target, logId
    Else
        EnqueueBuffer level, moduleName, funcName, msg, target, logId
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0578] clsLogger.Dispatch EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0579] clsLogger.Dispatch EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[clsLogger.Dispatch ERROR] " & Err.Description
End Sub

Private Sub EnqueueBuffer(ByVal level As String, ByVal moduleName As String, _
                          ByVal funcName As String, ByVal msg As String, _
                          ByVal target As String, ByVal logId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0580] clsLogger.EnqueueBuffer ENTER"  ' [ADR-0100]
    On Error Resume Next
    If m_buffer Is Nothing Then Set m_buffer = New Collection
    ' soft cap: prevent unbounded growth before Init completes
    If m_buffer.Count >= 500 Then Exit Sub
    Dim entry(0 To 6) As Variant
    entry(0) = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
    entry(1) = level
    entry(2) = moduleName
    entry(3) = funcName
    entry(4) = msg
    entry(5) = target
    entry(6) = logId
    m_buffer.Add entry
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0581] clsLogger.EnqueueBuffer EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub FlushBuffer()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0582] clsLogger.FlushBuffer ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_buffer Is Nothing Then Exit Sub
    If m_buffer.Count = 0 Then Exit Sub
    If m_logSheet Is Nothing Then Exit Sub
    Dim i As Long
    Dim entry As Variant
    For i = 1 To m_buffer.Count
        entry = m_buffer.Item(i)
        ' arrival order is preserved (Collection 1-based, FIFO)
        WriteLine CStr(entry(1)), CStr(entry(2)), CStr(entry(3)), _
                  CStr(entry(4)), CStr(entry(5)), CStr(entry(6))
    Next i
    Set m_buffer = New Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0583] clsLogger.FlushBuffer EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0584] clsLogger.FlushBuffer EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[clsLogger.FlushBuffer ERROR] " & Err.Description
End Sub
```
