---
title: clsLogger.cls
description: clsLogger.cls のソースコード（コピペ用）
---

# clsLogger.cls

**配置先**: `共通モジュール (3 ブック全て)` 用の VBA モジュール  
**種類**: クラス モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\`
- ファイル名: `clsLogger.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

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

' ----------------------------------------------------------------
' 初期化
' ----------------------------------------------------------------

Public Sub Init(ByVal logSheet As Worksheet, Optional ByVal debugLevelOverride As Long = -99)
    On Error GoTo ErrHandler
    Set m_logSheet = logSheet
    If debugLevelOverride = -99 Then
        m_debugLevel = modConfigHolder.GetDebugLevel()  ' Q7/Q17 既定 ERROR
    Else
        m_debugLevel = debugLevelOverride
    End If
    m_nextRow = 0
    m_rotationRows = CLng(modConfigHolder.GetValueOrDefault("logRotationRows", "10000"))

    ' Q2: LOG sheet 既定 Visible = xlSheetHidden（VeryHidden ではない、ユーザ右クリック再表示可）
    If m_logSheet.Visible <> xlSheetVeryHidden Then  ' R-2-Fix2b: LOG VeryHidden (Visible/Hidden both)
        m_logSheet.Visible = xlSheetVeryHidden
    End If
    Exit Sub
ErrHandler:
    Debug.Print "[clsLogger.Init ERROR] " & Err.Description
End Sub

Private Sub Class_Initialize()
    m_debugLevel = DEBUG_OFF  ' 既定 OFF（Init 前は何も出さない）
    m_nextRow = 0
    m_rotationRows = -1
End Sub

' ----------------------------------------------------------------
' Public API（v1 互換、内部実装のみ v2.1 化）
' ----------------------------------------------------------------

Public Sub LogError(ByVal moduleName As String, ByVal funcName As String, ByVal msg As String, _
                    Optional ByVal target As String = "", Optional ByVal logId As String = "")
    If m_debugLevel < DEBUG_ERROR Then Exit Sub
    WriteLine "ERROR", moduleName, funcName, msg, target, logId
End Sub

Public Sub LogWarn(ByVal moduleName As String, ByVal funcName As String, ByVal msg As String, _
                   Optional ByVal target As String = "", Optional ByVal logId As String = "")
    If m_debugLevel < DEBUG_WARN Then Exit Sub
    WriteLine "WARN", moduleName, funcName, msg, target, logId
End Sub

Public Sub LogInfo(ByVal moduleName As String, ByVal funcName As String, _
                   Optional ByVal message As String = "", _
                   Optional ByVal target As String = "", Optional ByVal logId As String = "")
    If m_debugLevel < DEBUG_INFO Then Exit Sub
    WriteLine "INFO", moduleName, funcName, message, target, logId
End Sub

Public Sub LogDebug(ByVal moduleName As String, ByVal funcName As String, ByVal msg As String, _
                    Optional ByVal target As String = "", Optional ByVal logId As String = "")
    If m_debugLevel < DEBUG_DEBUG Then Exit Sub
    WriteLine "DEBUG", moduleName, funcName, msg, target, logId
End Sub

Public Sub LogTrace(ByVal moduleName As String, ByVal funcName As String, ByVal msg As String, _
                    Optional ByVal target As String = "", Optional ByVal logId As String = "")
    If m_debugLevel < DEBUG_TRACE Then Exit Sub
    WriteLine "TRACE", moduleName, funcName, msg, target, logId
End Sub

Public Sub LogEntry(ByVal moduleName As String, ByVal funcName As String, Optional ByVal logId As String = "")
    If m_debugLevel < DEBUG_TRACE Then Exit Sub
    WriteLine "TRACE", moduleName, funcName, "[ENTRY]", "", logId
End Sub

Public Sub LogExit(ByVal moduleName As String, ByVal funcName As String, Optional ByVal logId As String = "")
    If m_debugLevel < DEBUG_TRACE Then Exit Sub
    WriteLine "TRACE", moduleName, funcName, "[EXIT]", "", logId
End Sub

Public Sub LogBranch(ByVal moduleName As String, ByVal funcName As String, ByVal branchInfo As String, Optional ByVal logId As String = "")
    If m_debugLevel < DEBUG_TRACE Then Exit Sub
    WriteLine "TRACE", moduleName, funcName, "[BRANCH] " & branchInfo, "", logId
End Sub

Public Sub LogLoopProgress(ByVal moduleName As String, ByVal funcName As String, ByVal current As Long, ByVal total As Long, Optional ByVal logId As String = "")
    If m_debugLevel < DEBUG_TRACE Then Exit Sub
    WriteLine "TRACE", moduleName, funcName, "[LOOP " & current & "/" & total & "]", "", logId
End Sub

Public Sub LogErr(ByVal moduleName As String, ByVal funcName As String, ByRef errObj As ErrObject, _
                  Optional ByVal target As String = "", Optional ByVal logId As String = "")
    WriteLine "ERROR", moduleName, funcName, _
              "Err#" & errObj.Number & " " & errObj.Description, target, logId
End Sub

' Q17: GetDebugLevel As Long、enum 値返却（modConfigHolder 経由で 5 秒 polling 想定）
Public Function GetDebugLevel() As Long
    GetDebugLevel = m_debugLevel
End Function

Public Function IsTraceEnabled() As Boolean
    IsTraceEnabled = (m_debugLevel >= DEBUG_TRACE)
End Function

' debugLevel 動的更新（5 秒 polling 用、Application.OnTime で起動）
Public Sub RefreshDebugLevel()
    On Error GoTo ErrHandler
    Dim newLevel As Long
    newLevel = modConfigHolder.GetDebugLevel()
    If newLevel <> m_debugLevel Then
        WriteLine "INFO", "clsLogger", "RefreshDebugLevel", _
                  "debugLevel 変更: " & m_debugLevel & " -> " & newLevel, "", "LOG-LV-006"
        m_debugLevel = newLevel
    End If
    Exit Sub
ErrHandler:
    Debug.Print "[clsLogger.RefreshDebugLevel ERROR] " & Err.Description
End Sub

Public Sub ClearLog()
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
    Exit Sub
ErrHandler:
    Debug.Print "[clsLogger.ClearLog ERROR] " & Err.Description
End Sub

' ----------------------------------------------------------------
' Private: シート書き込み + FIFO ローテーション
' ----------------------------------------------------------------

Private Sub WriteLine(ByVal level As String, ByVal moduleName As String, ByVal funcName As String, _
                       ByVal msg As String, ByVal target As String, ByVal logId As String)
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
    Exit Sub

ErrHandler:
    ' Q7 規約 X：error handler 内で LogError 呼ぶ → 自己再帰で無限ループ防ぐため Debug.Print のみ
    Debug.Print "[clsLogger.WriteLine ERROR] " & Err.Description
End Sub

Private Function NextLogRow() As Long
    If m_nextRow > 0 Then
        NextLogRow = m_nextRow
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
```
