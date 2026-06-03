---
title: clsLogger.cls
---

# clsLogger.cls

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | 操作ログをログシートへ出力。詳細度フィルタと行数上限ローテーション |
| 行数 | 226 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsLogger.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
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
    If m_logSheet.Visible <> xlSheetHidden And m_logSheet.Visible <> xlSheetVisible Then
        m_logSheet.Visible = xlSheetHidden
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
  