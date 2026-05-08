---
title: clsLogger.cls
---

# clsLogger.cls

| 項目 | 値 |
|---|---|
| 層 | ログ層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | ログエントリの蓄積 / シート出力 / TS タイムスタンプ |
| 行数 | 204 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsLogger` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsLogger（ビジネスロジック層）
' 概要:   ログシートへのログ出力を担当
'         デバッグレベル（OFF/ON）に応じて出力内容を制御
' 依存先: clsLogEntry, modDateUtil, modStringUtil, modCommon
' ================================================================

Private m_logSheet As Worksheet
Private m_debugLevel As String
Private m_nextRow As Long  ' M-3: O(N^2) 防止のためキャッシュ (0=未初期化)

' m-12: マジックナンバー 100000 を Const 化
Private Const MAX_LOG_SCAN_ROWS As Long = 100000

' --- Property Get ---

Public Property Get DebugLevel() As String
    DebugLevel = m_debugLevel
End Property

' ================================================================
' 関数名: Init
' 概要:   初期化（ログシート参照とデバッグレベルを保持）
' 引数:   logSheet    - ログ出力先のワークシート
'         debugLevel  - "OFF" または "ON"
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logSheet As Worksheet, ByVal debugLevel As String)
    Set m_logSheet = logSheet
    m_debugLevel = debugLevel
End Sub

' ================================================================
' 関数名: LogError
' 概要:   エラーログを出力（デバッグレベル OFF/ON に関わらず出力）
' 引数:   modName  - モジュール名
'         funcName - 関数名
'         message  - メッセージ内容
' 戻り値: なし
' ================================================================
Public Sub LogError(ByVal modName As String, _
                      ByVal funcName As String, _
                      ByVal message As String)
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_ERROR, message)
End Sub

' ================================================================
' 関数名: LogWarn
' 概要:   警告ログを出力（デバッグレベル OFF/ON に関わらず出力）
' 引数:   modName  - モジュール名
'         funcName - 関数名
'         message  - メッセージ内容
' 戻り値: なし
' ================================================================
Public Sub LogWarn(ByVal modName As String, _
                     ByVal funcName As String, _
                     ByVal message As String)
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_WARN, message)
End Sub

' ================================================================
' 関数名: LogInfo
' 概要:   情報ログを出力（デバッグレベル OFF/ON に関わらず出力）
' 引数:   modName  - モジュール名
'         funcName - 関数名
'         message  - メッセージ内容
' 戻り値: なし
' ================================================================
Public Sub LogInfo(ByVal modName As String, _
                     ByVal funcName As String, _
                     ByVal message As String)
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_INFO, message)
End Sub

' ================================================================
' 関数名: LogDebug
' 概要:   デバッグログを出力（デバッグレベル ON 時のみ出力）
' 引数:   modName  - モジュール名
'         funcName - 関数名
'         message  - メッセージ内容
' 戻り値: なし
' 備考:   OFF時は何もせず即座にリターン
' ================================================================
Public Sub LogDebug(ByVal modName As String, _
                      ByVal funcName As String, _
                      ByVal message As String)
    If m_debugLevel <> DEBUG_ON Then
        Exit Sub
    End If
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_DEBUG, message)
End Sub

' ================================================================
' 関数名: ClearLog
' 概要:   ログシートの2行目以降（データ行）を全削除
'         ヘッダー行は保持
' 引数:   なし
' 戻り値: なし
' ================================================================
Public Sub ClearLog()
    On Error GoTo ErrHandler
    
    Dim lastRow As Long
    lastRow = GetLastLogRow()
    
    If lastRow < 2 Then
        Exit Sub
    End If
    
    m_logSheet.Range(m_logSheet.Cells(2, 1), _
                      m_logSheet.Cells(lastRow, 5)).ClearContents
    m_nextRow = 2  ' M-3: cache reset to row 2 (after header)
    Exit Sub

ErrHandler:
    ' クリア失敗時は何もしない（致命的ではないため）
End Sub

' ================================================================
' 関数名: WriteLogInternal
' 概要:   ログエントリを作成してシートに書き込む内部関数
' 引数:   modName  - モジュール名
'         funcName - 関数名
'         level    - ログレベル
'         message  - メッセージ内容
' 戻り値: なし
' 備考:   「=」先頭メッセージは SafeCellText で安全化
'         NumberFormat = "@" で事前にテキスト書式に設定
' ================================================================
Private Sub WriteLogInternal(ByVal modName As String, _
                               ByVal funcName As String, _
                               ByVal level As String, _
                               ByVal message As String)
    On Error GoTo ErrHandler

    ' s-4: メッセージから改行コードを除去 (ログ injection 防止)
    message = Replace(message, vbCrLf, " | ")
    message = Replace(message, vbLf, " | ")
    message = Replace(message, vbCr, " | ")

    Dim entry As clsLogEntry
    Set entry = New clsLogEntry
    entry.Init NowStr(), modName, funcName, level, message
    
    Dim nextRow As Long
    nextRow = GetNextLogRow()
    
    ' テキスト書式に設定してから書き込み（"=" 先頭の数式化を防ぐ）
    m_logSheet.Range(m_logSheet.Cells(nextRow, 1), _
                      m_logSheet.Cells(nextRow, 5)).NumberFormat = "@"
    
    m_logSheet.Cells(nextRow, LOG_COL_TIMESTAMP).Value = entry.Timestamp
    m_logSheet.Cells(nextRow, LOG_COL_MODULE).Value = entry.ModuleName
    m_logSheet.Cells(nextRow, LOG_COL_FUNCTION).Value = entry.FunctionName
    m_logSheet.Cells(nextRow, LOG_COL_LEVEL).Value = entry.Level
    m_logSheet.Cells(nextRow, LOG_COL_MESSAGE).Value = SafeCellText(entry.Message)
    Exit Sub

ErrHandler:
    ' ログ書き込み失敗は握りつぶす（無限ループを防ぐため）
End Sub

' ================================================================
' 関数名: GetLastLogRow
' 概要:   ログシートの最終行を返す（Range.End 未使用、LibreOffice互換）
' 引数:   なし
' 戻り値: Long - 最終データ行の行番号（データなしの場合は1=ヘッダー行）
' 備考:   1行目から順に走査して、A列が空のセルの1つ手前を最終行とする
' ================================================================
Private Function GetLastLogRow() As Long
    Dim i As Long
    Dim maxScan As Long
    maxScan = MAX_LOG_SCAN_ROWS
    
    For i = 1 To maxScan
        If m_logSheet.Cells(i, 1).Value = "" Then
            GetLastLogRow = i - 1
            Exit Function
        End If
    Next i
    
    GetLastLogRow = maxScan
End Function

' ================================================================
' 関数名: GetNextLogRow
' 概要:   次に書き込むべきログ行の行番号を返す
' 引数:   なし
' 戻り値: Long - 次の書き込み行番号
' ================================================================
Private Function GetNextLogRow() As Long
    GetNextLogRow = GetLastLogRow() + 1
End Function
```

## 関連

- 呼び出す: `clsLogEntry`, `modDateUtil`
- 呼び出される: `全エントリポイント / ビジネスロジック`
