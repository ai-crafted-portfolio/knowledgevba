---
title: clsLogger.cls
---

# clsLogger.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | ログエントリの蓄積 / シート出力 / 外部ファイル出力 (TRACE 含む) |
| 行数 | 183 行 |

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
' 概要:   ログシート + 外部ファイル(C:\kvba\runtime.log) 二重出力。
'         レベル: TRACE / DEBUG / INFO / WARN / ERROR
'         step 識別子で「どこまで通って、どこで失敗したか」を残す。
' 依存先: clsLogEntry, modDateUtil, modStringUtil, modCommon
' 備考:   外部ファイル append は VBA の Open For Append で実装。
'         (coding-standards R12 は通常 I/O 用; ログ append は実用上の例外)
'         外部書込みが連続失敗した場合は m_externalDisabled で抑止。
' ================================================================

Private m_logSheet As Worksheet
Private m_debugLevel As String
Private m_nextRow As Long  ' M-3: O(N^2) 防止のためキャッシュ (0=未初期化)
Private m_externalPath As String
Private m_externalDisabled As Boolean

' m-12: マジックナンバー 100000 を Const 化
Private Const MAX_LOG_SCAN_ROWS As Long = 100000

' v20: M-14 リデザインに伴うデータ書込開始行。A1 はタイトル帯、A8 は表ヘッダ、
'      データ本体は A9 から。これにより画面描画 (タイトル/ヘッダ) と
'      ログ書込が衝突しない。
Private Const LOG_DATA_START_ROW As Long = 9

' --- Property Get ---

Public Property Get DebugLevel() As String
    DebugLevel = m_debugLevel
End Property

Public Property Get ExternalPath() As String
    ExternalPath = m_externalPath
End Property

' ================================================================
' 関数名: Init
' 概要:   初期化（ログシート参照とデバッグレベルを保持）
' 引数:   logSheet    - ログ出力先のワークシート
'         debugLevel  - "OFF" または "ON"
' 戻り値: なし
' 備考:   外部ログパスは EXTERNAL_LOG_PATH 定数を既定値とする。
'         呼び出し側で SetExternalPath を使えば差し替え可能。
' ================================================================
Public Sub Init(ByVal logSheet As Worksheet, ByVal debugLevel As String)
    Set m_logSheet = logSheet
    m_debugLevel = debugLevel
    m_externalPath = EXTERNAL_LOG_PATH
    m_externalDisabled = False
End Sub

' ================================================================
' 関数名: SetExternalPath
' 概要:   外部ログファイルのパスを差し替える（テスト等で利用）
' ================================================================
Public Sub SetExternalPath(ByVal newPath As String)
    m_externalPath = newPath
    m_externalDisabled = False
End Sub

' ================================================================
' 関数名: LogError
' 概要:   エラーログを出力（debugLevel 関係なく出力）
' ================================================================
Public Sub LogError(ByVal modName As String, _
                      ByVal funcName As String, _
                      ByVal message As String)
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_ERROR, message)
End Sub

' ================================================================
' 関数名: LogErrorWithErr
' 概要:   ErrHandler から呼ぶ用。step 名と Err 情報を組み立てて 1 行で記録。
' 引数:   stepName  - 失敗直前に通っていた step 名（呼び出し側で文字列管理）
'         errNum    - Err.Number
'         errDesc   - Err.Description
' ================================================================
Public Sub LogErrorWithErr(ByVal modName As String, _
                             ByVal funcName As String, _
                             ByVal stepName As String, _
                             ByVal errNum As Long, _
                             ByVal errDesc As String)
    Dim msg As String
    msg = "FAIL step=[" & stepName & "] errNum=" & errNum & " desc=" & errDesc
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_ERROR, msg)
End Sub

' ================================================================
' 関数名: LogWarn
' ================================================================
Public Sub LogWarn(ByVal modName As String, _
                     ByVal funcName As String, _
                     ByVal message As String)
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_WARN, message)
End Sub

' ================================================================
' 関数名: LogInfo
' ================================================================
Public Sub LogInfo(ByVal modName As String, _
                     ByVal funcName As String, _
                     ByVal message As String)
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_INFO, message)
End Sub

' ================================================================
' 関数名: LogTrace
' 概要:   ENTER/EXIT/step マーカー用のトレースログ。
'         シート書き込みは debugLevel=ON 時のみ（シート過多防止）。
'         外部ファイルには常に書き出して「ここまで通った」を残す。
' ================================================================
Public Sub LogTrace(ByVal modName As String, _
                      ByVal funcName As String, _
                      ByVal message As String)
    If m_debugLevel = DEBUG_ON Then
        Call WriteLogInternal(modName, funcName, LOG_LEVEL_TRACE, message)
    Else
        Call WriteToExternalFileSafe(modName, funcName, LOG_LEVEL_TRACE, message)
    End If
End Sub

' ================================================================
' 関数名: LogDebug
' 概要:   デバッグログを出力（debugLevel ON 時のみシートに出す）
'         外部ファイルには常に出力する（後追い解析用）。
' ================================================================
Public Sub LogDebug(ByVal modName As String, _
                      ByVal funcName As String, _
                      ByVal message As String)
    If m_debugLevel <> DEBUG_ON Then
        Call WriteToExternalFileSafe(modName, funcName, LOG_LEVEL_DEBUG, message)
        Exit Sub
    End If
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_DEBUG, message)
End Sub

' ================================================================
' 関数名: ClearLog
' 概要:   ログシートの A9 以降を全削除（ヘッダ A8 と A1 タイトルは保持）
' ================================================================
Public Sub ClearLog()
    On Error GoTo ErrHandler

    Dim lastRow As Long
    lastRow = GetLastLogRow()

    If lastRow < LOG_DATA_START_ROW Then
        ' データなし。外部ログには起動マーカー
        Call WriteToExternalFileSafe("clsLogger", "ClearLog", LOG_LEVEL_INFO, "ログクリア(空)")
        Exit Sub
    End If

    m_logSheet.Range(m_logSheet.Cells(LOG_DATA_START_ROW, 1), _
                      m_logSheet.Cells(lastRow, 5)).ClearContents
    m_nextRow = LOG_DATA_START_ROW  ' v20: cache reset to data start row
    Call WriteToExternalFileSafe("clsLogger", "ClearLog", LOG_LEVEL_INFO, _
                                  "ログクリア完了 (rows=" & (lastRow - LOG_DATA_START_ROW + 1) & ")")
    Exit Sub

ErrHandler:
    ' クリア失敗時は外部ログだけでも残す
    Call WriteToExternalFileSafe("clsLogger", "ClearLog", LOG_LEVEL_ERROR, _
                                  "ClearLog 失敗 errNum=" & Err.Number & " desc=" & Err.Description)
End Sub

' ================================================================
' 関数名: WriteLogInternal
' 概要:   シート + 外部ファイル両方への書き込みを統括する内部関数。
'         シート書込みが失敗しても外部ファイルへの試行は実施。
' ================================================================
Private Sub WriteLogInternal(ByVal modName As String, _
```
