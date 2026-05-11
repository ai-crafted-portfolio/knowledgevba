---
title: clsFormatManager.cls
---

# clsFormatManager.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | フォーマット一覧シート管理 / フィールド DSL 解決 |
| 行数 | 464 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsFormatManager` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormatManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' v12 注記: FL_COL_* / FL_START_ROW は modFormatColumns.bas へ移動 (ADR 0027)
'           VBA 仕様で .cls 内の Public Const が禁止のため標準モジュールに切出し。

' ================================================================
' Phase 6 レビュー: 17 subs/funcs に 9 error handlers。
' 状態管理の core class、各 Public method に ErrHandler ある程度完備。
' Private helpers (FindFormatRow 等) は呼出側で wrapped。指摘なし。
'
' D-2 注記: フォーマット一覧シート (SHEET_FORMAT_LIST) の列番号定数 FL_COL_*
' および FL_START_ROW は Public Const として外部公開する。
' modDemoSeeder 等の Cowork デモコードからは本クラスの定数を参照すること。
' (旧: 各モジュールで重複定義 → DRY 違反)
' ================================================================

' ================================================================
' クラス: clsFormatManager（ビジネスロジック層）
' 概要:   フォーマット定義の管理
'         フォーマット一覧・設計・プレビュー・採番・
'         バージョン管理を担当
' 依存先: clsLogger, modFileIO, modStringUtil, modCommon
' ================================================================

' --- フォーマット一覧シートの列番号 ---

' --- フォーマット設計シートのセル位置 ---
Private Const FD_ROW_FMT_ID As Long = 1
Private Const FD_COL_FMT_ID_VAL As Long = 2
Private Const FD_COL_VER_VAL As Long = 4
Private Const FD_COL_PATTERN_VAL As Long = 6
Private Const FD_ROW_FMT_NAME As Long = 2
Private Const FD_COL_FMT_NAME_VAL As Long = 2

' --- フィールドテーブルの位置 ---
Private Const FD_FIELDS_HEADER_ROW As Long = 5
Private Const FD_FIELDS_START_ROW As Long = 6
Private Const FD_FIELD_COL_NO As Long = 1
Private Const FD_FIELD_COL_NAME As Long = 2
Private Const FD_FIELD_COL_TYPE As Long = 3
Private Const FD_FIELD_COL_REQUIRED As Long = 4
Private Const FD_FIELD_COL_ROWS As Long = 5
Private Const FD_FIELD_COL_WIDTH As Long = 6
Private Const FD_FIELD_COL_HEIGHT As Long = 7
Private Const FD_FIELD_COL_REPEAT As Long = 8
Private Const FD_FIELD_COL_ROW_OP As Long = 9
Private Const FD_FIELD_COL_FORMAT As Long = 10
Private Const FD_FIELD_COL_DESC As Long = 11

Private m_logger As clsLogger

' ================================================================
' 関数名: Init
' 概要:   初期化
' 引数:   logger - ログ出力用
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger)
    Set m_logger = logger
End Sub

' ================================================================
' 関数名: GetSelectedFormatId
' 概要:   フォーマット一覧シートで選択されている行のフォーマットIDを返す
' 引数:   なし
' 戻り値: String - フォーマットID（選択がない場合は空文字列）
' ================================================================
Public Function GetSelectedFormatId() As String
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim selRow As Long
    selRow = ws.Application.Selection.Row
    
    If selRow < FL_START_ROW Then
        GetSelectedFormatId = ""
        Exit Function
    End If
    
    GetSelectedFormatId = CStr(ws.Cells(selRow, FL_COL_FMT_ID).Value)
    Exit Function

ErrHandler:
    GetSelectedFormatId = ""
End Function

' ================================================================
' 関数名: BeginCreate
' 概要:   フォーマット設計シートを空の新規作成状態にする
' 引数:   なし
' 戻り値: なし
' ================================================================
Public Sub BeginCreate()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    Call ClearDesignSheet(ws)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFormatManager", "BeginCreate", _
                          "新規フォーマット作成モード"
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFormatManager", "BeginCreate", Err.Description
    End If
End Sub

' ================================================================
' 関数名: BeginEdit
' 概要:   指定フォーマットIDをフォーマット設計シートに読み込んで編集モードにする
' 引数:   formatId - 編集対象のフォーマットID
' 戻り値: なし
' ================================================================
Public Sub BeginEdit(ByVal formatId As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    Call ClearDesignSheet(ws)
    Call LoadFormatIntoDesignSheet(ws, formatId)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFormatManager", "BeginEdit", _
                          "フォーマット編集モード: " & formatId
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFormatManager", "BeginEdit", Err.Description
    End If
End Sub

' ================================================================
' 関数名: SaveFormat
' 概要:   フォーマット設計シートの内容を検証してフォーマット一覧に保存
'         既存フォーマットの場合はバージョンを+1
' 引数:   なし
' 戻り値: Boolean - 保存成功なら True
' ================================================================
Public Function SaveFormat() As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    Dim formatId As String
    formatId = CStr(ws.Cells(FD_ROW_FMT_ID, FD_COL_FMT_ID_VAL).Value)
    
    If formatId = "" Then
        If Not m_logger Is Nothing Then
            m_logger.LogWarn "clsFormatManager", "SaveFormat", _
                              "フォーマットID未入力"
        End If
        SaveFormat = False
        Exit Function
    End If
    
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim existingRow As Long
    existingRow = FindFormatRow(listWs, formatId)
    
    If existingRow = 0 Then
        Call AppendFormatToList(listWs, ws, formatId)
    Else
        Call UpdateFormatInList(listWs, ws, formatId, existingRow)
    End If
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFormatManager", "SaveFormat", _
                          "保存完了: " & formatId
    End If
    
    SaveFormat = True
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFormatManager", "SaveFormat", Err.Description
    End If
    SaveFormat = False
End Function

' ================================================================
' 関数名: ShowPreview
' 概要:   指定フォーマットをプレビューシートに表示
' 引数:   formatId - 表示対象のフォーマットID
' 戻り値: なし
' ================================================================
Public Sub ShowPreview(ByVal formatId As String)
    On Error GoTo ErrHandler
    
    Dim previewWs As Worksheet
    Set previewWs = ThisWorkbook.Worksheets(SHEET_FORMAT_PREVIEW)
    
    Call ClearPreviewSheet(previewWs)
    Call RenderFormatPreview(previewWs, formatId)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFormatManager", "ShowPreview", _
                          "プレビュー表示: " & formatId
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFormatManager", "ShowPreview", Err.Description
    End If
End Sub

' ================================================================
' 関数名: IncrementAndGetNextNumber
' 概要:   指定フォーマットの採番値を+1して返す（保存時のユニーク保証）
' 引数:   formatId - フォーマットID
' 戻り値: Long - 採番値（失敗時は0）
' ================================================================
Public Function IncrementAndGetNextNumber(ByVal formatId As String) As Long
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim rowNum As Long
    rowNum = FindFormatRow(ws, formatId)
    
    If rowNum = 0 Then
        IncrementAndGetNextNumber = 0
        Exit Function
    End If
    
    Dim currentNum As Long
    currentNum = CLng(ws.Cells(rowNum, FL_COL_CURRENT_NUM).Value)
    
    Dim newNum As Long
    newNum = currentNum + 1
    
    ws.Cells(rowNum, FL_COL_CURRENT_NUM).Value = newNum
    ws.Cells(rowNum, FL_COL_NEXT_NUM).Value = newNum + 1
    
    IncrementAndGetNextNumber = newNum
    Exit Function

ErrHandler:
    IncrementAndGetNextNumber = 0
End Function

' ================================================================
' 関数名: GetFormatHeaderAsStanza
' 概要:   フォーマット定義をスタンザ形式文字列として取得
'         ナレッジファイル先頭に埋め込むための形式
' 引数:   formatId - フォーマットID
' 戻り値: String - スタンザ形式の文字列（失敗時は空文字列）
' ================================================================
Public Function GetFormatHeaderAsStanza(ByVal formatId As String) As String
    On Error GoTo ErrHandler
    
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim rowNum As Long
    rowNum = FindFormatRow(listWs, formatId)
    
    If rowNum = 0 Then
        GetFormatHeaderAsStanza = ""
        Exit Function
    End If
    
    Dim result As String
    result = "[FORMAT]" & vbCrLf
    result = result & "FormatID=" & formatId & vbCrLf
    result = result & "FormatName=" & CStr(listWs.Cells(rowNum, FL_COL_FMT_NAME).Value) & vbCrLf
    result = result & "Version=" & CStr(listWs.Cells(rowNum, FL_COL_VERSION).Value) & vbCrLf
    result = result & "IDPattern=" & CStr(listWs.Cells(rowNum, FL_COL_ID_PATTERN).Value) & vbCrLf
    
    GetFormatHeaderAsStanza = result
    Exit Function

ErrHandler:
    GetFormatHeaderAsStanza = ""
End Function

' ================================================================
' 関数名: FindFormatRow
' 概要:   フォーマット一覧から指定IDの行番号を返す
' 引数:   listWs   - フォーマット一覧シート
'         formatId - 探すフォーマットID
' 戻り値: Long - 行番号（見つからない場合は 0）
' ================================================================
Private Function FindFormatRow(ByVal listWs As Worksheet, _
                                 ByVal formatId As String) As Long
    Dim i As Long
    Dim maxScan As Long
    maxScan = 1000
    
    For i = FL_START_ROW To maxScan
        Dim id As String
        id = CStr(listWs.Cells(i, FL_COL_FMT_ID).Value)
        If id = "" Then
            Exit For
        End If
        If id = formatId Then
            FindFormatRow = i
            Exit Function
        End If
    Next i
    
    FindFormatRow = 0
End Function

' 新規フォーマットをフォーマット一覧に追記
Private Sub AppendFormatToList(ByVal listWs As Worksheet, _
                                 ByVal designWs As Worksheet, _
                                 ByVal formatId As String)
    Dim nextRow As Long
    nextRow = FindFirstEmptyRow(listWs, FL_COL_FMT_ID)
    
    listWs.Cells(nextRow, FL_COL_NO).Value = nextRow - FL_START_ROW + 1
    listWs.Cells(nextRow, FL_COL_FMT_ID).Value = formatId
    listWs.Cells(nextRow, FL_COL_FMT_NAME).Value = _
        designWs.Cells(FD_ROW_FMT_NAME, FD_COL_FMT_NAME_VAL).Value
    listWs.Cells(nextRow, FL_COL_ID_PATTERN).Value = _
        designWs.Cells(FD_ROW_FMT_ID, FD_COL_PATTERN_VAL).Value
    listWs.Cells(nextRow, FL_COL_CURRENT_NUM).Value = 0
    listWs.Cells(nextRow, FL_COL_NEXT_NUM).Value = 1
    listWs.Cells(nextRow, FL_COL_VERSION).Value = 1
    listWs.Cells(nextRow, FL_COL_FIELD_COUNT).Value = CountDesignFields(designWs)
    listWs.Cells(nextRow, FL_COL_KNW_COUNT).Value = 0
    listWs.Cells(nextRow, FL_COL_CREATED).Value = Format(Date, "yyyy-mm-dd")
    listWs.Cells(nextRow, FL_COL_UPDATED).Value = Format(Date, "yyyy-mm-dd")
End Sub

' 既存フォーマットを更新（バージョン+1）
Private Sub UpdateFormatInList(ByVal listWs As Worksheet, _
                                 ByVal designWs As Worksheet, _
                                 ByVal formatId As String, _
                                 ByVal rowNum As Long)
    Dim currentVer As Long
    currentVer = CLng(listWs.Cells(rowNum, FL_COL_VERSION).Value)
    
    listWs.Cells(rowNum, FL_COL_FMT_NAME).Value = _
        designWs.Cells(FD_ROW_FMT_NAME, FD_COL_FMT_NAME_VAL).Value
    listWs.Cells(rowNum, FL_COL_ID_PATTERN).Value = _
        designWs.Cells(FD_ROW_FMT_ID, FD_COL_PATTERN_VAL).Value
    listWs.Cells(rowNum, FL_COL_VERSION).Value = currentVer + 1
    listWs.Cells(rowNum, FL_COL_FIELD_COUNT).Value = CountDesignFields(designWs)
    listWs.Cells(rowNum, FL_COL_UPDATED).Value = Format(Date, "yyyy-mm-dd")
End Sub

' 指定列で最初の空セル行を返す
Private Function FindFirstEmptyRow(ByVal ws As Worksheet, _
                                     ByVal targetCol As Long) As Long
    Dim i As Long
    For i = FL_START_ROW To 10000
        If ws.Cells(i, targetCol).Value = "" Then
            FindFirstEmptyRow = i
            Exit Function
        End If
    Next i
    FindFirstEmptyRow = FL_START_ROW
End Function

' フォーマット設計シートのフィールド数を数える
Private Function CountDesignFields(ByVal ws As Worksheet) As Long
    Dim i As Long
    Dim count As Long
    count = 0
    For i = FD_FIELDS_START_ROW To 1000
        If ws.Cells(i, FD_FIELD_COL_NAME).Value = "" Then
            Exit For
        End If
        count = count + 1
    Next i
    CountDesignFields = count
End Function

' 設計シートをクリア（ヘッダー行は残す）
Private Sub ClearDesignSheet(ByVal ws As Worksheet)
    ws.Cells(FD_ROW_FMT_ID, FD_COL_FMT_ID_VAL).Value = ""
    ws.Cells(FD_ROW_FMT_ID, FD_COL_VER_VAL).Value = ""
    ws.Cells(FD_ROW_FMT_ID, FD_COL_PATTERN_VAL).Value = ""
    ws.Cells(FD_ROW_FMT_NAME, FD_COL_FMT_NAME_VAL).Value = ""
    
    Dim i As Long
    For i = FD_FIELDS_START_ROW To 1000
        If ws.Cells(i, FD_FIELD_COL_NAME).Value = "" Then Exit For
        ws.Range(ws.Cells(i, 1), ws.Cells(i, FD_FIELD_COL_DESC)).ClearContents
    Next i
End Sub

' フォーマット一覧の値を設計シートに反映
Private Sub LoadFormatIntoDesignSheet(ByVal designWs As Worksheet, _
                                        ByVal formatId As String)
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim rowNum As Long
    rowNum = FindFormatRow(listWs, formatId)
    If rowNum = 0 Then Exit Sub
    
    designWs.Cells(FD_ROW_FMT_ID, FD_COL_FMT_ID_VAL).Value = formatId
    designWs.Cells(FD_ROW_FMT_ID, FD_COL_VER_VAL).Value = _
        listWs.Cells(rowNum, FL_COL_VERSION).Value
    designWs.Cells(FD_ROW_FMT_ID, FD_COL_PATTERN_VAL).Value = _
        listWs.Cells(rowNum, FL_COL_ID_PATTERN).Value
    designWs.Cells(FD_ROW_FMT_NAME, FD_COL_FMT_NAME_VAL).Value = _
        listWs.Cells(rowNum, FL_COL_FMT_NAME).Value
End Sub

' プレビューシートをクリア
Private Sub ClearPreviewSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Cells.ClearContents
    On Error GoTo 0
End Sub

' フォーマットをプレビューシートに描画
Private Sub RenderFormatPreview(ByVal ws As Worksheet, _
                                  ByVal formatId As String)
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim rowNum As Long
    rowNum = FindFormatRow(listWs, formatId)
    If rowNum = 0 Then Exit Sub
    
    ws.Cells(1, 1).Value = "フォーマット"
    ws.Cells(1, 2).Value = formatId & " - " & _
                             CStr(listWs.Cells(rowNum, FL_COL_FMT_NAME).Value)
    
    Dim designWs As Worksheet
    Set designWs = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    Dim i As Long
    Dim previewRow As Long
    previewRow = 3
    
    For i = FD_FIELDS_START_ROW To 1000
        If designWs.Cells(i, FD_FIELD_COL_NAME).Value = "" Then Exit For
        ws.Cells(previewRow, 1).Value = _
            designWs.Cells(i, FD_FIELD_COL_NAME).Value
        ws.Cells(previewRow, 2).Value = _
            "(" & designWs.Cells(i, FD_FIELD_COL_TYPE).Value & " 入力エリア)"
        previewRow = previewRow + 1
    Next i
End Sub
```
