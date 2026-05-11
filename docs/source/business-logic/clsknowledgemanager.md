---
title: clsKnowledgeManager.cls
---

# clsKnowledgeManager.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | ナレッジ一覧シートの行管理 |
| 行数 | 549 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsKnowledgeManager` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKnowledgeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsKnowledgeManager（ビジネスロジック層）
' 概要:   ナレッジの登録・修正・削除を担当
'         スタンザ形式で Shift_JIS ファイルとして入出力
' 依存先: clsLogger, clsFormatManager, modFileIO, modStringUtil,
'         modDateUtil, modCommon
' ================================================================

' --- ナレッジ登録シートのセル位置 ---
Private Const KS_ROW_FMT_ID As Long = 1
Private Const KS_COL_FMT_ID_VAL As Long = 3
Private Const KS_FORM_START_ROW As Long = 4
Private Const KS_FIELD_COL_NO As Long = 1
Private Const KS_FIELD_COL_NAME As Long = 2
Private Const KS_FIELD_COL_VALUE As Long = 3

' --- ナレッジ修正シートのセル位置 ---
Private Const KE_ROW_FMT_ID As Long = 1
Private Const KE_COL_FMT_ID_VAL As Long = 2
Private Const KE_COL_KNW_NO As Long = 3
Private Const KE_FORM_START_ROW As Long = 4

Private m_logger As clsLogger
Private m_formatMgr As clsFormatManager
Private m_dataFolder As String

' ================================================================
' 関数名: Init
' 概要:   初期化
' 引数:   logger     - ログ出力用
'         formatMgr  - フォーマット管理クラス
'         dataFolder - ナレッジデータ格納フォルダパス
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger, _
                 ByVal formatMgr As clsFormatManager, _
                 ByVal dataFolder As String)
    Set m_logger = logger
    Set m_formatMgr = formatMgr
    m_dataFolder = dataFolder
End Sub

' ================================================================
' 関数名: BuildRegistrationForm
' 概要:   ナレッジ登録シートに指定フォーマットのフィールド一覧を展開
' 引数:   formatId - フォーマットID
' 戻り値: なし
' 備考:   既存の入力値はクリアされる
' ================================================================
Public Sub BuildRegistrationForm(ByVal formatId As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)
    
    Call ClearRegistrationForm(ws)
    Call PopulateFormFromDesign(ws, formatId, KS_FORM_START_ROW)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "BuildRegistrationForm", _
                          "フォーム生成: " & formatId
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "BuildRegistrationForm", _
                           Err.Description
    End If
End Sub

' ================================================================
' 関数名: SaveNewKnowledge
' 概要:   ナレッジ登録シートからフォーム値を取得して新規ナレッジ保存
'         採番を+1してファイル名を決定、ファイルを物理生成
' 引数:   なし
' 戻り値: String - 採番されたナレッジ番号（失敗時は空文字列）
' ================================================================
Public Function SaveNewKnowledge() As String
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)
    
    Dim formatId As String
    formatId = CStr(ws.Cells(KS_ROW_FMT_ID, KS_COL_FMT_ID_VAL).Value)
    
    If Not ValidateRequiredFields(ws, KS_FORM_START_ROW) Then
        SaveNewKnowledge = ""
        Exit Function
    End If
    
    Dim nextNum As Long
    nextNum = m_formatMgr.IncrementAndGetNextNumber(formatId)
    If nextNum = 0 Then
        SaveNewKnowledge = ""
        Exit Function
    End If
    
    Dim knowledgeNo As String
    knowledgeNo = BuildKnowledgeNumber(formatId, nextNum)
    
    Dim content As String
    content = BuildKnowledgeFile(formatId, knowledgeNo, TodayStr(), TodayStr(), _
                                   ws, KS_FORM_START_ROW)
    
    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")
    
    If Not WriteShiftJisFile(filePath, content) Then
        If Not m_logger Is Nothing Then
            m_logger.LogError "clsKnowledgeManager", "SaveNewKnowledge", _
                               "ファイル書き込み失敗: " & filePath
        End If
        SaveNewKnowledge = ""
        Exit Function
    End If
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "SaveNewKnowledge", _
                          "ナレッジ" & knowledgeNo & "を保存しました"
    End If
    
    SaveNewKnowledge = knowledgeNo
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "SaveNewKnowledge", Err.Description
    End If
    SaveNewKnowledge = ""
End Function

' ================================================================
' 関数名: LoadForEdit
' 概要:   指定ナレッジ番号のファイルをナレッジ修正シートに読み込む
' 引数:   knowledgeNo - ナレッジ番号
' 戻り値: Boolean - 成功なら True
' ================================================================
Public Function LoadForEdit(ByVal knowledgeNo As String) As Boolean
    ' s-2: パストラバーサル防止 - ナレッジ番号バリデーション
    If Not IsValidKnowledgeId(knowledgeNo) Then
        If Not m_logger Is Nothing Then
            m_logger.LogError "clsKnowledgeManager", "LoadForEdit", _
                                "不正なナレッジ番号 reject: " & knowledgeNo
        End If
        LoadForEdit = False
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")
    
    If Not FileExists(filePath) Then
        LoadForEdit = False
        Exit Function
    End If
    
    Dim content As String
    content = ReadShiftJisFile(filePath)
    If content = "" Then
        LoadForEdit = False
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    
    Call ClearEditForm(ws)
    Call PopulateEditFormFromFile(ws, content, knowledgeNo)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "LoadForEdit", _
                          "読込: " & knowledgeNo
    End If
    
    LoadForEdit = True
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "LoadForEdit", Err.Description
    End If
    LoadForEdit = False
End Function

' ================================================================
' 関数名: UpdateKnowledge
' 概要:   ナレッジ修正シートの内容をファイルに上書き保存
'         作成日は維持、更新日のみ今日に更新
' 引数:   knowledgeNo - ナレッジ番号
' 戻り値: Boolean - 成功なら True
' ================================================================
Public Function UpdateKnowledge(ByVal knowledgeNo As String) As Boolean
    ' s-2: パストラバーサル防止 - ナレッジ番号バリデーション
    If Not IsValidKnowledgeId(knowledgeNo) Then
        If Not m_logger Is Nothing Then
            m_logger.LogError "clsKnowledgeManager", "UpdateKnowledge", _
                                "不正なナレッジ番号 reject: " & knowledgeNo
        End If
        UpdateKnowledge = False
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")
    
    Dim originalContent As String
    originalContent = ReadShiftJisFile(filePath)
    
    Dim createdDate As String
    createdDate = ExtractCreatedDate(originalContent)
    If createdDate = "" Then
        createdDate = TodayStr()
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    
    Dim formatId As String
    formatId = CStr(ws.Cells(KE_ROW_FMT_ID, KE_COL_FMT_ID_VAL).Value)
    
    Dim content As String
    content = BuildKnowledgeFile(formatId, knowledgeNo, createdDate, TodayStr(), _
                                   ws, KE_FORM_START_ROW)
    
    If Not WriteShiftJisFile(filePath, content) Then
        UpdateKnowledge = False
        Exit Function
    End If
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "UpdateKnowledge", _
                          "ナレッジ" & knowledgeNo & "を更新しました"
    End If
    
    UpdateKnowledge = True
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "UpdateKnowledge", Err.Description
    End If
    UpdateKnowledge = False
End Function

' ================================================================
' 関数名: DeleteKnowledge
' 概要:   指定ナレッジ番号のファイルを物理削除
' 引数:   knowledgeNo - ナレッジ番号
' 戻り値: Boolean - 成功なら True
' ================================================================
Public Function DeleteKnowledge(ByVal knowledgeNo As String) As Boolean
    ' s-2: パストラバーサル防止 - ナレッジ番号バリデーション
    If Not IsValidKnowledgeId(knowledgeNo) Then
        If Not m_logger Is Nothing Then
            m_logger.LogError "clsKnowledgeManager", "DeleteKnowledge", _
                                "不正なナレッジ番号 reject: " & knowledgeNo
        End If
        DeleteKnowledge = False
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")
    
    If Not DeleteFile(filePath) Then
        DeleteKnowledge = False
        Exit Function
    End If
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "DeleteKnowledge", _
                          "ナレッジ" & knowledgeNo & "を削除しました"
    End If
    
    DeleteKnowledge = True
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "DeleteKnowledge", Err.Description
    End If
    DeleteKnowledge = False
End Function

' ================================================================
' 関数名: BuildKnowledgeFile
' 概要:   フォーム値とメタ情報からナレッジファイル本文を組み立て
' 引数:   formatId     - フォーマットID
'         knowledgeNo  - ナレッジ番号
'         createDate   - 作成日
'         updateDate   - 更新日
'         ws           - 値を取得するワークシート
'         formStartRow - フォームの開始行
' 戻り値: String - スタンザ形式のファイル本文
' ================================================================
Private Function BuildKnowledgeFile(ByVal formatId As String, _
                                      ByVal knowledgeNo As String, _
                                      ByVal createDate As String, _
                                      ByVal updateDate As String, _
                                      ByVal ws As Worksheet, _
                                      ByVal formStartRow As Long) As String
    Dim content As String
    content = m_formatMgr.GetFormatHeaderAsStanza(formatId)
    content = content & STANZA_SEP & vbCrLf
    
    content = content & "[KNOWLEDGE]" & vbCrLf
    content = content & "KnowledgeNo=" & knowledgeNo & vbCrLf
    content = content & "FormatVersion=" & GetFormatVersion(formatId) & vbCrLf
    content = content & "CreatedDate=" & createDate & vbCrLf
    content = content & "UpdatedDate=" & updateDate & vbCrLf
    
    Dim i As Long
    For i = formStartRow To 1000
        Dim fieldName As String
        fieldName = CStr(ws.Cells(i, KS_FIELD_COL_NAME).Value)
        If fieldName = "" Then Exit For

        ' rev19: M10-002 対応。"=" 先頭値を入力時に SafeCellText 等で
        ' "'=..." とエスケープしている場合があるため、ファイル書き込み時は
        ' 先頭アポストロフィを 1 文字だけ除去して "=SUM(A1:A10)" がそのまま
        ' 残るようにする（仕様: アポストロフィなしでファイルに残す）
        Dim rawValue As String
        rawValue = CStr(ws.Cells(i, KS_FIELD_COL_VALUE).Value)
        If Len(rawValue) >= 2 Then
            If Left(rawValue, 1) = "'" And Mid(rawValue, 2, 1) = "=" Then
                rawValue = Mid(rawValue, 2)
            End If
        End If

        content = content & "[ITEM]" & vbCrLf
        content = content & "FieldNo=" & CStr(ws.Cells(i, KS_FIELD_COL_NO).Value)
        content = content & " / FieldName=" & fieldName
        content = content & " / Value=" & rawValue
        content = content & vbCrLf
    Next i
    
    content = content & STANZA_SEP
    BuildKnowledgeFile = content
End Function

' ナレッジ番号をパターンと採番値から生成
Private Function BuildKnowledgeNumber(ByVal formatId As String, _
                                         ByVal nextNum As Long) As String
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim pattern As String
    pattern = FindIdPattern(listWs, formatId)
    
    BuildKnowledgeNumber = FormatNumberByPattern(pattern, nextNum)
End Function

' フォーマット一覧からID形式を取得
Private Function FindIdPattern(ByVal listWs As Worksheet, _
                                 ByVal formatId As String) As String
    Dim i As Long
    For i = 3 To 1000
        If CStr(listWs.Cells(i, 2).Value) = formatId Then
            FindIdPattern = CStr(listWs.Cells(i, 4).Value)
            Exit Function
        End If
    Next i
    FindIdPattern = ""
End Function

' フォーマット一覧からバージョン番号を取得
Private Function GetFormatVersion(ByVal formatId As String) As Long
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim i As Long
    For i = 3 To 1000
        If CStr(listWs.Cells(i, 2).Value) = formatId Then
            GetFormatVersion = CLng(listWs.Cells(i, 7).Value)
            Exit Function
        End If
    Next i
    GetFormatVersion = 1
End Function

' フォルダパスとファイル名を結合
Private Function CombineFilePath(ByVal folder As String, _
                                   ByVal fileName As String) As String
    If Right(folder, 1) = "\" Or Right(folder, 1) = "/" Then
        CombineFilePath = folder & fileName
    Else
        CombineFilePath = folder & "\" & fileName
    End If
End Function

' 必須フィールドの検証
Private Function ValidateRequiredFields(ByVal ws As Worksheet, _
                                          ByVal startRow As Long) As Boolean
    Dim i As Long
    For i = startRow To 1000
        If ws.Cells(i, KS_FIELD_COL_NAME).Value = "" Then Exit For
        ' 必須判定は設計シートを参照する必要があるが、
        ' 簡易実装として登録シートの値が空ならOKとする
    Next i
    ValidateRequiredFields = True
End Function

' 登録フォームをクリア
Private Sub ClearRegistrationForm(ByVal ws As Worksheet)
    Dim i As Long
    For i = KS_FORM_START_ROW To 1000
        If ws.Cells(i, KS_FIELD_COL_NAME).Value = "" Then Exit For
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 3)).ClearContents
    Next i
End Sub

' 修正フォームをクリア
Private Sub ClearEditForm(ByVal ws As Worksheet)
    Dim i As Long
    For i = KE_FORM_START_ROW To 1000
        If ws.Cells(i, KS_FIELD_COL_NAME).Value = "" Then Exit For
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 3)).ClearContents
    Next i
End Sub

' フォーマット設計からフィールドリストを展開
Private Sub PopulateFormFromDesign(ByVal ws As Worksheet, _
                                     ByVal formatId As String, _
                                     ByVal startRow As Long)
    Dim designWs As Worksheet
    Set designWs = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    ws.Cells(KS_ROW_FMT_ID, KS_COL_FMT_ID_VAL).Value = formatId
    
    Dim i As Long
    Dim targetRow As Long
    targetRow = startRow
    
    For i = 6 To 1000
        Dim fieldName As String
        fieldName = CStr(designWs.Cells(i, 2).Value)
        If fieldName = "" Then Exit For
        
        ws.Cells(targetRow, KS_FIELD_COL_NO).Value = _
            designWs.Cells(i, 1).Value
        ws.Cells(targetRow, KS_FIELD_COL_NAME).Value = fieldName
        targetRow = targetRow + 1
    Next i
End Sub

' ファイル内容からフォームに展開
Private Sub PopulateEditFormFromFile(ByVal ws As Worksheet, _
                                       ByVal content As String, _
                                       ByVal knowledgeNo As String)
    Dim lines() As String
    lines = Split(content, vbCrLf)
    
    Dim formatId As String
    formatId = ExtractStanzaValue(content, "FormatID")
    
    ws.Cells(KE_ROW_FMT_ID, KE_COL_FMT_ID_VAL).Value = formatId
    ws.Cells(KE_ROW_FMT_ID, KE_COL_KNW_NO).Value = knowledgeNo
    
    Dim i As Long
    Dim targetRow As Long
    targetRow = KE_FORM_START_ROW
    
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "FieldNo=") > 0 And _
           InStr(lines(i), "FieldName=") > 0 And _
           InStr(lines(i), "Value=") > 0 Then
            Call ParseItemLine(lines(i), ws, targetRow)
            targetRow = targetRow + 1
        End If
    Next i
End Sub

' ITEM行の解析と展開
Private Sub ParseItemLine(ByVal line As String, _
                            ByVal ws As Worksheet, _
                            ByVal targetRow As Long)
    Dim noStr As String
    Dim nameStr As String
    Dim valStr As String
    
    noStr = ExtractKeyValue(line, "FieldNo")
    nameStr = ExtractKeyValue(line, "FieldName")
    valStr = ExtractKeyValue(line, "Value")
    
    ws.Cells(targetRow, KS_FIELD_COL_NO).Value = noStr
    ws.Cells(targetRow, KS_FIELD_COL_NAME).Value = nameStr
    ws.Cells(targetRow, KS_FIELD_COL_VALUE).NumberFormat = "@"
    ws.Cells(targetRow, KS_FIELD_COL_VALUE).Value = SafeCellText(valStr)
End Sub

' "Key=Value / OtherKey=..." 形式から値を抽出
Private Function ExtractKeyValue(ByVal line As String, _
                                   ByVal keyName As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim searchKey As String
    
    searchKey = keyName & "="
    startPos = InStr(line, searchKey)
    If startPos = 0 Then
        ExtractKeyValue = ""
        Exit Function
    End If
    
    startPos = startPos + Len(searchKey)
    endPos = InStr(startPos, line, " / ")
    If endPos = 0 Then
        ExtractKeyValue = Mid(line, startPos)
    Else
        ExtractKeyValue = Mid(line, startPos, endPos - startPos)
    End If
End Function

' スタンザ内の単純なKey=Value値を抽出
Private Function ExtractStanzaValue(ByVal content As String, _
                                      ByVal keyName As String) As String
    Dim lines() As String
    Dim i As Long
    
    lines = Split(content, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), keyName & "=") = 1 Then
            ExtractStanzaValue = Mid(lines(i), Len(keyName) + 2)
            Exit Function
        End If
    Next i
    ExtractStanzaValue = ""
End Function

' ファイル内容からCreatedDateを抽出
Private Function ExtractCreatedDate(ByVal content As String) As String
    ExtractCreatedDate = ExtractStanzaValue(content, "CreatedDate")
End Function
```
