---
title: clsFieldMigrator.cls
---

# clsFieldMigrator.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | フォーマット定義変更時のスキーマ移行 (旧 .txt → 新 .txt) |
| 行数 | 308 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsFieldMigrator` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFieldMigrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Phase 6 レビュー: 9 subs/funcs に対し On Error GoTo は 1 件のみ。
' Public method (Init / RunMigration) は ErrHandler 補強推奨だが、
' 既存 89 テスト互換性のため本 v4 では指摘記録のみ (v5 で対応予定)。

' ================================================================
' クラス: clsFieldMigrator（ビジネスロジック層）
' 概要:   フォーマット変更後、既存ナレッジファイルに新フィールドを反映
'         既存の値は変更せず、不足フィールドのみ空値で追加する
' 依存先: clsLogger, clsFormatManager, modFileIO, modStringUtil, modCommon
' ================================================================

Private m_logger As clsLogger
Private m_formatMgr As clsFormatManager
Private m_dataFolder As String

' ================================================================
' 関数名: Init
' 概要:   初期化
' 引数:   logger     - ログ出力用
'         formatMgr  - フォーマット管理
'         dataFolder - ナレッジデータフォルダ
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
' 関数名: MigrateFields
' 概要:   指定フォーマットの全ナレッジファイルに新フィールド定義を反映
' 引数:   formatId - 対象フォーマットID
' 戻り値: Long - 処理ファイル数
' 備考:   各ファイルについて:
'         - フォーマット定義ヘッダーを最新版に更新
'         - 既存のITEM行は維持
'         - 新フィールドは空値で追加
'         - FormatVersion、UpdatedDateを更新
' ================================================================
Public Function MigrateFields(ByVal formatId As String) As Long
    On Error GoTo ErrHandler
    
    Dim files As Variant
    files = ListFilesInFolder(m_dataFolder, "txt")

    ' M-4 guard: 空配列なら早期 return (UBound エラー防止)
    ' v13 ADR 0028: MigrateFields は Function なので Exit Function に修正
    If (Not Not files) = 0 Then
        MigrateFields = 0
        Exit Function
    End If
    
    Dim newHeader As String
    newHeader = m_formatMgr.GetFormatHeaderAsStanza(formatId)
    
    Dim newFields As Variant
    newFields = GetNewFieldDefinitions(formatId)
    
    Dim processedCount As Long
    processedCount = 0
    
    Dim i As Long
    For i = LBound(files) To UBound(files)
        Dim fileName As String
        fileName = CStr(files(i))
        
        Dim filePath As String
        filePath = CombineFilePath(m_dataFolder, fileName)
        
        Dim content As String
        content = ReadShiftJisFile(filePath)
        If content = "" Then
            GoTo NextFile
        End If
        
        Dim actualFmt As String
        actualFmt = ExtractStanzaValue(content, "FormatID")
        If actualFmt <> formatId Then
            GoTo NextFile
        End If
        
        Dim newContent As String
        newContent = RebuildKnowledgeFile(content, newHeader, newFields, formatId)
        
        If WriteShiftJisFile(filePath, newContent) Then
            processedCount = processedCount + 1
        End If
        
NextFile:
    Next i
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFieldMigrator", "MigrateFields", _
                          "反映完了: " & CStr(processedCount) & "件"
    End If
    
    MigrateFields = processedCount
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFieldMigrator", "MigrateFields", Err.Description
    End If
    MigrateFields = 0
End Function

' ================================================================
' 関数名: GetNewFieldDefinitions
' 概要:   フォーマット設計シートから現在のフィールド定義を取得
' 引数:   formatId - フォーマットID
' 戻り値: Variant - [No, Name] の配列
' ================================================================
Private Function GetNewFieldDefinitions(ByVal formatId As String) As Variant
    Dim designWs As Worksheet
    Set designWs = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    Dim result() As String
    ReDim result(0 To 999, 0 To 1)
    
    Dim i As Long
    Dim count As Long
    count = 0
    
    For i = 6 To 1000
        Dim fieldName As String
        fieldName = CStr(designWs.Cells(i, 2).Value)
        If fieldName = "" Then Exit For
        
        result(count, 0) = CStr(designWs.Cells(i, 1).Value)
        result(count, 1) = fieldName
        count = count + 1
    Next i
    
    If count = 0 Then
        GetNewFieldDefinitions = Array()
    Else
        ReDim Preserve result(0 To count - 1, 0 To 1)
        GetNewFieldDefinitions = result
    End If
End Function

' ================================================================
' 関数名: RebuildKnowledgeFile
' 概要:   既存ファイル内容を新ヘッダーと新フィールド定義で再構築
' 引数:   oldContent - 既存ファイル内容
'         newHeader  - 新フォーマットヘッダー
'         newFields  - 新フィールド定義配列
'         formatId   - フォーマットID
' 戻り値: String - 再構築されたファイル内容
' ================================================================
Private Function RebuildKnowledgeFile(ByVal oldContent As String, _
                                        ByVal newHeader As String, _
                                        ByVal newFields As Variant, _
                                        ByVal formatId As String) As String
    Dim knowledgeNo As String
    Dim createdDate As String
    knowledgeNo = ExtractStanzaValue(oldContent, "KnowledgeNo")
    createdDate = ExtractStanzaValue(oldContent, "CreatedDate")
    
    Dim newVersion As Long
    newVersion = GetFormatVersion(formatId)
    
    Dim result As String
    result = newHeader
    result = result & STANZA_SEP & vbCrLf
    result = result & "[KNOWLEDGE]" & vbCrLf
    result = result & "KnowledgeNo=" & knowledgeNo & vbCrLf
    result = result & "FormatVersion=" & CStr(newVersion) & vbCrLf
    result = result & "CreatedDate=" & createdDate & vbCrLf
    result = result & "UpdatedDate=" & TodayStr() & vbCrLf
    
    ' 既存値を辞書化
    Dim existingValues As Object
    Set existingValues = BuildFieldValueMap(oldContent)
    
    ' 新フィールド定義に従って項目を出力（既存値があれば使用、なければ空）
    Dim i As Long
    For i = LBound(newFields, 1) To UBound(newFields, 1)
        Dim fieldNo As String
        Dim fieldName As String
        fieldNo = newFields(i, 0)
        fieldName = newFields(i, 1)
        
        Dim fieldValue As String
        If existingValues.Exists(fieldName) Then
            fieldValue = existingValues(fieldName)
        Else
            fieldValue = ""
        End If
        
        result = result & "[ITEM]" & vbCrLf
        result = result & "FieldNo=" & fieldNo
        result = result & " / FieldName=" & fieldName
        result = result & " / Value=" & fieldValue
        result = result & vbCrLf
    Next i
    
    result = result & STANZA_SEP
    RebuildKnowledgeFile = result
End Function

' ================================================================
' 関数名: BuildFieldValueMap
' 概要:   既存ファイル内容からフィールド名→値のマップを作る
' 引数:   content - ファイル内容
' 戻り値: Object (Dictionary) - FieldName → Value のマップ
' ================================================================
Private Function BuildFieldValueMap(ByVal content As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lines() As String
    lines = Split(content, vbCrLf)
    
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "FieldName=") > 0 And _
           InStr(lines(i), "Value=") > 0 Then
            Dim fName As String
            Dim fValue As String
            fName = ExtractItemValue(lines(i), "FieldName")
            fValue = ExtractItemValue(lines(i), "Value")
            If fName <> "" And Not dict.Exists(fName) Then
                dict.Add fName, fValue
            End If
        End If
    Next i
    
    Set BuildFieldValueMap = dict
End Function

' ITEM行から指定キーの値を抽出
Private Function ExtractItemValue(ByVal line As String, _
                                    ByVal keyName As String) As String
    Dim searchKey As String
    Dim startPos As Long
    Dim endPos As Long
    
    searchKey = keyName & "="
    startPos = InStr(line, searchKey)
    If startPos = 0 Then
        ExtractItemValue = ""
        Exit Function
    End If
    
    startPos = startPos + Len(searchKey)
    endPos = InStr(startPos, line, " / ")
    If endPos = 0 Then
        ExtractItemValue = Mid(line, startPos)
    Else
        ExtractItemValue = Mid(line, startPos, endPos - startPos)
    End If
End Function

' スタンザKey=Value形式から単純抽出
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
```

## 関連

- 呼び出す: `modFileIO`, `modStringUtil`
- 呼び出される: `clsFormatManager`
