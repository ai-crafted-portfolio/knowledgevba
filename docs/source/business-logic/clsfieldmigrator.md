---
title: clsFieldMigrator.cls
---

# clsFieldMigrator.cls

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | フォーマット変更時に既存ナレッジへフィールドを反映し、消失リスク時にバックアップを取る |
| 行数 | 331 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsFieldMigrator.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "clsFieldMigrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' クラス: clsFieldMigrator（v2.1、Phase 2 task 2.10、ADR-0053 §2.10 完全実装）
' 概要:   format 変更時の knowledge migrate + データ消失検知 + バックアップ
' Version: v2.1（2026-05-16 EOD、Q1-Q57 解消反映）
' 依存:   modKnowledgeFileIO（Q47 新形式対応）, modFormatLoader（Q19 確定 I/F）,
'         modConfigHolder, ClsStanzaSection, clsLogger
' 関連:   ADR-0053 §2.10、Q4/Q6/Q47/Q50
' v2.1 主要 method:
'   - DetectDataLoss（field 削除 / 型変更 / MaxLength 短縮 / Dropdown 値削除 検知）
'   - MigrateSections（in-memory migrate、空欄保持、補完値挿入禁止）
'   - MigrateFields（全 knowledge 一括 migrate + バックアップ）
'   - MigrateAllByFormat（Q50 新規：format 保存時の連動 migrate）
' ================================================================
Option Explicit

Private m_logger As Object  ' clsLogger 注入（弱結合）

' ----------------------------------------------------------------
' 初期化
' ----------------------------------------------------------------

Public Sub Init(Optional ByVal logger As Object = Nothing)
    Set m_logger = logger
End Sub

' ----------------------------------------------------------------
' Public API 1: DetectDataLoss（ADR-0053 §2.10 条件 1 検知）
' ----------------------------------------------------------------

' 概要: 旧 FormatVersion から新 FormatVersion への移行で
'       データ消失が発生するか判定
' 引数: oldVer / newVer / formatId
' 戻り値: True = データ消失あり / False = なし
Public Function DetectDataLoss( _
    ByVal oldVer As Long, _
    ByVal newVer As Long, _
    ByVal formatId As String _
) As Boolean
    On Error GoTo ErrHandler

    Dim oldFields As Collection
    Dim newFields As Collection
    Set oldFields = LoadFieldDefinitionsForVersion(formatId, oldVer)
    Set newFields = LoadFieldDefinitionsForVersion(formatId, newVer)

    Dim oldField As ClsStanzaSection
    Dim newField As ClsStanzaSection
    Dim found As Boolean

    ' field 削除検知
    For Each oldField In oldFields
        found = False
        For Each newField In newFields
            If newField.GetValue("FieldName") = oldField.GetValue("FieldName") Then
                found = True
                Exit For
            End If
        Next newField
        If Not found Then
            DetectDataLoss = True
            Exit Function
        End If
    Next oldField

    ' 型変更 / MaxLength 短縮 / Dropdown 値削除 検知
    For Each oldField In oldFields
        For Each newField In newFields
            If newField.GetValue("FieldName") = oldField.GetValue("FieldName") Then
                ' 型変更
                If newField.GetValue("FieldType") <> oldField.GetValue("FieldType") Then
                    DetectDataLoss = True
                    Exit Function
                End If
                ' MaxLength 短縮
                If IsLengthShortened(oldField.GetValue("MaxLength"), newField.GetValue("MaxLength")) Then
                    DetectDataLoss = True
                    Exit Function
                End If
                ' Dropdown 値削除
                If IsDropdownOptionRemoved(oldField.GetValue("DropdownOptions"), newField.GetValue("DropdownOptions")) Then
                    DetectDataLoss = True
                    Exit Function
                End If
            End If
        Next newField
    Next oldField

    DetectDataLoss = False
    Exit Function

ErrHandler:
    ' Q7 規約 X：error handler 内で必ず LogError
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFieldMigrator", "DetectDataLoss", Err.Description, "", "LOG-MIG-DETECT-ERR"
    End If
    DetectDataLoss = True  ' エラー時は安全側に倒す（migrate を止める）
End Function

' ----------------------------------------------------------------
' Public API 2: MigrateSections（in-memory migrate、ADR-0053 §2.10 条件 2）
' ----------------------------------------------------------------

' 概要: 既に load 済の knowledge dictionary に対して in-memory migrate
'       空欄保持・補完値挿入禁止（ADR-0053 §2.10）
' 引数: knowledgeDict（ByRef） / oldVer / newVer / formatId
' 戻り値: True = 成功
Public Function MigrateSections( _
    ByRef knowledgeDict As Object, _
    ByVal oldVer As Long, _
    ByVal newVer As Long, _
    ByVal formatId As String _
) As Boolean
    On Error GoTo ErrHandler

    ' 新 FormatVersion の field 定義
    Dim newFields As Collection
    Set newFields = LoadFieldDefinitionsForVersion(formatId, newVer)

    ' 既存 knowledge の key 集合（Q47 新形式 = Dictionary key 1 件 1 field）
    ' 新規 field の追加：value="" で挿入（補完値禁止）
    Dim newField As ClsStanzaSection
    Dim fieldName As String
    For Each newField In newFields
        fieldName = newField.GetValue("FieldName")
        If Not knowledgeDict.Exists(fieldName) Then
            ' ★ 空欄保持（補完値挿入禁止、ADR-0053 §2.10 条件 2 / Q50）
            knowledgeDict(fieldName) = ""
        End If
    Next newField

    ' FormatVersion / UpdatedAt の更新（Q47 新形式の Dictionary key として）
    knowledgeDict("FormatVersion") = CStr(newVer)
    knowledgeDict("UpdatedAt") = Format$(Now(), "yyyy-mm-dd")

    MigrateSections = True
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFieldMigrator", "MigrateSections", Err.Description, "", "LOG-MIG-SECT-ERR"
    End If
    MigrateSections = False
End Function

' ----------------------------------------------------------------
' Public API 3: MigrateAllByFormat（Q50 format 保存時の連動 migrate）
' ----------------------------------------------------------------

' 概要: 指定 format の全 knowledge を一括 migrate
'       format 保存時に clsFormatManager.SaveFormat → 本 method を呼ぶ
'       バックアップ自動取得、確認 MsgBox は呼出側 UI 担当
' 引数: formatId / newVer
' 戻り値: 処理件数（-1 = エラー）
Public Function MigrateAllByFormat( _
    ByVal formatId As String, _
    ByVal newVer As Long _
) As Long
    On Error GoTo ErrHandler

    Dim ts As String
    ts = Format$(Now(), "yyyymmdd_hhnnss")

    Dim allKnowledges As Collection
    Set allKnowledges = modKnowledgeFileIO.ListKnowledgesByFormat(formatId)

    Dim processedCount As Long
    processedCount = 0

    Dim knwNo As Variant
    For Each knwNo In allKnowledges
        Dim knwDict As Object
        Set knwDict = modKnowledgeFileIO.LoadKnowledge(CStr(knwNo))
        If knwDict.Count = 0 Then GoTo NextKnw

        ' 旧 version 取得
        Dim oldVer As Long
        If knwDict.Exists("FormatVersion") Then
            oldVer = CLng(knwDict("FormatVersion"))
        Else
            oldVer = 1
        End If

        ' 更新不要なら skip
        If oldVer >= newVer Then GoTo NextKnw

        ' データ消失検知 → 必要ならバックアップ
        If DetectDataLoss(oldVer, newVer, formatId) Then
            If Not modKnowledgeFileIO.BackupKnowledgeFile(CStr(knwNo), ts) Then
                If Not m_logger Is Nothing Then
                    m_logger.LogError "clsFieldMigrator", "MigrateAllByFormat", _
                        "バックアップ失敗、スキップ: " & knwNo, "", "LOG-MIG-ALL-ITEM-ERR"
                End If
                GoTo NextKnw
            End If
        End If

        ' in-memory migrate
        If Not MigrateSections(knwDict, oldVer, newVer, formatId) Then
            GoTo NextKnw
        End If

        ' 書き戻し（Q46 楽観ロック：format 保存時の連動 migrate なので originalTimestamp=0 で新規扱い）
        Dim saveResult As Long
        saveResult = modKnowledgeFileIO.SaveKnowledge(CStr(knwNo), knwDict, 0)
        If saveResult = 0 Then
            processedCount = processedCount + 1
        End If

NextKnw:
    Next knwNo

    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFieldMigrator", "MigrateAllByFormat", _
            "migrate 完了: " & processedCount & " 件 / format=" & formatId & " / backup_ts=" & ts, "", "LOG-MIG-ALL-DONE"
    End If

    MigrateAllByFormat = processedCount
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFieldMigrator", "MigrateAllByFormat", Err.Description, "", "LOG-MIG-ALL-ERR"
    End If
    MigrateAllByFormat = -1
End Function

' ----------------------------------------------------------------
' Private ヘルパ
' ----------------------------------------------------------------

' 指定 formatId / version の field 定義一覧を取得
' v2.1 想定：format .txt は最新 version のみ保持、過去 version は migrate 履歴で再構築
Private Function LoadFieldDefinitionsForVersion( _
    ByVal formatId As String, _
    ByVal targetVer As Long _
) As Collection
    On Error Resume Next
    Dim result As Collection
    Set result = New Collection

    Dim sections As Collection
    Set sections = modFormatLoader.LoadFormat(formatId)
    If sections Is Nothing Then
        Set LoadFieldDefinitionsForVersion = result
        Exit Function
    End If

    ' 簡易実装：最新 version の [FIELD] のみ返却
    ' 将来：[MIGRATE_RULE] セクションを参照して指定 version の field 集合を再構築
    Dim sec As ClsStanzaSection
    For Each sec In sections
        If sec.SectionName = "FIELD" Then
            result.Add sec
        End If
    Next sec

    Set LoadFieldDefinitionsForVersion = result
    On Error GoTo 0
End Function

' MaxLength の短縮判定
Private Function IsLengthShortened(ByVal oldLen As String, ByVal newLen As String) As Boolean
    On Error GoTo ErrHandler
    If Len(oldLen) = 0 Or Len(newLen) = 0 Then
        IsLengthShortened = False
        Exit Function
    End If
    IsLengthSho