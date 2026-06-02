---
title: clsFormatManager.cls
description: clsFormatManager.cls のソースコード（コピペ用）
---

# clsFormatManager.cls

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsFormatManager.cls`
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
Attribute VB_Name = "clsFormatManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' クラス: clsFormatManager（v2.1、Phase 4、format 業務ロジック）
' 概要:   format 定義 CRUD（管理.xlsm 限定）+ 保存時 migrate 連動
' Version: v2.1（2026-05-16 EOD、Q6/Q19/Q41/Q43/Q50/Q55 反映）
' 関連:   modFormatLoader（Q19、ThisWorkbook.Name enforce 内蔵）
'         clsFieldMigrator（Q50 MigrateAllByFormat）
'         modKnowledgeFileIO（Q55 ListKnowledgesByFormat）
' v2.1 主要更新:
'   - SHEET_FORMAT_LIST 廃止 → modFormatLoader 経由
'   - Q41 M-09 削除、Q43 M-07 削除 → M-08 で新規/編集/削除集約
'   - Q50 SaveFormat 時に clsFieldMigrator.MigrateAllByFormat 自動呼出
'   - Q55 DeleteFormat 前に紐づく knowledge 件数 check（modFormatLoader 内 enforce）
'   - Q6 FieldType 5 種前提
' ================================================================
Option Explicit

Private m_logger As Object
' v2.3 deprecated: clsFieldMigrator is excluded from the package.
' The type is changed to Variant; no instance is created.
Private m_migrator As Variant

' ----------------------------------------------------------------
' 初期化
' ----------------------------------------------------------------

Public Sub Init(Optional ByVal logger As Object = Nothing)
    Set m_logger = logger
    ' v2.3 deprecated: clsFieldMigrator is excluded from the package.
    ' Set m_migrator = New clsFieldMigrator
    ' m_migrator.Init logger
    m_migrator = Empty
End Sub

' ----------------------------------------------------------------
' Public API: 読込
' ----------------------------------------------------------------

' format 定義 1 件を取得
Public Function LoadFormat(ByVal formatId As String) As Collection
    Set LoadFormat = modFormatLoader.LoadFormat(formatId)
End Function

' format 一覧（M-02 ドロップダウン Q28 用）
Public Function LoadFormatList() As Collection
    Set LoadFormatList = modFormatLoader.LoadFormatList()
End Function

' format ID 一覧（拡張子なし）
Public Function ListAllFormats() As Collection
    Set ListAllFormats = modFormatLoader.ListAllFormats()
End Function

' format 存在確認
Public Function FormatExists(ByVal formatId As String) As Boolean
    FormatExists = modFormatLoader.FormatExists(formatId)
End Function

' ----------------------------------------------------------------
' Public API: 保存（Q50 自動 migrate 連動）
' ----------------------------------------------------------------

' 概要: format 定義を保存、既存 knowledge を Q50 で自動 migrate
' 引数: formatId / sections / newVer（新 FormatVersion）
' 戻り値: 0 = 成功 / 1 = 管理.xlsm 外 reject / 2 = その他エラー
'         migrate 件数は m_logger に LogInfo
Public Function SaveFormat( _
    ByVal formatId As String, _
    ByVal sections As Collection, _
    ByVal newVer As Long _
) As Long
    On Error GoTo ErrHandler

    ' 旧 version 取得（既存 format からの増分判定）
    Dim oldVer As Long
    If modFormatLoader.FormatExists(formatId) Then
        oldVer = ExtractFormatVersion(modFormatLoader.LoadFormat(formatId))
    Else
        oldVer = 0
    End If

    ' format 保存（管理.xlsm 限定 enforce は modFormatLoader 側）
    Dim saveResult As Long
    saveResult = modFormatLoader.SaveFormat(formatId, sections)
    If saveResult <> 0 Then
        SaveFormat = saveResult
        Exit Function
    End If

    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFormatManager", "SaveFormat", _
            "format 保存: " & formatId & " v" & oldVer & "->v" & newVer, _
            formatId, "LOG-FMT-SAVE-OK"
    End If

    ' Q50: 保存時に該当 format の全 knowledge を自動 migrate（version 増分時のみ）
    If newVer > oldVer Then
        ' v2.3 deprecated: auto migrate via clsFieldMigrator is disabled.
        Dim migrated As Long
        migrated = 0
        ' migrated = m_migrator.MigrateAllByFormat(formatId, newVer)
        If Not m_logger Is Nothing Then
            m_logger.LogInfo "clsFormatManager", "SaveFormat", _
                "連動 migrate: " & migrated & " 件 (Q50 自動)", _
                formatId, "LOG-FMT-MIG-OK"
        End If
    End If

    SaveFormat = 0
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFormatManager", "SaveFormat", Err.Description, "", "LOG-FMT-SAVE-ERR"
    End If
    SaveFormat = 2
End Function

' ----------------------------------------------------------------
' Public API: 削除（Q55 knowledge 件数 check）
' ----------------------------------------------------------------

' 概要: format 削除（紐づく knowledge があれば reject、Q55）
' 戻り値: 0 = 成功 / 1 = 管理.xlsm 外 reject / 2 = knowledge 残存 reject / 3 = その他
Public Function DeleteFormat(ByVal formatId As String) As Long
    On Error GoTo ErrHandler

    ' modFormatLoader.DeleteFormat 内で Q55 件数 check + Q19 管理.xlsm enforce を実施
    DeleteFormat = modFormatLoader.DeleteFormat(formatId)

    Select Case DeleteFormat
        Case 0
            If Not m_logger Is Nothing Then
                m_logger.LogInfo "clsFormatManager", "DeleteFormat", _
                    "format 削除: " & formatId, formatId, "LOG-FMT-DEL-OK"
            End If
        Case 1
            If Not m_logger Is Nothing Then
                m_logger.LogWarn "clsFormatManager", "DeleteFormat", _
                    "管理.xlsm 外から呼出 reject: " & formatId, formatId, "LOG-FMT-DEL-INUSE"
            End If
        Case 2
            If Not m_logger Is Nothing Then
                m_logger.LogWarn "clsFormatManager", "DeleteFormat", _
                    "knowledge 残存 reject: " & formatId, formatId, "LOG-FMT-DEL-NOTFOUND"
            End If
    End Select
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFormatManager", "DeleteFormat", Err.Description, "", "LOG-FMT-DEL-ERR"
    End If
    DeleteFormat = 3
End Function

' ----------------------------------------------------------------
' Private ヘルパ
' ----------------------------------------------------------------

Private Function ExtractFormatVersion(ByVal sections As Collection) As Long
    On Error Resume Next
    Dim sec As ClsStanzaSection
    For Each sec In sections
        If sec.SectionName = "FORMAT" Then
            ExtractFormatVersion = CLng(sec.GetValue("FormatVersion"))
            Exit Function
        End If
    Next sec
    ExtractFormatVersion = 1
    If Err.Number <> 0 Then Err.Clear
End Function
```
