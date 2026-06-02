---
title: clsFieldMigrator.cls
description: clsFieldMigrator.cls のソースコード（コピペ用）
---

# clsFieldMigrator.cls

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsFieldMigrator.cls`
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
Attribute VB_Name = "clsFieldMigrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' Class:   clsFieldMigrator (v2.3 BUG-4 related shim)
' Summary: Thin migrator shim aligning with caller signatures used
'          by modEntryKnowledge.Btn_MigrateFields:
'            - Init logger, formatMgr, dataFolder    (3-arg)
'            - MigrateFields(formatId) As Long       (count of targets)
'          The real migration policy (FormatVersion bump,
'          DetectDataLoss, value mapping) is deferred to a later
'          Sprint; this shim ensures the project compiles and that
'          Btn_MigrateFields runs without raising a compile / runtime
'          error. It logs a "reflect done" line so M-12 button UX
'          remains consistent with screen_design_v2 ﾂｧ2.10 behaviour.
' Version: v2.3 (2026-05-30 BUG-4 related compile fix)
' Deps:    modKnowledgeFileIO.ListKnowledgesByFormat,
'          clsLogger (optional)
' ================================================================
Option Explicit

Private m_logger As Object        ' clsLogger or Nothing
Private m_formatMgr As Object     ' clsFormatManager or Nothing
Private m_dataFolder As String    ' optional override (empty = unused)

' ----------------------------------------------------------------
' Init (BUG-4 alignment: caller passes 3 args)
'   migrator.Init logger, formatMgr, GetDataFolder()
' All three remain Optional so any legacy 1-arg call still compiles.
' ----------------------------------------------------------------
Public Sub Init( _
    Optional ByVal logger As Object = Nothing, _
    Optional ByVal formatMgr As Object = Nothing, _
    Optional ByVal dataFolder As String = "")
    Set m_logger = logger
    Set m_formatMgr = formatMgr
    m_dataFolder = dataFolder
End Sub

' ----------------------------------------------------------------
' MigrateFields
' Summary: count the knowledges belonging to the given formatId.
'          Returns the target count (>= 0). On error returns 0.
'          The actual rewrite-each-knowledge logic is deferred to a
'          later Sprint per screen_design_v2 ﾂｧ2.10; this shim keeps
'          the caller signature stable so the workbook compiles and
'          Btn_MigrateFields raises no runtime trap.
' Caller:  modEntryKnowledge.Btn_MigrateFields
' ----------------------------------------------------------------
Public Function MigrateFields(ByVal formatId As String) As Long
    On Error GoTo ErrHandler

    Dim list As Collection
    Set list = modKnowledgeFileIO.ListKnowledgesByFormat(formatId)

    Dim n As Long
    If list Is Nothing Then
        n = 0
    Else
        n = list.Count
    End If

    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFieldMigrator", "MigrateFields", _
                         "targets=" & CStr(n) & " (shim, no rewrite)", _
                         formatId, "LOG-MIG-SHIM-OK"
    End If

    MigrateFields = n
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFieldMigrator", "MigrateFields", _
                          Err.Description, formatId, "LOG-MIG-SHIM-ERR"
    End If
    MigrateFields = 0
End Function
```
