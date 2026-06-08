---
title: clsFieldMigrator.cls
description: clsFieldMigrator.cls 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# clsFieldMigrator.cls

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 繧ｯ繝ｩ繧ｹ繝｢繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`clsFieldMigrator.cls`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
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
'          remains consistent with screen_design_v2 §2.10 behaviour.
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0434] clsFieldMigrator.Init ENTER"  ' [ADR-0100]
    Set m_logger = logger
    Set m_formatMgr = formatMgr
    m_dataFolder = dataFolder
End Sub

' ----------------------------------------------------------------
' MigrateFields
' Summary: count the knowledges belonging to the given formatId.
'          Returns the target count (>= 0). On error returns 0.
'          The actual rewrite-each-knowledge logic is deferred to a
'          later Sprint per screen_design_v2 §2.10; this shim keeps
'          the caller signature stable so the workbook compiles and
'          Btn_MigrateFields raises no runtime trap.
' Caller:  modEntryKnowledge.Btn_MigrateFields
' ----------------------------------------------------------------
Public Function MigrateFields(ByVal formatId As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0435] clsFieldMigrator.MigrateFields ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0436] clsFieldMigrator.MigrateFields EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0437] clsFieldMigrator.MigrateFields EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFieldMigrator", "MigrateFields", _
                          Err.Description, formatId, "LOG-MIG-SHIM-ERR"
    End If
    MigrateFields = 0
End Function
```