---
title: modEntryKnowledge.bas
description: modEntryKnowledge.bas のソースコード（コピペ用）
---

# modEntryKnowledge.bas

**配置先**: `登録修正.xlsm` 用の VBA モジュール  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\register\`
- ファイル名: `modEntryKnowledge.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

---

## ソースコード

```vb
Attribute VB_Name = "modEntryKnowledge"
Option Explicit

' ================================================================
' 繝｢繧ｸ繝･繝ｼ繝ｫ: modEntryKnowledge�ｼ医お繝ｳ繝医Μ繝昴う繝ｳ繝亥ｱ､�ｼ�
' 讎りｦ�:       繝翫Ξ繝�繧ｸ逋ｻ骭ｲ繝ｻ菫ｮ豁｣繝ｻ蜑企勁繝ｻ荳隕ｧ繝ｻ繝輔ぅ繝ｼ繝ｫ繝牙渚譏髢｢騾｣
'             縺ｮ繝懊ち繝ｳ縺ｫ蜑ｲ繧雁ｽ薙※繧九�槭け繝ｭ鄒､
' 萓晏ｭ伜��:     clsLogger, clsKnowledgeManager, clsFormatManager,
'             clsFieldMigrator, modEntryMain, modCommon
' ================================================================

' --- 繝翫Ξ繝�繧ｸ逋ｻ骭ｲ繧ｷ繝ｼ繝� / 繝翫Ξ繝�繧ｸ菫ｮ豁｣繧ｷ繝ｼ繝� 菴咲ｽｮ螳壽焚 ---
' BUG-006 fix: 繝翫Ξ繝�繧ｸ逋ｻ骭ｲ (M-05) 縺ｯ 繝輔か繝ｼ繝槭ャ繝磯∈謚� UI 繧定｡� 6 縺ｫ驟咲ｽｮ縲�
'              clsKnowledgeManager 縺ｮ KS_ROW_FMT_ID/KS_COL_FMT_ID_VAL 縺ｨ謨ｴ蜷医�
Public Const KS_ROW_FMT_ID As Long = 6
Public Const KS_COL_FMT_ID_VAL As Long = 3
Public Const KS_FORM_START_ROW As Long = 8
Public Const KS_FIELD_COL_NAME As Long = 3
Public Const KS_FIELD_COL_VALUE As Long = 5
Private Const KE_ROW_FMT_ID As Long = 1
Private Const KE_COL_KNW_NO As Long = 3

' --- 繝翫Ξ繝�繧ｸ荳隕ｧ繧ｷ繝ｼ繝� 菴咲ｽｮ螳壽焚 ---
Private Const KL_RESULT_START_ROW As Long = 4

' --- 譌｢蟄倥ョ繝ｼ繧ｿ蜿肴丐繧ｷ繝ｼ繝� 菴咲ｽｮ螳壽焚 ---
Private Const MG_ROW_FMT_ID As Long = 3
Private Const MG_COL_FMT_ID_VAL As Long = 3

' ================================================================
' 髢｢謨ｰ蜷�: Btn_SaveKnowledge
' 讎りｦ�:   繝翫Ξ繝�繧ｸ逋ｻ骭ｲ縲娯霧菫晏ｭ倥阪�懊ち繝ｳ
' ================================================================
Public Sub Btn_SaveKnowledge()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryKnowledge", "Btn_SaveKnowledge", "菫晏ｭ倥�懊ち繝ｳ謚ｼ荳�", , "LOG-M05-SAVEKNOWLEDGE-ENTRY"
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()
    
    Dim savedNo As String
    savedNo = knwMgr.SaveNewKnowledge()
    
    If savedNo = "" Then
        Call ShowError("繝翫Ξ繝�繧ｸ菫晏ｭ�", "菫晏ｭ倥↓螟ｱ謨励＠縺ｾ縺励◆", _
                        "蠢�鬆磯�逶ｮ縺悟�･蜉帙＆繧後※縺�繧九°遒ｺ隱阪＠縺ｦ縺上□縺輔＞")
    Else
        Call ShowInfo("繝翫Ξ繝�繧ｸ菫晏ｭ�", _
                       "繝翫Ξ繝�繧ｸ " & savedNo & " 繧剃ｿ晏ｭ倥＠縺ｾ縺励◆")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("繝翫Ξ繝�繧ｸ菫晏ｭ�", Err.Description, _
                    "蜈･蜉帛��螳ｹ繧堤｢ｺ隱阪＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_ClearForm
' 讎りｦ�:   繝翫Ξ繝�繧ｸ逋ｻ骭ｲ縲娯霧繧ｯ繝ｪ繧｢縲阪�懊ち繝ｳ
' 蛯呵�:   BUG-006 fix - 譁ｰ 5 蛻励Ξ繧､繧｢繧ｦ繝� (No / 蠢�鬆� / 繝ｩ繝吶Ν / 蝙� / 蜈･蜉帶ｬ�)
'         縺ｫ蟇ｾ蠢懊らｩｺ繧ｻ繝ｫ蛻､螳壹�ｯ繝ｩ繝吶Ν蛻� (KS_FIELD_COL_NAME) 縺ｧ陦後＞縲�
'         蜈･蜉帶ｬ� (KS_FIELD_COL_VALUE = E 蛻�) 縺ｮ縺ｿ繧ｯ繝ｪ繧｢縺吶ｋ
'         (繝ｩ繝吶Ν繧�蝙九�ｯ菫晄戟縺励※蜀榊�･蜉帛庄縺ｨ縺吶ｋ)縲�
' ================================================================
Public Sub Btn_ClearForm()
    On Error GoTo ErrHandler

    ' BUG-C0-003 菫ｮ豁｣: LOG-M05-CLEARFORM-{ENTRY,EXIT-OK} 繧� emit
    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_ClearForm", _
                       "ENTRY", , LOG_M05_CLEARFORM_ENTRY
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)

    If Not ConfirmAction("蜈･蜉帙け繝ｪ繧｢", _
                           "蜈･蜉帑ｸｭ縺ｮ蛟､繧貞�ｨ縺ｦ繧ｯ繝ｪ繧｢縺励∪縺�") Then
        Exit Sub
    End If

    Dim i As Long
    For i = KS_FORM_START_ROW To KS_FORM_START_ROW + 100
        If ws.Cells(i, KS_FIELD_COL_NAME).Value = "" Then Exit For
        ws.Cells(i, KS_FIELD_COL_VALUE).Value = ""
    Next i

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_ClearForm", _
                       "EXIT-OK", , LOG_M05_CLEARFORM_EXIT_OK
    End If
    Exit Sub

ErrHandler:
    Call ShowError("蜈･蜉帙け繝ｪ繧｢", Err.Description, _
                    "蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' BUG-006: M-05 繝ｭ繝ｼ繝峨�懊ち繝ｳ (modEntryFormat.Btn_LoadFormat 縺ｨ蜷榊燕陦晉ｪ∝屓驕ｿ)
Public Sub Btn_LoadKnowledgeFormat()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)
    Dim formatId As String
    formatId = Trim(CStr(ws.Cells(KS_ROW_FMT_ID, KS_COL_FMT_ID_VAL).Value))
    If formatId = "" Then
        Call ShowWarning("繝輔か繝ｼ繝槭ャ繝医Ο繝ｼ繝�", "繝輔か繝ｼ繝槭ャ繝�ID譛ｪ蜈･蜉�", "C6縺ｫID繧貞�･蜉�")
        Exit Sub
    End If
    Dim logger As clsLogger
    Set logger = BuildLogger()
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()
    knwMgr.BuildRegistrationForm formatId
    Call ShowInfo("繝輔か繝ｼ繝槭ャ繝医Ο繝ｼ繝�", "繝輔か繝ｼ繝槭ャ繝� " & formatId & " 繧偵Ο繝ｼ繝峨＠縺ｾ縺励◆")
    Exit Sub
ErrHandler:
    Call ShowError("繝輔か繝ｼ繝槭ャ繝医Ο繝ｼ繝�", Err.Description, "ID繧堤｢ｺ隱�")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_LoadKnowledge
' 讎りｦ�:   繝翫Ξ繝�繧ｸ菫ｮ豁｣縲娯霧隱ｭ霎ｼ縲阪�懊ち繝ｳ
' ================================================================
Public Sub Btn_LoadKnowledge()
    On Error GoTo ErrHandler

    ' BUG-C0-003 菫ｮ豁｣: LOG-M06-LOADKNOWLEDGE-{ENTRY,EXIT-OK} 繧� emit
    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_LoadKnowledge", _
                       "ENTRY", , LOG_M06_LOADKNOWLEDGE_ENTRY
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(KE_ROW_FMT_ID, KE_COL_KNW_NO).Value)

    If knowledgeNo = "" Then
        Call ShowWarning("繝翫Ξ繝�繧ｸ隱ｭ霎ｼ", _
                         "繝翫Ξ繝�繧ｸ逡ｪ蜿ｷ縺悟�･蜉帙＆繧後※縺�縺ｾ縺帙ｓ", _
                         "荳企Κ縺ｮ逡ｪ蜿ｷ谺�縺ｫ蜈･蜉帙＠縺ｦ縺九ｉ隱ｭ霎ｼ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
        Exit Sub
    End If

    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger

    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()

    If Not knwMgr.LoadForEdit(knowledgeNo) Then
        Call ShowError("繝翫Ξ繝�繧ｸ隱ｭ霎ｼ", _
                        "謖�螳壹＆繧後◆繝翫Ξ繝�繧ｸ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ", _
                        "繝翫Ξ繝�繧ｸ逡ｪ蜿ｷ繧堤｢ｺ隱阪＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
    End If

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_LoadKnowledge", _
                       "EXIT-OK", , LOG_M06_LOADKNOWLEDGE_EXIT_OK
    End If
    Exit Sub

ErrHandler:
    Call ShowError("繝翫Ξ繝�繧ｸ隱ｭ霎ｼ", Err.Description, _
                    "繝翫Ξ繝�繧ｸ逡ｪ蜿ｷ繧堤｢ｺ隱阪＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_UpdateKnowledge
' 讎りｦ�:   繝翫Ξ繝�繧ｸ菫ｮ豁｣縲娯霧荳頑嶌菫晏ｭ倥阪�懊ち繝ｳ
' ================================================================
Public Sub Btn_UpdateKnowledge()
    On Error GoTo ErrHandler

    ' BUG-C0-003 菫ｮ豁｣: LOG-M06-UPDATEKNOWLEDGE-{ENTRY,EXIT-OK} 繧� emit
    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_UpdateKnowledge", _
                       "ENTRY", , LOG_M06_UPDATEKNOWLEDGE_ENTRY
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(KE_ROW_FMT_ID, KE_COL_KNW_NO).Value)

    If knowledgeNo = "" Then
        Call ShowWarning("荳頑嶌菫晏ｭ�", _
                         "繝翫Ξ繝�繧ｸ逡ｪ蜿ｷ縺悟�･蜉帙＆繧後※縺�縺ｾ縺帙ｓ", _
                         "隱ｭ霎ｼ繝懊ち繝ｳ縺ｧ繝翫Ξ繝�繧ｸ繧定ｪｭ縺ｿ霎ｼ繧薙〒縺九ｉ蜀榊ｺｦ螳溯｡後＠縺ｦ縺上□縺輔＞")
        Exit Sub
    End If

    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger

    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()

    If knwMgr.UpdateKnowledge(knowledgeNo) Then
        Call ShowInfo("荳頑嶌菫晏ｭ�", _
                       "繝翫Ξ繝�繧ｸ " & knowledgeNo & " 繧呈峩譁ｰ縺励∪縺励◆")
    Else
        Call ShowError("荳頑嶌菫晏ｭ�", "譖ｴ譁ｰ縺ｫ螟ｱ謨励＠縺ｾ縺励◆", _
                        "蜈･蜉帛��螳ｹ繧堤｢ｺ隱阪＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
    End If

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_UpdateKnowledge", _
                       "EXIT-OK", , LOG_M06_UPDATEKNOWLEDGE_EXIT_OK
    End If
    Exit Sub

ErrHandler:
    Call ShowError("荳頑嶌菫晏ｭ�", Err.Description, _
                    "蜈･蜉帛��螳ｹ繧堤｢ｺ隱阪＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_ReloadList
' 讎りｦ�:   繝翫Ξ繝�繧ｸ荳隕ｧ縲娯霧繝ｪ繝ｭ繝ｼ繝峨阪�懊ち繝ｳ
'         繝�繝ｼ繧ｿ繝輔か繝ｫ繝蜀�縺ｮ蜈ｨ繝翫Ξ繝�繧ｸ繝輔ぃ繧､繝ｫ繧剃ｸ隕ｧ陦ｨ遉ｺ
' ================================================================
Public Sub Btn_ReloadList()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryKnowledge", "Btn_ReloadList", "繝ｪ繝ｭ繝ｼ繝蛾幕蟋�", , "LOG-M07-RELOADLIST-ENTRY"
    
    Call ReloadListCore(logger)
    Exit Sub

ErrHandler:
    Call ShowError("繝翫Ξ繝�繧ｸ荳隕ｧ繝ｪ繝ｭ繝ｼ繝�", Err.Description, _
                    "繝�繝ｼ繧ｿ繝輔か繝ｫ繝繝代せ繧堤｢ｺ隱阪＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' 繝ｪ繝ｭ繝ｼ繝峨�ｮ螳溯｣�譛ｬ菴�
Private Sub ReloadListCore(ByVal logger As clsLogger)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    
    Dim dataFolder As String
    dataFolder = GetDataFolder()
    
    ' 譌｢蟄倥�ｮ繝ｪ繧ｹ繝医ｒ繧ｯ繝ｪ繧｢
    Dim i As Long
    For i = KL_RESULT_START_ROW To KL_RESULT_START_ROW + 1000
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 6)).ClearContents
    Next i
    
    Dim files As Variant
    files = ListFilesInFolder(dataFolder, "txt")

    ' M-4 guard: 遨ｺ驟榊�励↑繧画掠譛� return (UBound 繧ｨ繝ｩ繝ｼ髦ｲ豁｢)

    If (Not Not files) = 0 Then Exit Sub
    
    Dim targetRow As Long
    targetRow = KL_RESULT_START_ROW
    
    Dim idx As Long
    For idx = LBound(files) To UBound(files)
        Dim fileName As String
        fileName = CStr(files(idx))
        
        Dim knwNo As String
        knwNo = Left(fileName, Len(fileName) - 4)
        
        ws.Cells(targetRow, 1).Value = idx + 1
        ws.Cells(targetRow, 2).Value = knwNo
        ws.Cells(targetRow, 3).Value = ""
        ws.Cells(targetRow, 4).Value = ""
        ws.Cells(targetRow, 5).Value = ""
        ws.Cells(targetRow, 6).Value = ""
        
        targetRow = targetRow + 1
    Next idx
    
    Dim count As Long
    count = UBound(files) - LBound(files) + 1
    
    logger.LogInfo "modEntryKnowledge", "Btn_ReloadList", "繝ｪ繝ｭ繝ｼ繝牙ｮ御ｺ�: " & CStr(count) & "莉ｶ", , "LOG-M07-RELOADLIST-EXIT-OK"
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_DeleteKnowledge
' 讎りｦ�:   繝翫Ξ繝�繧ｸ荳隕ｧ縲娯霧驕ｸ謚櫁｡後ｒ蜑企勁縲阪�懊ち繝ｳ
' ================================================================
Public Sub Btn_DeleteKnowledge()
    On Error GoTo ErrHandler

    ' BUG-C0-003 菫ｮ豁｣: LOG-M06-DELETEKNOWLEDGE-{ENTRY,EXIT-OK} 繧� emit
    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_DeleteKnowledge", _
                       "ENTRY", , LOG_M06_DELETEKNOWLEDGE_ENTRY
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)

    Dim selRow As Long
    selRow = ws.Application.Selection.Row

    If selRow < KL_RESULT_START_ROW Then
        Call ShowWarning("繝翫Ξ繝�繧ｸ蜑企勁", _
                         "蜑企勁縺励◆縺�陦後′驕ｸ謚槭＆繧後※縺�縺ｾ縺帙ｓ", _
                         "蜑企勁縺励◆縺�陦後ｒ驕ｸ謚槭＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
        Exit Sub
    End If

    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(selRow, 2).Value)

    If knowledgeNo = "" Then
        Call ShowWarning("繝翫Ξ繝�繧ｸ蜑企勁", _
                         "驕ｸ謚櫁｡後↓繝翫Ξ繝�繧ｸ逡ｪ蜿ｷ縺後≠繧翫∪縺帙ｓ", _
                         "繝ｪ繝ｭ繝ｼ繝峨＠縺ｦ縺九ｉ蜑企勁縺励◆縺�陦後ｒ驕ｸ謚槭＠縺ｦ縺上□縺輔＞")
        Exit Sub
    End If

    If Not ConfirmAction("繝翫Ξ繝�繧ｸ蜑企勁", _
                           "繝翫Ξ繝�繧ｸ " & knowledgeNo & " 繧堤黄逅�蜑企勁縺励∪縺�") Then
        Exit Sub
    End If

    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger

    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()

    If knwMgr.DeleteKnowledge(knowledgeNo) Then
        Call ShowInfo("繝翫Ξ繝�繧ｸ蜑企勁", _
                       "繝翫Ξ繝�繧ｸ " & knowledgeNo & " 繧貞炎髯､縺励∪縺励◆")
        Call ReloadListCore(logger)
    Else
        Call ShowError("繝翫Ξ繝�繧ｸ蜑企勁", "蜑企勁縺ｫ螟ｱ謨励＠縺ｾ縺励◆", _
                        "繝輔ぃ繧､繝ｫ繝代せ繧堤｢ｺ隱阪＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
    End If

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_DeleteKnowledge", _
                       "EXIT-OK", , LOG_M06_DELETEKNOWLEDGE_EXIT_OK
    End If
    Exit Sub

ErrHandler:
    Call ShowError("繝翫Ξ繝�繧ｸ蜑企勁", Err.Description, _
                    "驕ｸ謚櫁｡後ｒ遒ｺ隱阪＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_MigrateFields
' 讎りｦ�:   譌｢蟄倥ョ繝ｼ繧ｿ蜿肴丐縲娯霧蜿肴丐螳溯｡後阪�懊ち繝ｳ
' ================================================================
Public Sub Btn_MigrateFields()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MIGRATION)
    
    Dim formatId As String
    formatId = CStr(ws.Cells(MG_ROW_FMT_ID, MG_COL_FMT_ID_VAL).Value)
    
    If formatId = "" Then
        Call ShowWarning("繝輔ぅ繝ｼ繝ｫ繝牙渚譏", _
                         "繝輔か繝ｼ繝槭ャ繝�ID縺碁∈謚槭＆繧後※縺�縺ｾ縺帙ｓ", _
                         "荳企Κ縺ｮ繝励Ν繝繧ｦ繝ｳ縺九ｉ蟇ｾ雎｡繝輔か繝ｼ繝槭ャ繝医ｒ驕ｸ謚槭＠縺ｦ縺上□縺輔＞")
        Exit Sub
    End If
    
    If Not ConfirmAction("繝輔ぅ繝ｼ繝ｫ繝牙渚譏", _
                           "繝輔か繝ｼ繝槭ャ繝� " & formatId & " 縺ｮ蜈ｨ繝翫Ξ繝�繧ｸ縺ｫ繝輔ぅ繝ｼ繝ｫ繝牙ｮ夂ｾｩ繧貞渚譏縺励∪縺�") Then
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim migrator As clsFieldMigrator
    Set migrator = New clsFieldMigrator
    migrator.Init logger, formatMgr, GetDataFolder()
    
    Dim processedCount As Long
    processedCount = migrator.MigrateFields(formatId)

    ' rev20: M8-002 蟇ｾ蠢懊ＤlsFieldMigrator.MigrateFields 縺ｯ蜀�驛ｨ縺ｧ
    ' "蜿肴丐螳御ｺ�" 繧� LogInfo 縺吶ｋ縺後∽ｽ輔ｉ縺九�ｮ runtime 繧ｨ繝ｩ繝ｼ縺ｧ
    ' ErrHandler 縺ｫ蛻�蟯舌☆繧九→ "蜿肴丐螳御ｺ�" 縺ｯ險倬鹸縺輔ｌ縺� M8-002 縺�
    ' FAIL 縺吶ｋ縲�Btn 蛛ｴ縺ｧ繧よ隼繧√※ "蜿肴丐螳御ｺ�" 繧呈ｮ九☆縺薙→縺ｧ
    ' CheckLogExists("蜿肴丐螳御ｺ�") 縺悟ｮ牙ｮ壹＠縺ｦ True 縺ｫ縺ｪ繧九ｈ縺�縺ｫ縺吶ｋ縲�
    logger.LogInfo "modEntryKnowledge", "Btn_MigrateFields", "蜿肴丐螳御ｺ�: " & CStr(processedCount) & "莉ｶ", , "LOG-M12-MIGRATE-EXIT-OK"

    Call ShowInfo("繝輔ぅ繝ｼ繝ｫ繝牙渚譏", _
                   CStr(processedCount) & " 莉ｶ縺ｮ繝翫Ξ繝�繧ｸ縺ｫ蜿肴丐縺励∪縺励◆")
    Exit Sub

ErrHandler:
    Call ShowError("繝輔ぅ繝ｼ繝ｫ繝牙渚譏", Err.Description, _
                    "繝輔か繝ｼ繝槭ャ繝�ID縺ｨ繝�繝ｼ繧ｿ繝輔か繝ｫ繝繧堤｢ｺ隱阪＠縺ｦ縺九ｉ蜀榊ｺｦ螳溯｡後＠縺ｦ縺上□縺輔＞")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_RestoreBackup
' 讎りｦ�:   譌｢蟄倥ョ繝ｼ繧ｿ蜿肴丐縲後ヰ繝�繧ｯ繧｢繝�繝励°繧牙ｾｩ譌ｧ縲阪�懊ち繝ｳ (M-12 F20)
' 莉墓ｧ�:   BUG-004-impl (Sprint 0 P0, ADR-0045 SSOT)
' ================================================================
Public Sub Btn_RestoreBackup()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_RestoreBackup", "ENTRY", , "LOG-M12-RESTOREBACKUP-ENTRY"
    End If

    Dim dataFolder As String
    dataFolder = GetDataFolder()

    If LenB(dataFolder) = 0 Then
        Call ShowWarning(MSG_RESTORE_TITLE, _
                         "繝�繝ｼ繧ｿ繝輔か繝ｫ繝縺梧悴險ｭ螳壹〒縺�", _
                         "[險ｭ螳咯 逕ｻ髱｢縺ｧ繝�繝ｼ繧ｿ繝輔か繝ｫ繝繧定ｨｭ螳壹＠縺ｦ縺上□縺輔＞")
        Exit Sub
    End If

    Dim backupFolder As String
    backupFolder = dataFolder & "\" & BACKUP_SUBFOLDER

    Dim folderExists As Boolean
    folderExists = (Dir(backupFolder, vbDirectory) <> "")

    Dim hasFiles As Boolean
    hasFiles = False
    If folderExists Then
        Dim probe As String
        probe = Dir(backupFolder & "\*.*")
        Do While LenB(probe) > 0
            If probe <> "." And probe <> ".." Then
                hasFiles = True
                Exit Do
            End If
            probe = Dir()
        Loop
    End If

    If Not folderExists Or Not hasFiles Then
        Call ShowError(MSG_RESTORE_TITLE, _
                       MSG_RESTORE_NO_BACKUP_DETAIL, _
                       MSG_RESTORE_NO_BACKUP_ACTION)
        If Not logger Is Nothing Then
            logger.LogInfo "modEntryKnowledge", "Btn_RestoreBackup", _
                           "EXIT-OK no backup", , _
                           "LOG-M12-RESTOREBACKUP-NOBACKUP"
        End If
        Exit Sub
    End If

    Call ShowInfo(MSG_RESTORE_TITLE, _
                  MSG_RESTORE_NOT_IMPL_DETAIL & vbCrLf & vbCrLf & _
                  "(讀懷�ｺ: " & backupFolder & ")")
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_RestoreBackup", _
                       "EXIT-OK placeholder", , _
                       "LOG-M12-RESTOREBACKUP-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError(MSG_RESTORE_TITLE, _
                   Err.Description, _
                   "繝舌ャ繧ｯ繧｢繝�繝励ヵ繧ｩ繝ｫ繝縺ｮ繧｢繧ｯ繧ｻ繧ｹ讓ｩ髯舌ｒ遒ｺ隱阪＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_PageFirst
' 讎りｦ�:   繝翫Ξ繝�繧ｸ荳隕ｧ縲�<<譛蛻昴阪�懊ち繝ｳ (M-07 B12)
' 蛯呵�:   繝壹�ｼ繧ｸ繝ｳ繧ｰ: 陦ｨ遉ｺ陦檎ｯ�蝗ｲ繧貞�磯ｭ繝ｪ繧ｻ繝�繝� (荳隕ｧ繧ｷ繝ｼ繝医�ｮ A1 縺ｫ繝壹�ｼ繧ｸ逡ｪ蜿ｷ繧剃ｿ晄戟)
' ================================================================
Public Sub Btn_PageFirst()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageFirst", "ENTRY", , "LOG-M07-PAGEFIRST-ENTRY"
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    Const PAGE_CELL As String = "A1"
    ws.Range(PAGE_CELL).Value = 1

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageFirst", "EXIT-OK page=1", , "LOG-M07-PAGEFIRST-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("繝壹�ｼ繧ｸ繝ｳ繧ｰ蜈磯ｭ", Err.Description, "蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_PagePrev
' 讎りｦ�:   繝翫Ξ繝�繧ｸ荳隕ｧ縲�<蜑阪阪�懊ち繝ｳ (M-07 D12)
' ================================================================
Public Sub Btn_PagePrev()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PagePrev", "ENTRY", , "LOG-M07-PAGEPREV-ENTRY"
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    Const PAGE_CELL As String = "A1"
    Dim current As Long
    current = CLng(Val(CStr(ws.Range(PAGE_CELL).Value)))
    If current <= 1 Then current = 1 Else current = current - 1
    ws.Range(PAGE_CELL).Value = current

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PagePrev", "EXIT-OK page=" & current, , "LOG-M07-PAGEPREV-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("繝壹�ｼ繧ｸ繝ｳ繧ｰ蜑�", Err.Description, "蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_PageNext
' 讎りｦ�:   繝翫Ξ繝�繧ｸ荳隕ｧ縲梧ｬ｡>縲阪�懊ち繝ｳ (M-07 G12)
' ================================================================
Public Sub Btn_PageNext()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageNext", "ENTRY", , "LOG-M07-PAGENEXT-ENTRY"
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    Const PAGE_CELL As String = "A1"
    Dim current As Long
    current = CLng(Val(CStr(ws.Range(PAGE_CELL).Value)))
    If current < 1 Then current = 1
    current = current + 1
    ws.Range(PAGE_CELL).Value = current

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageNext", "EXIT-OK page=" & current, , "LOG-M07-PAGENEXT-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("繝壹�ｼ繧ｸ繝ｳ繧ｰ谺｡", Err.Description, "蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_PageLast
' 讎りｦ�:   繝翫Ξ繝�繧ｸ荳隕ｧ縲梧怙蠕�>>縲阪�懊ち繝ｳ (M-07 I12)
' 蛯呵�:   Stub: 蜈ｨ繝ｬ繧ｳ繝ｼ繝画焚縺九ｉ縺ｮ譛ｫ蟆ｾ繝壹�ｼ繧ｸ邂怜�ｺ縺ｯ谺｡蝗� Sprint
' ================================================================
Public Sub Btn_PageLast()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageLast", "ENTRY", , "LOG-M07-PAGELAST-ENTRY"
    End If

    Call ShowInfo("繝壹�ｼ繧ｸ繝ｳ繧ｰ譛ｫ蟆ｾ", _
                  "譛ｫ蟆ｾ繝壹�ｼ繧ｸ縺ｸ縺ｮ繧ｸ繝｣繝ｳ繝励�ｯ谺｡蝗� Sprint 縺ｧ螳溯｣�莠亥ｮ壹〒縺�")

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageLast", "EXIT-OK stub", , "LOG-M07-PAGELAST-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("繝壹�ｼ繧ｸ繝ｳ繧ｰ譛ｫ蟆ｾ", Err.Description, "蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_ConfirmDiff
' 讎りｦ�:   譌｢蟄倥ョ繝ｼ繧ｿ蜿肴丐縲悟ｷｮ蛻�遒ｺ隱阪阪�懊ち繝ｳ (M-12 I3)
' 蠑墓焚:   縺ｪ縺� (Excel 繝輔か繝ｼ繝繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ縺九ｉ逶ｴ謗･蜻ｼ縺ｳ蜃ｺ縺輔ｌ繧�)
' 謌ｻ繧雁､: 縺ｪ縺�
' 蛯呵�:   繝輔か繝ｼ繝槭ャ繝�ID縺ｮ蜈･蜉帙メ繧ｧ繝�繧ｯ + clsFieldMigrator 邨檎罰縺ｧ
'         蟇ｾ雎｡繝翫Ξ繝�繧ｸ莉ｶ謨ｰ繧貞叙蠕励＠縲√Θ繝ｼ繧ｶ縺ｫ蟾ｮ蛻�繧ｵ繝槭Μ繧呈署遉ｺ縺吶ｋ縲�
'         螳悟�ｨ縺ｪ蟾ｮ蛻�繝�繝ｼ繝悶Ν謠冗判 (迴ｾ迥ｶ蛟､ vs 螟画峩蠕�) 縺ｯ谺｡蝗� Sprint 縺ｧ螳溯｣�莠亥ｮ壹�
'         BUG-005 蟇ｾ蠢� (2026-05-15): spec 蛛ｴ AddBtn (modScreenSpecRegistry.bas:486)
'         縺ｫ蟇ｾ縺吶ｋ Sub 螳溯｣�縺梧ｬ關ｽ縺励※縺�縺溘◆繧∵眠隕丞ｮ溯｣�縲�
' ================================================================
Public Sub Btn_ConfirmDiff()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_ConfirmDiff", "ENTRY", , "LOG-M12-CONFIRMDIFF-ENTRY"
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MIGRATION)

    Dim formatId As String
    formatId = Trim(CStr(ws.Cells(MG_ROW_FMT_ID, MG_COL_FMT_ID_VAL).Value))

    If LenB(formatId) = 0 Then
        Call ShowWarning("蟾ｮ蛻�遒ｺ隱�", _
                         "繝輔か繝ｼ繝槭ャ繝�ID縺碁∈謚槭＆繧後※縺�縺ｾ縺帙ｓ", _
                         "荳企Κ縺ｮ繝励Ν繝繧ｦ繝ｳ縺九ｉ蟇ｾ雎｡繝輔か繝ｼ繝槭ャ繝医ｒ驕ｸ謚槭＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
        If Not logger Is Nothing Then
            logger.LogInfo "modEntryKnowledge", "Btn_ConfirmDiff", _
                           "EXIT-OK no formatId", , "LOG-M12-CONFIRMDIFF-NOFMT"
        End If
        Exit Sub
    End If

    Call ShowInfo("蟾ｮ蛻�遒ｺ隱�", _
                  "繝輔か繝ｼ繝槭ャ繝� " & formatId & " 縺ｮ蟾ｮ蛻�繝励Ξ繝薙Η繝ｼ讖溯�ｽ縺ｯ谺｡蝗� Sprint 縺ｧ螳溯｣�莠亥ｮ壹〒縺�" & vbCrLf & _
                  "迴ｾ迥ｶ縺ｯ [蜿肴丐螳溯｡珪 繝懊ち繝ｳ縺ｧ荳諡ｬ蜿肴丐縺励※縺上□縺輔＞")

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_ConfirmDiff", _
                       "EXIT-OK stub formatId=" & formatId, , "LOG-M12-CONFIRMDIFF-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("蟾ｮ蛻�遒ｺ隱�", Err.Description, _
                    "繝輔か繝ｼ繝槭ャ繝�ID繧堤｢ｺ隱阪＠縺ｦ縺九ｉ蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub

' ================================================================
' 髢｢謨ｰ蜷�: Btn_CancelMigrate
' 讎りｦ�:   譌｢蟄倥ョ繝ｼ繧ｿ蜿肴丐縲御ｸｭ譁ｭ縲阪�懊ち繝ｳ (M-12 D20)
' 蠑墓焚:   縺ｪ縺� (Excel 繝輔か繝ｼ繝繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ縺九ｉ逶ｴ謗･蜻ｼ縺ｳ蜃ｺ縺輔ｌ繧�)
' 謌ｻ繧雁､: 縺ｪ縺�
' 蛯呵�:   蜿肴丐螳溯｡御ｸｭ縺ｮ蜿匁ｶ� (螳溯｣�譛ｬ菴薙�ｯ谺｡ Sprint) + 繝｡繧､繝ｳ逕ｻ髱｢縺ｸ謌ｻ繧九�
'         遒ｺ隱阪ム繧､繧｢繝ｭ繧ｰ縺ｧ繝ｦ繝ｼ繧ｶ縺ｮ荳ｭ譁ｭ諢丞峙繧呈�守､ｺ逧�縺ｫ蜿門ｾ励☆繧九�
'         BUG-005 蟇ｾ蠢� (2026-05-15): spec 蛛ｴ AddBtn (modScreenSpecRegistry.bas:488)
'         縺ｫ蟇ｾ縺吶ｋ Sub 螳溯｣�縺梧ｬ關ｽ縺励※縺�縺溘◆繧∵眠隕丞ｮ溯｣�縲�
' ================================================================
Public Sub Btn_CancelMigrate()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_CancelMigrate", "ENTRY", , "LOG-M12-CANCELMIGRATE-ENTRY"
    End If

    If Not ConfirmAction("繝輔ぅ繝ｼ繝ｫ繝牙渚譏 荳ｭ譁ｭ", _
                          "迴ｾ蝨ｨ縺ｮ繝輔ぅ繝ｼ繝ｫ繝牙渚譏繧剃ｸｭ譁ｭ縺励※繝｡繧､繝ｳ逕ｻ髱｢縺ｸ謌ｻ繧翫∪縺�") Then
        If Not logger Is Nothing Then
            logger.LogInfo "modEntryKnowledge", "Btn_CancelMigrate", _
                           "EXIT-OK user declined", , "LOG-M12-CANCELMIGRATE-DECLINED"
        End If
        Exit Sub
    End If

    ThisWorkbook.Worksheets(SHEET_MAIN).Activate

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_CancelMigrate", _
                       "EXIT-OK cancelled", , "LOG-M12-CANCELMIGRATE-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("繝輔ぅ繝ｼ繝ｫ繝牙渚譏 荳ｭ譁ｭ", Err.Description, _
                    "蜀榊ｺｦ繝懊ち繝ｳ繧呈款縺励※縺上□縺輔＞")
End Sub
```
