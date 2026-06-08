---
title: modPreviewForm.bas
description: modPreviewForm.bas 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# modPreviewForm.bas

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`modPreviewForm.bas`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
```vb
Attribute VB_Name = "modPreviewForm"
' ================================================================
' modPreviewForm (v2.5, 2026-06-03, ADR-0099)
'
' Role:
'   M-04 format preview entry point. Hosts the public Sub
'   ShowPreviewForm(formatId) called from M-02 Btn_PreviewFormat
'   and M-03 Btn_PreviewInDesign. Wraps clsUserFormRenderer with
'   mode="preview" so the format's input form layout renders as
'   a modal UserForm with read-only controls and a single close
'   button.
'
' ADR-0099 (2026-06-03):
'   The previous M-04 worksheet ("プレビュー" sheet) is retired.
'   clsSetupOrchestrator.SHEETS_KANRI no longer includes M-04 so
'   the sheet is never created. ui_seed/管理/M-04.txt is renamed
'   to M-04.txt.deprecated. The wired buttons now route here.
'
' Why not call clsUserFormRenderer.ShowForm("view") instead:
'   "view" mode resolves m_formatId from the loaded knowledge
'   record. Preview has no knowledge record (knowledgeId="") so
'   m_formatId would never be populated. The dedicated "preview"
'   mode treats formatId as an explicit parameter (m_formatId)
'   and skips the knowledge load, the register-style buttons, and
'   the edit/view buttons -- only "閉じる" remains.
'
' Dependencies:
'   clsUserFormRenderer (existing, ADR-0073 dynamic UserForm)
'   modEntryUserForm.GetSharedRenderer (single-instance lookup)
' ================================================================
Option Explicit

' xlsm display name used as scope label only (the form is workbook-local).
' Use ChrW so the .bas source is ASCII-safe through CP932 round trips.
Private Function KANRI_NAME() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1827] modPreviewForm.KANRI_NAME ENTER"  ' [ADR-0100]
    KANRI_NAME = ChrW(&H7BA1) & ChrW(&H7406)   ' 管理
End Function

' ================================================================
' Public Sub: ShowPreviewForm
' Role:
'   Render the format's input form layout as a modal UserForm.
'   Each field of <formatId> becomes a labeled, read-only control
'   per clsUserFormRenderer's existing FieldType -> control map:
'     単一行 -> TextBox (MultiLine=False)
'     複数行 -> TextBox (MultiLine=True)
'     日時   -> TextBox (single line)
'     選択   -> ComboBox (DropdownList)
'   Required fields show the "必須" badge. The form title is
'   "プレビュー: <formatId>". One "閉じる" button dismisses the
'   form (Unload Me). No persistence (read-only display only).
' Args:
'   formatId - target format id. If empty the call is a no-op.
' Returns: nothing.
' Notes:
'   - clsUserFormRenderer.ShowForm returns "" for preview mode
'     because no return id is set (no save / no delete buttons).
'   - On error a Debug.Print line is emitted (no MsgBox so the
'     headless E2E harness does not block).
' ================================================================
Public Sub ShowPreviewForm(ByVal formatId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1828] modPreviewForm.ShowPreviewForm ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If Len(Trim(formatId)) = 0 Then
        Debug.Print "[modPreviewForm.ShowPreviewForm] empty formatId, no-op"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1829] modPreviewForm.ShowPreviewForm EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    Dim renderer As clsUserFormRenderer
    Set renderer = New clsUserFormRenderer
    ' xlsmName=管理 / mode=preview / knowledgeId="" / readOnlyFormat=True.
    ' readOnlyFormat=True keeps the format selector locked (preview is
    ' scoped to one format -- changing it would lose the M-03 context).
    Dim ignored As String
    ignored = renderer.ShowFormPreview(KANRI_NAME(), formatId)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1830] modPreviewForm.ShowPreviewForm EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1831] modPreviewForm.ShowPreviewForm EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] modPreviewForm.ShowPreviewForm: " & Err.Number & " " & Err.Description
End Sub
```