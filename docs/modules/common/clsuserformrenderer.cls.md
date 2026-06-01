---
title: clsUserFormRenderer.cls
description: clsUserFormRenderer.cls のソースコード（コピペ用）
---

# clsUserFormRenderer.cls

**配置先**: `共通モジュール (3 ブック全て)` 用の VBA モジュール  
**種類**: クラス モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\`
- ファイル名: `clsUserFormRenderer.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

---

## ソースコード

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUserFormRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' Class: clsUserFormRenderer (v2.3 Phase N-3, 2026-05-27)
' Purpose:
'   Runtime dynamic UserForm generation for M-05 (register) / M-06 (edit)
'   / M-09 (view). Per ADR-0073 the form layout is generated at runtime
'   via VBComponents.Add(vbext_ct_MSForm) + Designer.Controls.Add (no
'   static .frm/.frx because the distribution is text-only per ADR-0008
'   and ADR-0013). The IScreenRenderer interface is kept as a stub so
'   existing sheet-renderer callers do not break.
' Dependencies:
'   modFormatLoader (format definition reader)
'   modKnowledgeFileIO (knowledge file read/write/delete)
'   modConfigHolder (data_dir/format_dir lookup)
' Field-control mapping (design book "UserForm ???I???? ?d?l (???f)" A3):
'   ?P??s -> Forms.TextBox.1 (MultiLine=False)
'   ?????s -> Forms.TextBox.1 (MultiLine=True, ScrollBars=fmScrollBarsVertical)
'   ????   -> Forms.TextBox.1 (single line + date validation in change event)
'   ?I??   -> Forms.ComboBox.1 (Style=fmStyleDropDownList)
' Layout (design book A4):
'   formWidth = 540 (譯�A險よｭ｣ mock 720px=540pt), label width = 120, control width = formWidth - 140
'   headerHeight = 48, rowHeight ?P??s/????/?I?? = 24, rowHeight ?????s = 60
'   buttonBarHeight = 40, bottomMargin = 16
' ================================================================
Implements IScreenRenderer
Option Explicit

' --- Layout DEFAULTS (ﾂｧA4 譯�A險よｭ｣蠕後�ｻmock 貅匁侠 540pt/720px; overridable via [USERFORM] stanza) ---
' ﾂｧA4 螳溯｣� SSOT: formWidth=540(pt; mock 720px 貅匁侠 譯�A 2026-05-28), labelWidth=120, headerHeight=48, rowHeight(蜊倅ｸ陦�/譌･譎�/驕ｸ謚�)=24,
' rowHeight(隍�謨ｰ陦�)=60, margin=10, buttonBarHeight=40, button 鬮倥＆=24/髢馴囈=8, bottomMargin=16.
' All values can be overridden per-screen via ui_seed/<role>/M-XX.txt [USERFORM] section.
Private Const DEFAULT_FORM_WIDTH As Long = 486
Private Const DEFAULT_FORM_HEIGHT As Long = 0                 ' 0 = auto-compute from rows
Private Const DEFAULT_HEADER_HEIGHT As Long = 48
Private Const DEFAULT_LABEL_WIDTH As Long = 107
Private Const DEFAULT_MARGIN As Long = 18
Private Const DEFAULT_BADGE_WIDTH As Long = 29
Private Const DEFAULT_BADGE_HEIGHT As Long = 16
Private Const DEFAULT_BADGE_GAP As Long = 3
' Phase R-3-ﾏ�-3 (2026-05-28): 邵ｦ荳ｦ縺ｳ layout (label 陦御ｸ� + data 陦悟�ｨ蟷�荳�)縲�
' row pitch(total) = labelZone(18) + dataHeight + rowSpacing(6)縲�
'   single 48 = 18 + 24 + 6 / multi 114 = 18 + 90 + 6 / multiLong 129 = 18 + 105 + 6
Private Const DEFAULT_ROW_HEIGHT_SINGLE As Long = 48
Private Const DEFAULT_ROW_HEIGHT_MULTI As Long = 114
Private Const DEFAULT_ROW_HEIGHT_MULTI_LONG As Long = 129
' 邵ｦ荳ｦ縺ｳ layout 螳壽焚 (label control 鬮� / label-data 髢� gap / data 荳� spacing)縲�
' label zone = VLABEL_H + VLABEL_GAP = 18縲ＥataHeight = rowPitch - 18 - VROW_SPACING縲�
Private Const VLABEL_H As Long = 16
Private Const VLABEL_GAP As Long = 2
Private Const VROW_SPACING As Long = 6
Private Const DEFAULT_SUBHEADER_HEIGHT As Long = 28
Private Const DEFAULT_BUTTON_BAR_HEIGHT As Long = 48
Private Const DEFAULT_BUTTON_WIDTH As Long = 113
Private Const DEFAULT_BUTTON_HEIGHT As Long = 26
Private Const DEFAULT_BUTTON_GAP As Long = 11
Private Const DEFAULT_BOTTOM_MARGIN As Long = 16
Private Const DEFAULT_FONT_NAME As String = "繝｡繧､繝ｪ繧ｪ"
Private Const DEFAULT_FONT_SIZE As Long = 10
Private Const DEFAULT_SUBHEADER_FONT_SIZE As Long = 12
Private Const DEFAULT_MULTILINE_SCROLLBARS As String = "vertical"
' Phase R-1-j: form surface color. Mock M-05/M-06/M-09 use a white form
' interior; VBA UserForm default is OS button-face grey. Overridable via
' [USERFORM].backColor (RRGGBB hex). Empty/blank = leave OS default.
Private Const DEFAULT_BACK_COLOR As String = "FFFFFF"
' Phase R-2 F-1: comma-separated header field ids (e.g. "knowledgeId,createdAt")
' rendered as readonly label+textbox rows between the format row and the
' subheader. Empty = none (R-1-j behaviour).
Private Const DEFAULT_HEADER_FIELDS As String = ""
' Phase R-2 F-2: field help line styling.
Private Const HELP_TEXT_COLOR As Long = &H808080      ' RGB(128,128,128) grey
Private Const HELP_FONT_SIZE As Long = 9
Private Const HELP_LINE_HEIGHT As Long = 15
' Phase R-2 F-4: placeholder gray text styling.
Private Const PLACEHOLDER_COLOR As Long = &H808080    ' RGB(128,128,128) grey
Private Const PLACEHOLDER_TAG As String = "PLACEHOLDER"
Private Const BADGE_BACK_COLOR As Long = &HE6E6FA
Private Const BADGE_TEXT_COLOR As Long = &H3C3CC8
Private Const SUBHEADER_COLOR As Long = &H643819
Private Const PRIMARY_BACK_COLOR As Long = &H643819
Private Const PRIMARY_FORE_COLOR As Long = &HFFFFFF
Private Const DESTRUCTIVE_BACK_COLOR As Long = &H3C3CDC
' V3 fix (2026-05-29): mock title bar
Private Const TITLE_BAR_COLOR As Long = &H643819
Private Const TITLE_BAR_FORE_COLOR As Long = &HFFFFFF
Private Const TITLE_BAR_HEIGHT As Long = 30
Private Const TITLE_BAR_FONT_SIZE As Long = 12

' --- Resolved layout values (set in InitFormConfig from defaults + [USERFORM] stanza override) ---
Private m_formWidth As Long
Private m_formHeight As Long
Private m_headerHeight As Long
Private m_labelWidth As Long
Private m_margin As Long
Private m_badgeWidth As Long
Private m_badgeHeight As Long
Private m_badgeGap As Long
Private m_rowHeightSingle As Long
Private m_rowHeightMulti As Long
Private m_rowHeightMultiLong As Long
Private m_subheaderHeight As Long
Private m_buttonBarHeight As Long
Private m_buttonWidth As Long
Private m_buttonHeight As Long
Private m_buttonGap As Long
Private m_bottomMargin As Long
Private m_fontName As String
Private m_fontSize As Long
Private m_subheaderFontSize As Long
Private m_multiLineScrollBars As String   ' "vertical" or "none"
Private m_captionOverride As String       ' empty = use mode-derived title
Private m_backColor As String             ' RRGGBB hex; empty = OS default
Private m_headerFields As String          ' Phase R-2 F-1: comma-separated ids
Private m_formatSelectorType As String    ' Phase R-2 F-3: "textbox" | "dropdown"
Private m_formatRowEnabled As Boolean      ' Phase R-3-ﾏ�-5: format 陦瑚｡ｨ遉ｺ(default mode=register 縺ｮ縺ｿ縲「i_seed formatRowEnabled 縺ｧ荳頑嶌縺榊庄)
Private m_knowledgeData As Object         ' Phase R-2: header field value source
Private m_formatHelp As String            ' Phase R-2 F-2: help line under format selector
Private m_headerHelp As Object            ' Phase R-2 F-2: Dict id->help line text
Private m_reformatRequested As Boolean    ' Phase R-2 F-3: dropdown asked to re-render
Private m_knowledgeNoRow As String        ' R-3-a: off|editable|readonly
Private m_knowledgeNoLabel As String
Private m_loadButtonLabel As String
Private m_knowledgeNoHelp As String
Private m_headerYOffset As Long           ' R-3-a: y shift for knowledge-no row above format
Private m_loadRequested As Boolean        ' R-3-a: btnLoad asked to reload by number
Private m_buttons As String               ' R-3-e: [USERFORM] buttons override (csv)

' --- vbext_ct_MSForm constant (avoid Microsoft Visual Basic for Applications Extensibility ref) ---
Private Const VBEXT_CT_MSFORM As Long = 3

' --- ProgIDs (per A3 mapping) ---
Private Const PROGID_TEXTBOX As String = "Forms.TextBox.1"
Private Const PROGID_COMBOBOX As String = "Forms.ComboBox.1"
Private Const PROGID_LABEL As String = "Forms.Label.1"
Private Const PROGID_BUTTON As String = "Forms.CommandButton.1"
' Phase R-3-ﾏ�-4 (2026-05-28): 譯�C/譯�A scroll 蛹� 窶� 蝗ｺ螳� header + frScroll Frame縲�
Private Const PROGID_FRAME As String = "Forms.Frame.1"
' V5 fix (2026-05-30) BUG-1: 1080px 逕ｻ髱｢縺ｧ button bar 縺檎判髱｢螟悶↓蜃ｺ繧� regression 繧帝亟縺舌◆繧�
' SCROLL_FORM_CAP 繧� 900竊�720 縺ｫ荳九￡ (72pt/inch ﾃ� 10in = 720pt = ~960px縲・xcel chrome + taskbar 謗ｧ髯､蠕後�ｮ
' 螳牙�ｨ蝨�)縲りｶ�驕主��縺ｯ frScroll 蜀�繧ｹ繧ｯ繝ｭ繝ｼ繝ｫ縺ｧ蜷ｸ蜿弱☆繧九�
' V4 fix (2026-05-29) #3 螻･豁ｴ: M-09縲悟次蝗縲埼國繧悟屓驕ｿ縺ｧ 900pt 縺ｫ諡｡蠑ｵ縺励◆縺後�1080px screen 縺ｧ
' button bar 逕ｻ髱｢螟紋ｺ区腐縺悟�ｺ縺溘◆繧∝�咲ｸｮ蟆上る聞 field 縺ｯ frame scroll 縺ｧ隕九○繧九�
Private Const SCROLL_FORM_CAP As Long = 720   ' form inside 鬮倥�ｮ荳企剞(pt)縲りｶ�驕主��縺ｯ frame 蜀�繧ｹ繧ｯ繝ｭ繝ｼ繝ｫ縲�

' --- Module-level state (single-instance UserForm session) ---
Private m_returnId As String          ' return value from form: knowledgeId / "DELETED" / ""
Private m_xlsmName As String          ' xlsm name (kanri / touroku / etc)
Private m_mode As String              ' "register" / "edit" / "view"
Private m_formatId As String          ' selected format id
Private m_knowledgeId As String       ' current knowledge id
Private m_readOnlyFormat As Boolean   ' format selection locked (edit/view)
Private m_fieldCount As Long
Private m_dynFormName As String

' Phase R-3-ﾏ�-4: scroll 蛹� 窶� 蝗ｺ螳� header 縺ｮ荳九↓鄂ｮ縺� frScroll Frame 縺ｨ蟇ｸ豕輔�
Private m_scrollFrame As Object       ' frScroll Frame (field 鄒､縺ｮ container)
Private m_headerHeightPx As Long      ' 蝗ｺ螳� header 鬮�(format selector + button bar)
Private m_scrollHeightPx As Long      ' frame 蜀� content total 鬮� (= Frame.ScrollHeight)

' Phase R-1-j (2026-05-28): path-verify dump output path. When non-empty,
' BuildAndShow performs the build phase, dumps designer.Controls to this
' file, removes the VBComponent, and returns WITHOUT calling Show.
Private m_dumpToFile As String

' Phase R-1-j (2026-05-28): modeless show flag. When True, BuildAndShow
' calls ufInstance.Show vbModeless (non-blocking) and stashes the form +
' vbc references so the caller can screenshot then close via
' CloseModelessForm. When False, normal vbModal + immediate dispose.
Private m_modelessRequest As Boolean
Private m_modelessFormInstance As Object
Private m_modelessVbc As Object

' Phase O-2 (2026-05-27): field index -> fieldName map. We name controls
' as ctl_<NNN> instead of ctl_<safeName(fieldName)> so multi-byte field
' names like  / cannot collide, and the persistence layer
' looks up the fieldName via this dictionary.
Private m_fieldNamesByIdx As Object       ' Dictionary: ctlName -> fieldName
Private m_dateFieldIndices As Object       ' Dictionary: idxStr -> True for ???? fields
' Phase P fix (2026-05-27): ComboBox items added on Designer.Controls.Add
' do not persist to the runtime UserForm instance returned by
' VBA.UserForms.Add. Stash items per ctlName during AddFieldRow, then
' populate the live instance after UserForms.Add (before Show).
Private m_comboItemsByCtl As Object       ' Dictionary: ctlName -> Array of items
Private m_comboInitialByCtl As Object     ' Dictionary: ctlName -> initial selected value
Private m_placeholderByCtl As Object      ' Phase R-2 F-4: ctlName -> placeholder text

' --- Public API (Phase N-3, ADR-0073 ??3.1) ---
' xlsmName: "???" / "?o?^?C??" / "????" (used for log scope only; form is workbook-local)
' mode    : "register" | "edit" | "view" | "preview"
' knowledgeId: edit/view only (register/preview pass "")
' readOnlyFormat: TRUE pins format selection (edit/view/preview). FALSE allows change (register).
' formatId: preview (M-04) only 窶� fixed format to preview; data is not loaded.
Public Function ShowForm( _
        ByVal xlsmName As String, _
        ByVal mode As String, _
        ByVal knowledgeId As String, _
        Optional ByVal readOnlyFormat As Boolean = False, _
        Optional ByVal formatId As String = "" _
) As String

    ' --- VBE flicker suppression (Phase P revised, 2026-05-27) ---
    ' DIAGNOSTIC: Phase P found Show vbModal failing silently when
    ' Application.VBE.MainWindow.Visible was set to False before Show.
    ' Disable the toggle to keep Show functional; VBE may flicker briefly.
    ' To re-enable suppression once the root cause is fixed, restore the
    ' toggle here.
    Dim wasVbeVisible As Boolean
    Dim vbeAvail As Boolean
    vbeAvail = False

    m_xlsmName = xlsmName
    m_mode = LCase$(mode)
    m_knowledgeId = knowledgeId
    m_readOnlyFormat = readOnlyFormat
    m_returnId = ""

    ' Register this instance with the callback bridge so dynamic-form
    ' button handlers (Application.Run "modUserFormCallback.OnXxx") can
    ' communicate the return id back to ShowForm.
    modUserFormCallback.SetRenderer Me

    ' Resolve formatId: register=blank, edit/view=load from knowledge file
    Dim knowledgeData As Object
    If Len(knowledgeId) > 0 Then
        Set knowledgeData = modKnowledgeFileIO.LoadKnowledge(knowledgeId)
        If Not knowledgeData Is Nothing Then
            If knowledgeData.Exists("FormatID") Then
                m_formatId = CStr(knowledgeData("FormatID"))
            End If
        End If
    Else
        Set knowledgeData = Nothing
    End If

    ' Phase R-3-ﾏ�-2: preview mode (M-04) 窶� fixed format from caller, no data load.
    ' M-09 縺ｨ蜷� interface (陦ｨ遉ｺ蟆ら畑 popup) 縺ｧ縲∝ｷｮ縺ｯ縲悟ｮ溘ョ繝ｼ繧ｿ繧� load 縺励↑縺�縲咲せ縺ｮ縺ｿ縲�
    If m_mode = "preview" And Len(formatId) > 0 Then m_formatId = formatId

    ' Build dynamic form, generate controls, show, dispose. Result in m_returnId.
    ' Phase R-2 F-3: loop when the format dropdown requests a re-render with a
    ' new format (m_reformatRequested set by modUserFormCallback.OnFormatChange).
    Do
        m_reformatRequested = False
        ' R-3-a: btnLoad requested a reload by number -> refetch data + format.
        If m_loadRequested Then
            m_loadRequested = False
            If Len(m_knowledgeId) > 0 Then
                Set knowledgeData = modKnowledgeFileIO.LoadKnowledge(m_knowledgeId)
                If Not knowledgeData Is Nothing Then
                    If knowledgeData.Exists("FormatID") Then m_formatId = CStr(knowledgeData("FormatID"))
                End If
            End If
        End If
        BuildAndShow knowledgeData
    Loop While m_reformatRequested

    ShowForm = m_returnId
    GoTo Cleanup

ErrHandler:
    Debug.Print "[ERR] clsUserFormRenderer.ShowForm: " & Err.Number & " " & Err.Description
    ShowForm = ""

Cleanup:
    ' Restore VBE state. Wrap in Resume Next so a single failure does not
    ' leave the workbook in an unusable state for the user.
    On Error Resume Next
    If vbeAvail Then Application.VBE.MainWindow.Visible = wasVbeVisible
    On Error GoTo 0
End Function

' Phase R-1-f: Set m_* config to defaults, then override from ui_seed [USERFORM] stanza.
Private Sub InitFormConfig()
    m_formWidth = DEFAULT_FORM_WIDTH
    m_formHeight = DEFAULT_FORM_HEIGHT
    m_headerHeight = DEFAULT_HEADER_HEIGHT
    m_labelWidth = DEFAULT_LABEL_WIDTH
    m_margin = DEFAULT_MARGIN
    m_badgeWidth = DEFAULT_BADGE_WIDTH
    m_badgeHeight = DEFAULT_BADGE_HEIGHT
    m_badgeGap = DEFAULT_BADGE_GAP
    m_rowHeightSingle = DEFAULT_ROW_HEIGHT_SINGLE
    m_rowHeightMulti = DEFAULT_ROW_HEIGHT_MULTI
    m_rowHeightMultiLong = DEFAULT_ROW_HEIGHT_MULTI_LONG
    m_subheaderHeight = DEFAULT_SUBHEADER_HEIGHT
    m_buttonBarHeight = DEFAULT_BUTTON_BAR_HEIGHT
    m_buttonWidth = DEFAULT_BUTTON_WIDTH
    m_buttonHeight = DEFAULT_BUTTON_HEIGHT
    m_buttonGap = DEFAULT_BUTTON_GAP
    m_bottomMargin = DEFAULT_BOTTOM_MARGIN
    m_fontName = DEFAULT_FONT_NAME
    m_fontSize = DEFAULT_FONT_SIZE
    m_subheaderFontSize = DEFAULT_SUBHEADER_FONT_SIZE
    m_multiLineScrollBars = DEFAULT_MULTILINE_SCROLLBARS
    m_captionOverride = ""
    m_backColor = DEFAULT_BACK_COLOR
    m_headerFields = DEFAULT_HEADER_FIELDS
    m_formatSelectorType = "textbox"
    ' Phase R-3-ﾏ�-5: format 陦後�ｯ譌｢螳壹〒 register 縺ｮ縺ｿ陦ｨ遉ｺ(M-04/M-06/M-09 縺ｯ譛ｬ菴薙°繧蛾勁螟�)縲�
    m_formatRowEnabled = (m_mode = "register")
    m_formatHelp = ""
    Set m_headerHelp = CreateObject("Scripting.Dictionary")
    m_knowledgeNoRow = "off"
    m_knowledgeNoLabel = ""
    m_loadButtonLabel = ""
    m_knowledgeNoHelp = ""
    m_headerYOffset = 0
    m_buttons = ""
End Sub

' Parse an RRGGBB hex string to a VBA color Long (&HBBGGRR). Returns -1 if
' the string is empty or unparseable (caller leaves the OS default).
Private Function ParseHexColor(ByVal rrggbb As String) As Long
    On Error GoTo Fail
    Dim s As String
    s = Trim$(rrggbb)
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)
    If Len(s) <> 6 Then GoTo Fail
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid$(s, 1, 2))
    g = CLng("&H" & Mid$(s, 3, 2))
    b = CLng("&H" & Mid$(s, 5, 2))
    ParseHexColor = RGB(r, g, b)
    Exit Function
Fail:
    ParseHexColor = -1
End Function

' Map (xlsmName, mode) -> screenId for ui_seed lookup (Phase R-1-f).
Private Function ResolveScreenId() As String
    Select Case m_mode
        Case "register": ResolveScreenId = "M-05"
        Case "edit":     ResolveScreenId = "M-06"
        Case "view":     ResolveScreenId = "M-09"
        Case "preview":  ResolveScreenId = "M-04"
        Case Else:       ResolveScreenId = ""
    End Select
End Function

' Phase R-1-f: load ui_seed/<role>/<screenId>.txt, find [USERFORM] section, override m_*.
Private Sub ApplyUserformStanza()
    On Error Resume Next
    Dim screenId As String
    screenId = ResolveScreenId()
    If Len(screenId) = 0 Then Exit Sub
    Dim sections As Collection
    Set sections = modUILoader.LoadUiDefinition(m_xlsmName, screenId)
    If sections Is Nothing Then Exit Sub
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To sections.count
        Set sec = sections.Item(i)
        If sec.sectionName = "USERFORM" Then
            ApplyConfigFromStanza sec
            Exit For
        End If
    Next i
    On Error GoTo 0
End Sub

' Phase R-1-h: format-level [USERFORM] override (highest priority, 3-stage fallback).
' format > sheet > default. Each key independent (per-key fallback).
Private Sub ApplyUserformStanzaFromFormat()
    On Error Resume Next
    If Len(m_formatId) = 0 Then Exit Sub
    Dim sections As Collection
    Set sections = modFormatLoader.LoadFormat(m_formatId)
    If sections Is Nothing Then Exit Sub
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To sections.count
        Set sec = sections.Item(i)
        If sec.sectionName = "USERFORM" Then
            ApplyConfigFromStanza sec
            Exit For
        End If
    Next i
    On Error GoTo 0
End Sub

' Per-key override (only set if key present + parseable). Phase R-1-f.
Private Sub ApplyConfigFromStanza(ByVal sec As ClsStanzaSection)
    Dim v As String
    v = Trim$(sec.GetValue("formWidth"))
    If Len(v) > 0 And IsNumeric(v) Then m_formWidth = CLng(v)
    v = Trim$(sec.GetValue("formHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_formHeight = CLng(v)
    v = Trim$(sec.GetValue("headerHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_headerHeight = CLng(v)
    v = Trim$(sec.GetValue("labelWidth"))
    If Len(v) > 0 And IsNumeric(v) Then m_labelWidth = CLng(v)
    v = Trim$(sec.GetValue("margin"))
    If Len(v) > 0 And IsNumeric(v) Then m_margin = CLng(v)
    v = Trim$(sec.GetValue("rowHeightSingle"))
    If Len(v) > 0 And IsNumeric(v) Then m_rowHeightSingle = CLng(v)
    v = Trim$(sec.GetValue("rowHeightMulti"))
    If Len(v) > 0 And IsNumeric(v) Then m_rowHeightMulti = CLng(v)
    v = Trim$(sec.GetValue("rowHeightMultiLong"))
    If Len(v) > 0 And IsNumeric(v) Then m_rowHeightMultiLong = CLng(v)
    v = Trim$(sec.GetValue("subheaderHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_subheaderHeight = CLng(v)
    v = Trim$(sec.GetValue("buttonBarHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_buttonBarHeight = CLng(v)
    v = Trim$(sec.GetValue("buttonWidth"))
    If Len(v) > 0 And IsNumeric(v) Then m_buttonWidth = CLng(v)
    v = Trim$(sec.GetValue("buttonHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_buttonHeight = CLng(v)
    v = Trim$(sec.GetValue("buttonGap"))
    If Len(v) > 0 And IsNumeric(v) Then m_buttonGap = CLng(v)
    v = Trim$(sec.GetValue("bottomMargin"))
    If Len(v) > 0 And IsNumeric(v) Then m_bottomMargin = CLng(v)
    v = Trim$(sec.GetValue("badgeWidth"))
    If Len(v) > 0 And IsNumeric(v) Then m_badgeWidth = CLng(v)
    v = Trim$(sec.GetValue("badgeHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_badgeHeight = CLng(v)
    v = Trim$(sec.GetValue("fontSize"))
    If Len(v) > 0 And IsNumeric(v) Then m_fontSize = CLng(v)
    v = Trim$(sec.GetValue("subheaderFontSize"))
    If Len(v) > 0 And IsNumeric(v) Then m_subheaderFontSize = CLng(v)
    v = Trim$(sec.GetValue("fontName"))
    If Len(v) > 0 Then m_fontName = v
    v = Trim$(sec.GetValue("multiLineScrollBars"))
    If Len(v) > 0 Then m_multiLineScrollBars = LCase$(v)
    v = Trim$(sec.GetValue("caption"))
    If Len(v) > 0 Then m_captionOverride = v
    v = Trim$(sec.GetValue("backColor"))
    If Len(v) > 0 Then m_backColor = v
    v = Trim$(sec.GetValue("headerFields"))
    If Len(v) > 0 Then m_headerFields = v
    v = Trim$(sec.GetValue("formatSelectorType"))
    If Len(v) > 0 Then m_formatSelectorType = LCase$(v)
    ' Phase R-3-ﾏ�-5: format 陦後�ｮ陦ｨ遉ｺ蜿ｯ蜷ｦ繧� ui_seed 縺ｧ譏守､ｺ荳頑嶌縺�(M-04/06/09=false)縲�
    v = LCase$(Trim$(sec.GetValue("formatRowEnabled")))
    If Len(v) > 0 Then m_formatRowEnabled = (v = "true" Or v = "1" Or v = "yes")
    v = Trim$(sec.GetValue("formatHelp"))
    If Len(v) > 0 Then m_formatHelp = v
    v = Trim$(sec.GetValue("knowledgeNoRow"))
    If Len(v) > 0 Then m_knowledgeNoRow = LCase$(v)
    v = Trim$(sec.GetValue("knowledgeNoLabel"))
    If Len(v) > 0 Then m_knowledgeNoLabel = v
    v = Trim$(sec.GetValue("loadButtonLabel"))
    If Len(v) > 0 Then m_loadButtonLabel = v
    v = Trim$(sec.GetValue("knowledgeNoHelp"))
    If Len(v) > 0 Then m_knowledgeNoHelp = v
    v = Trim$(sec.GetValue("buttons"))
    If Len(v) > 0 Then m_buttons = v
    ' Phase R-2 F-2: per-header-field help lines (headerHelp_<id>).
    If Len(Trim$(m_headerFields)) > 0 Then
        Dim hids() As String
        hids = Split(m_headerFields, ",")
        Dim hk As Long
        For hk = LBound(hids) To UBound(hids)
            Dim hid As String
            hid = Trim$(hids(hk))
            If Len(hid) > 0 Then
                Dim hv As String
                hv = Trim$(sec.GetValue("headerHelp_" & hid))
                If Len(hv) > 0 Then m_headerHelp(hid) = hv
            End If
        Next hk
    End If
End Sub

' --- Internal: build form, populate, modal show, then dispose ---
Private Sub BuildAndShow(ByVal knowledgeData As Object)
    On Error GoTo ErrHandler

    ' Phase R-1-f/h: 3-stage fallback config load. default 竊・sheet 竊・format.
    InitFormConfig
    ApplyUserformStanza             ' sheet-level override
    ApplyUserformStanzaFromFormat   ' format-level override (highest priority)

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    If vbProj Is Nothing Then
        Debug.Print "[ERR] VBProject not available"
        Exit Sub
    End If

    LogToSheet "BuildAndShow", "step 1 VBComponents.Add", "LOG-UF-STEP-01"
    Dim vbc As Object
    Set vbc = vbProj.VBComponents.Add(VBEXT_CT_MSFORM)
    m_dynFormName = "frmDyn_" & Format$(Now(), "yyyymmddhhnnss") & "_" & vbc.Name
    vbc.Name = m_dynFormName
    LogToSheet "BuildAndShow", "step 2 named " & m_dynFormName, "LOG-UF-STEP-02"

    ' Phase O-2: reset per-form maps (control-name -> field-name; ???? set)
    Set m_fieldNamesByIdx = CreateObject("Scripting.Dictionary")
    Set m_dateFieldIndices = CreateObject("Scripting.Dictionary")
    Set m_comboItemsByCtl = CreateObject("Scripting.Dictionary")
    Set m_comboInitialByCtl = CreateObject("Scripting.Dictionary")
    Set m_placeholderByCtl = CreateObject("Scripting.Dictionary")
    LogToSheet "BuildAndShow", "step 3 dicts created", "LOG-UF-STEP-03"

    Dim designer As Object
    Set designer = vbc.designer
    LogToSheet "BuildAndShow", "step 4 designer ok", "LOG-UF-STEP-04"

    ' Resolve format fields. Even register mode needs format choice; if formatId blank
    ' we leave control area empty until user picks one in the dropdown.
    Dim fmtFields As Collection
    Set fmtFields = LoadFormatFields(m_formatId)
    m_fieldCount = 0
    If Not fmtFields Is Nothing Then m_fieldCount = fmtFields.count
    LogToSheet "BuildAndShow", "step 5 fields=" & m_fieldCount & " fmtId=" & m_formatId, "LOG-UF-STEP-05"

    ' Phase R-2: stash knowledge data for header-field value lookup.
    Set m_knowledgeData = knowledgeData

    ' Phase P fix (2026-05-27): UserForm size/caption properties live on the
    ' VBComponent.Properties collection, NOT on Designer.Properties. Width and
    ' caption can be set now; Height is set AFTER rendering (Phase R-2: help
    ' lines make row heights variable, so we size the form to the actual layout).
    vbc.Properties("Caption") = ResolveFormTitle()
    vbc.Properties("Width") = m_formWidth
    ' Phase R-1-j: form surface color from [USERFORM].backColor (RRGGBB).
    Dim bc As Long
    bc = ParseHexColor(m_backColor)
    If bc >= 0 Then
        On Error Resume Next
        vbc.Properties("BackColor") = bc
        On Error GoTo ErrHandler
    End If

    ' V3 fix (2026-05-29): titlebar(navy) + format selector header. button bar moved to bottom.
    AddTitleBar designer
    m_headerYOffset = TITLE_BAR_HEIGHT
    Dim hdrFmtBottom As Long
    If m_formatRowEnabled Then
        hdrFmtBottom = AddHeaderRow(designer)
    Else
        hdrFmtBottom = TITLE_BAR_HEIGHT + m_margin
    End If
    m_headerHeightPx = hdrFmtBottom
    LogToSheet "BuildAndShow", "step 7 fixed header(title+format) headerH=" & m_headerHeightPx, "LOG-UF-STEP-07"

    ' chrome 螳滓ｸｬ縺ｮ縺溘ａ Width/Height 繧呈圻螳夊ｨｭ螳�
    Dim chromeW As Long, chromeH As Long
    chromeW = 0: chromeH = 0
    vbc.Properties("Width") = m_formWidth
    vbc.Properties("Height") = SCROLL_FORM_CAP
    On Error Resume Next
    chromeW = CLng(vbc.Properties("Width").Value) - CLng(vbc.Properties("InsideWidth").Value)
    chromeH = CLng(vbc.Properties("Height").Value) - CLng(vbc.Properties("InsideHeight").Value)
    On Error GoTo ErrHandler
    If chromeW < 1 Then chromeW = 12
    If chromeH < 1 Then chromeH = 29

    ' scroll area Frame (header 縺ｮ荳�)
    Dim fr As Object
    Set fr = designer.Controls.Add(PROGID_FRAME, "frScroll", True)
    fr.caption = ""
    fr.top = m_headerHeightPx
    fr.left = 0
    fr.Width = m_formWidth
    ' V4 fix (2026-05-29) #1/#5 (revised): Frame 縺ｮ sunken/etched 譫 + 轣ｰ濶ｲ閭梧勹繧呈ｶ医☆縲�
    ' V4-iter1 縺ｧ BorderStyle=fmBorderStyleNone 繧定ｿｽ蜉縺励◆繧� M-05 register 縺�
    ' 1 field 縺励°謠冗判縺輔ｌ縺ｪ縺� regression 縺悟�ｺ縺溘◆繧� (frame 縺ｮ inset+clipping 莉墓ｧ伜､牙喧)縲�
    ' BorderStyle 縺ｯ default (0=fmBorderStyleNone 縺ｯ螟峨∴縺�) + SpecialEffect=0(flat)
    ' + BackColor=逋ｽ 縺ｧ縲梧棧隕九◆逶ｮ縺縺代肴椛縺医ｋ譁ｹ蜷代↓謖ｯ繧九�
    On Error Resume Next
    fr.SpecialEffect = 0       ' fmSpecialEffectFlat (etched 譫繧呈ｶ医☆)
    Dim bcFrame As Long
    bcFrame = ParseHexColor(m_backColor)
    If bcFrame >= 0 Then fr.BackColor = bcFrame
    On Error GoTo ErrHandler
    Set m_scrollFrame = fr

    ' scroll content (frame-relative 蠎ｧ讓吶〒 frScroll 蜀�縺ｫ逕滓��)
    Dim y As Long
    Dim kOff As Long
    kOff = AddKnowledgeNoRow(fr)                        ' 繝翫Ξ繝�繧ｸ逡ｪ蜿ｷ陦�(edit/view); M-05=0
    y = m_margin + kOff
    y = AddHeaderFields(fr, y)                          ' 莠亥ｮ夂分蜿ｷ 遲� header fields
    AddSubheaderRow fr, y
    y = y + m_subheaderHeight
    Dim i As Long
    If Not fmtFields Is Nothing Then
        For i = 1 To fmtFields.count
            y = y + AddFieldRow(fr, fmtFields.Item(i), y, knowledgeData)
        Next i
    End If
    Dim contentBottom As Long
    contentBottom = y + m_bottomMargin
    m_scrollHeightPx = contentBottom
    LogToSheet "BuildAndShow", "step 8 scroll content bottom=" & contentBottom, "LOG-UF-STEP-08"

    ' 蝗ｺ螳� form inside 鬮�: 蜀�螳ｹ縺悟庶縺ｾ繧後�ｰ縺昴�ｮ縺ｾ縺ｾ縲∬ｶ�縺医◆繧� cap 縺ｧ frame 蜀�繧ｹ繧ｯ繝ｭ繝ｼ繝ｫ
    Dim formInsideH As Long
    Dim fullNeeded As Long
    fullNeeded = m_headerHeightPx + contentBottom + m_buttonBarHeight + 8
    If m_formHeight > 0 Then
        formInsideH = m_formHeight
    ElseIf fullNeeded > SCROLL_FORM_CAP Then
        formInsideH = SCROLL_FORM_CAP
    Else
        formInsideH = fullNeeded
    End If
    ' V3 fix: frame height reserves bottom button bar zone.
    fr.Height = formInsideH - m_headerHeightPx - m_buttonBarHeight - 8
    fr.ScrollBars = 2
    fr.ScrollHeight = contentBottom
    On Error Resume Next
    fr.ScrollWidth = m_formWidth - 2
    On Error GoTo ErrHandler

    ' V3 fix: button bar fixed at form bottom.
    AddButtonBar designer, formInsideH - m_buttonBarHeight - 4

    Dim formH As Long
    formH = formInsideH
    vbc.Properties("Width") = m_formWidth + chromeW
    vbc.Properties("Height") = formInsideH + chromeH
    LogToSheet "BuildAndShow", "step 6 props insideH=" & formInsideH & " frameH=" & fr.Height & " scrollH=" & contentBottom, "LOG-UF-STEP-06"
    LogToSheet "BuildAndShow", "step 9 header buttons mode=" & m_mode, "LOG-UF-STEP-09"

    ' Inject minimal init/handler code into the form's code module so that
    ' clicking a button can set m_returnId and call UserForm.Hide.
    InjectFormCode vbc
    LogToSheet "BuildAndShow", "step 10 code injected", "LOG-UF-STEP-10"

    ' Phase R-1-j path-verify: if dump path is set, dump designer state and
    ' return without Show. Caller (TestR1j_DumpForm) inspects the file.
    If Len(m_dumpToFile) > 0 Then
        DumpFormToFile vbc, designer, fmtFields, formH
        On Error Resume Next
        vbProj.VBComponents.Remove vbc
        On Error GoTo ErrHandler
        Exit Sub
    End If

    ' Phase P fix: get the live form instance, populate ComboBox items at
    ' runtime (Designer.Controls.Add items don't survive), then Show.
    LogToSheet "BuildAndShow", "step 11 about to Show formName=" & m_dynFormName, "LOG-UF-SHOW-PRE"
    Dim ufInstance As Object
    Set ufInstance = VBA.UserForms.Add(m_dynFormName)
    PopulateComboBoxesOnInstance ufInstance

    ' V4 fix (2026-05-29) #M-05: register 繝｢繝ｼ繝峨〒 frScroll 蜀�縺ｮ field 鄒､縺�
    ' 1 莉ｶ縺励° paint 縺輔ｌ縺ｪ縺� regression 縺檎匱逕� (蟄� agent dump 縺ｧ 11 莉ｶ蜈ｨ讒狗ｯ画ｸ�
    ' 遒ｺ隱�)縲Ｇrame 縺ｮ ScrollTop 繧� 0 縺ｫ reset + Repaint nudge 縺ｧ蛻晄悄謠冗判繧貞ｼｷ蛻ｶ縲�
    On Error Resume Next
    Dim frLive As Object
    Set frLive = ufInstance.Controls("frScroll")
    If Not frLive Is Nothing Then
        frLive.ScrollTop = 0
        frLive.Repaint
    End If
    On Error GoTo ErrHandler

    ' Phase R-1-j: modeless path for screenshot harness. Caller closes via
    ' CloseModelessForm to dispose the VBComponent + form instance.
    If m_modelessRequest Then
        Set m_modelessFormInstance = ufInstance
        Set m_modelessVbc = vbc
        ufInstance.Show vbModeless
        Exit Sub
    End If

    ufInstance.Show vbModal
    LogToSheet "BuildAndShow", "step 12 closed returnId=" & m_returnId, "LOG-UF-SHOW-POST"

    ' Dispose dynamic form
    On Error Resume Next
    vbProj.VBComponents.Remove vbc
    On Error GoTo ErrHandler

    Exit Sub

ErrHandler:
    Debug.Print "[ERR] BuildAndShow: " & Err.Number & " " & Err.Description
    On Error Resume Next
    LogToSheet "BuildAndShow", "Show failed errNum=" & Err.Number & " desc=" & Err.Description & " formName=" & m_dynFormName, "LOG-UF-SHOW-ERR"
    On Error GoTo 0
End Sub

' Phase R-3-ﾏ�-4: find a control by name on the live form, searching direct
' controls then the frScroll frame's children (one level). field 邉ｻ縺ｯ frScroll 蜀�縲�
Private Function FindCtlOnForm(ByVal uf As Object, ByVal ctlName As String) As Object
    On Error Resume Next
    Dim c As Object
    Set c = uf.Controls(ctlName)
    If Not c Is Nothing Then
        Set FindCtlOnForm = c
        Exit Function
    End If
    Dim f As Object
    For Each f In uf.Controls
        If TypeName(f) = "Frame" Then
            Dim cc As Object
            Set cc = Nothing
            Set cc = f.Controls(ctlName)
            If Not cc Is Nothing Then
                Set FindCtlOnForm = cc
                Exit Function
            End If
        End If
    Next f
End Function

' Phase P fix: walk the live UserForm instance and add ComboBox items per
' the stashed m_comboItemsByCtl map. Then set initial selected value if any.
Private Sub PopulateComboBoxesOnInstance(ByVal uf As Object)
    On Error Resume Next
    If m_comboItemsByCtl Is Nothing Then Exit Sub
    Dim ctlName As Variant
    For Each ctlName In m_comboItemsByCtl.Keys
        Dim items As Variant
        items = m_comboItemsByCtl(CStr(ctlName))
        Dim cb As Object
        Set cb = FindCtlOnForm(uf, CStr(ctlName))   ' R-3-ﾏ�-4: frScroll frame 蜀�繧よ爾邏｢
        If Not cb Is Nothing Then
            Dim i As Long
            For i = LBound(items) To UBound(items)
                cb.AddItem CStr(items(i))
            Next i
            If m_comboInitialByCtl.Exists(CStr(ctlName)) Then
                cb.value = CStr(m_comboInitialByCtl(CStr(ctlName)))
            End If
        End If
    Next ctlName
    On Error GoTo 0
End Sub

' Diagnostic logger - write to the active workbook's LOG sheet via clsLogger.
Private Sub LogToSheet(ByVal funcName As String, ByVal msg As String, ByVal logId As String)
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets("LOG")
    If Not lg Is Nothing Then
        lg.LogInfo "clsUserFormRenderer", funcName, msg, "", logId
    End If
    On Error GoTo 0
End Sub

Private Function ResolveFormTitle() As String
    ' Phase R-1-f: stanza caption override wins over mode-derived JP title.
    If Len(m_captionOverride) > 0 Then
        ResolveFormTitle = m_captionOverride
        Exit Function
    End If
    Select Case m_mode
        Case "register": ResolveFormTitle = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H767B) & ChrW(&H9332)
        Case "edit":     ResolveFormTitle = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H4FEE) & ChrW(&H6B63)
        Case "view":     ResolveFormTitle = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H8868) & ChrW(&H793A)
        Case "preview"
            ' 繝励Ξ繝薙Η繝ｼ (V4 fix 2026-05-29: drop ": <formatId>" suffix per隕ｪ謖�遉ｺ#4)
            ResolveFormTitle = ChrW(&H30D7) & ChrW(&H30EC) & ChrW(&H30D3) & ChrW(&H30E5) & ChrW(&H30FC)
        Case Else:       ResolveFormTitle = "UserForm"
    End Select
End Function

Private Function LoadFormatFields(ByVal fmtId As String) As Collection
    On Error GoTo ErrHandler
    If Len(fmtId) = 0 Then
        Set LoadFormatFields = Nothing
        Exit Function
    End If
    Dim secs As Collection
    Set secs = modFormatLoader.LoadFormat(fmtId)
    If secs Is Nothing Then
        Set LoadFormatFields = Nothing
        Exit Function
    End If
    Dim out As Collection
    Set out = New Collection
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To secs.count
        Set sec = secs.Item(i)
        If sec.sectionName = "FIELD" Then
            out.Add sec
        End If
    Next i
    Set LoadFormatFields = out
    Exit Function
ErrHandler:
    Set LoadFormatFields = Nothing
End Function

Private Function ComputeRowsHeight(ByVal flds As Collection) As Long
    ' Includes subheader row "�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽi�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽb�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽW�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽﾌ難ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽe" + per-field row heights.
    Dim h As Long
    h = m_subheaderHeight
    If flds Is Nothing Then
        ComputeRowsHeight = h + m_rowHeightSingle
        Exit Function
    End If
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To flds.count
        Set sec = flds.Item(i)
        If sec.GetValue("FieldType") = ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C) Then  ' �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽs
            ' Long multi-line for known "long" fields (�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽﾆ手順, �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ, �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ); others use short
            If IsLongMultilineField(sec.GetValue("FieldName")) Then
                h = h + m_rowHeightMultiLong
            Else
                h = h + m_rowHeightMulti
            End If
        Else
            h = h + m_rowHeightSingle
        End If
    Next i
    ComputeRowsHeight = h
End Function

' Heuristic: identify field names that mock shows as tall multi-line boxes.
' These get m_rowHeightMultiLong; other multi-line fields get m_rowHeightMulti.
Private Function IsLongMultilineField(ByVal fieldName As String) As Boolean
    Dim s As String
    s = fieldName
    ' �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽﾆ手順 / �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ / �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ / �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽﾛ詳搾ｿｽ / �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽﾚ搾ｿｽ / �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽe �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽc
    If InStr(s, ChrW(&H4F5C) & ChrW(&H696D) & ChrW(&H624B) & ChrW(&H9806)) > 0 Then IsLongMultilineField = True: Exit Function
    If InStr(s, ChrW(&H4E8B) & ChrW(&H8C61)) > 0 Then IsLongMultilineField = True: Exit Function
    If InStr(s, ChrW(&H539F) & ChrW(&H56E0)) > 0 Then IsLongMultilineField = True: Exit Function
    If InStr(s, ChrW(&H8A73) & ChrW(&H7D30)) > 0 Then IsLongMultilineField = True: Exit Function
    If InStr(s, ChrW(&H5185) & ChrW(&H5BB9)) > 0 Then IsLongMultilineField = True: Exit Function
End Function

' Phase R-2 F-2: render a grey 9pt help line at (leftX, y) of given width.
' Returns the vertical space consumed (HELP_LINE_HEIGHT if text non-empty,
' else 0 so callers that skip empty help do not change row heights).
Private Function RenderHelpLine(ByVal designer As Object, ByVal ctlName As String, _
                                 ByVal y As Long, ByVal leftX As Long, ByVal w As Long, _
                                 ByVal text As String) As Long
    If Len(Trim$(text)) = 0 Then
        RenderHelpLine = 0
        Exit Function
    End If
    On Error Resume Next
    Dim hl As Object
    Set hl = designer.Controls.Add(PROGID_LABEL, ctlName, True)
    hl.caption = text
    hl.top = y
    hl.left = leftX
    hl.Width = w
    hl.Height = HELP_LINE_HEIGHT
    hl.Font.Name = m_fontName
    hl.Font.Size = HELP_FONT_SIZE
    hl.ForeColor = HELP_TEXT_COLOR
    On Error GoTo 0
    RenderHelpLine = HELP_LINE_HEIGHT
End Function

' R-3-a (2026-05-28): knowledge-number row above the format row.
'   editable (M-06) -> label + input textbox + load button + help (load existing)
'   readonly (M-09) -> label + readonly value textbox + help
'   off (M-05)      -> not rendered (returns 0). config-driven, no hardcode.
Private Function AddKnowledgeNoRow(ByVal designer As Object) As Long
    If m_knowledgeNoRow <> "editable" And m_knowledgeNoRow <> "readonly" Then
        AddKnowledgeNoRow = 0
        Exit Function
    End If
    ' Phase R-3-ﾏ�-3 邵ｦ荳ｦ縺ｳ: label 陦御ｸ翫�ｻdata 陦�(蜈ｨ蟷�)荳九�
    Dim y As Long
    y = m_margin
    Dim fullWK As Long, dataTopK As Long, dataHK As Long
    fullWK = m_formWidth - m_margin * 2
    dataTopK = y + VLABEL_H + VLABEL_GAP
    dataHK = m_rowHeightSingle - VLABEL_H - VLABEL_GAP - VROW_SPACING
    Dim lblK As Object
    Set lblK = designer.Controls.Add(PROGID_LABEL, "lblKnowledgeNo", True)
    lblK.caption = m_knowledgeNoLabel
    lblK.top = y
    lblK.left = m_margin
    lblK.Width = fullWK
    lblK.Height = VLABEL_H
    ApplyBaseFont lblK
    Dim ctlK As Object
    Set ctlK = designer.Controls.Add(PROGID_TEXTBOX, "txtKnowledgeNo", True)
    ctlK.top = dataTopK
    ctlK.left = m_margin
    ctlK.Height = dataHK
    ApplyBaseFont ctlK
    If m_knowledgeNoRow = "editable" And Len(m_loadButtonLabel) > 0 Then
        ctlK.Width = fullWK - m_buttonWidth - m_buttonGap
        Dim btnL As Object
        Set btnL = designer.Controls.Add(PROGID_BUTTON, "btnLoad", True)
        btnL.caption = m_loadButtonLabel
        btnL.top = dataTopK
        btnL.left = m_margin + (fullWK - m_buttonWidth)
        btnL.Width = m_buttonWidth
        btnL.Height = dataHK
        On Error Resume Next
        btnL.Font.Name = m_fontName
        btnL.Font.Size = m_fontSize
        On Error GoTo 0
    Else
        ctlK.Width = fullWK
        ctlK.Locked = True
        ctlK.BackColor = RGB(240, 240, 240)
    End If
    ' BUG-6 fix (2026-05-30): seed the visible textbox with the current
    ' knowledge id. Without this, M-09 (readonly) and edit-mode M-06
    ' (editable) reopen showed an empty input box even though the
    ' internal m_knowledgeId was set. The textbox is always populated
    ' when m_knowledgeId has a value; otherwise it stays blank.
    If Len(m_knowledgeId) > 0 Then
        On Error Resume Next
        ctlK.Value = m_knowledgeId
        On Error GoTo 0
    End If
    y = dataTopK + dataHK
    If Len(m_knowledgeNoHelp) > 0 Then
        y = y + 2 + RenderHelpLine(designer, "lblKnowledgeNoHelp", y + 2, m_margin, fullWK, m_knowledgeNoHelp)
    End If
    AddKnowledgeNoRow = (y - m_margin) + 6
End Function

' Returns the y position after the format row (+ optional format help line).
Private Function AddHeaderRow(ByVal designer As Object) As Long
    Dim lbl As Object
    Set lbl = designer.Controls.Add(PROGID_LABEL, "lblFormat", True)
    lbl.caption = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8)
    lbl.top = m_margin + m_headerYOffset
    lbl.left = m_margin
    lbl.Width = m_formWidth - m_margin * 2
    lbl.Height = VLABEL_H
    ApplyBaseFont lbl

    ' Phase R-3-ﾏ�-3 邵ｦ荳ｦ縺ｳ: format selector 繧� label 荳翫�ｻdata 蜈ｨ蟷�荳九�
    Dim ctlW As Long
    ctlW = m_formWidth - m_margin * 2
    Dim dataTopH As Long, dataHH As Long
    dataTopH = m_margin + m_headerYOffset + VLABEL_H + VLABEL_GAP
    dataHH = m_rowHeightSingle - VLABEL_H - VLABEL_GAP - VROW_SPACING

    ' Phase R-2 F-3: format selector as dropdown or textbox.
    Dim ctl As Object
    If m_formatSelectorType = "dropdown" Then
        Set ctl = designer.Controls.Add(PROGID_COMBOBOX, "cboFormatId", True)
        ctl.Style = 2   ' fmStyleDropDownList
        Dim fmts As Variant
        fmts = ListFormatIds()
        Dim fi As Long
        For fi = LBound(fmts) To UBound(fmts)
            ctl.AddItem CStr(fmts(fi))
        Next fi
        m_comboItemsByCtl("cboFormatId") = fmts
        If Len(m_formatId) > 0 Then m_comboInitialByCtl("cboFormatId") = m_formatId
        ctl.Locked = m_readOnlyFormat
    Else
        Set ctl = designer.Controls.Add(PROGID_TEXTBOX, "txtFormatId", True)
        ctl.Text = m_formatId
        ctl.Locked = m_readOnlyFormat
        If m_readOnlyFormat Then ctl.BackColor = RGB(240, 240, 240)
    End If
    ctl.top = dataTopH
    ctl.left = m_margin
    ctl.Width = ctlW
    ctl.Height = dataHH
    ApplyBaseFont ctl

    ' Phase R-2 F-2 / R-3-ﾏ�-3: help line under the format selector (data 陦檎峩荳�)縲�
    Dim helpH As Long
    helpH = RenderHelpLine(designer, "lblFormatHelp", dataTopH + dataHH + 2, m_margin, ctlW, m_formatHelp)
    ' R-3-g1A: return the actual format-block bottom (control + formatHelp + gap)
    ' so the header-field row never overlaps the format help line (D-7 overlap fix).
    Dim hdrBottom As Long
    hdrBottom = dataTopH + dataHH
    If helpH > 0 Then hdrBottom = hdrBottom + 2 + helpH
    AddHeaderRow = hdrBottom + 6
End Function

' Phase R-2 F-3: list available format ids (basenames of formats/*.txt).
' Returns a string array (possibly empty).
Private Function ListFormatIds() As Variant
    On Error GoTo Fail
    Dim col As Collection
    Set col = modFormatLoader.ListAllFormats()
    If col Is Nothing Then GoTo Fail
    If col.count = 0 Then
        ListFormatIds = Array()
        Exit Function
    End If
    Dim arr() As String
    ReDim arr(0 To col.count - 1)
    Dim i As Long
    For i = 1 To col.count
        arr(i - 1) = CStr(col.Item(i))
    Next i
    ListFormatIds = arr
    Exit Function
Fail:
    ListFormatIds = Array()
End Function

' Phase R-2 F-1: number of configured header fields.
Private Function HeaderFieldCount() As Long
    Dim s As String
    s = Trim$(m_headerFields)
    If Len(s) = 0 Then
        HeaderFieldCount = 0
    Else
        HeaderFieldCount = UBound(Split(s, ",")) - LBound(Split(s, ",")) + 1
    End If
End Function

' Phase R-2 F-1: map a logical header-field id to its JP label.
Private Function ResolveHeaderFieldLabel(ByVal id As String) As String
    Select Case LCase$(Trim$(id))
        Case "knowledgeid": ResolveHeaderFieldLabel = ChrW(&H4E88) & ChrW(&H5B9A) & ChrW(&H756A) & ChrW(&H53F7)        ' 莠亥ｮ夂分蜿ｷ
        Case "createdat":   ResolveHeaderFieldLabel = ChrW(&H767B) & ChrW(&H9332) & ChrW(&H65E5) & ChrW(&H6642)        ' 逋ｻ骭ｲ譌･譎�
        Case "updatedat":   ResolveHeaderFieldLabel = ChrW(&H66F4) & ChrW(&H65B0) & ChrW(&H65E5) & ChrW(&H6642)        ' 譖ｴ譁ｰ譌･譎�
        Case Else:          ResolveHeaderFieldLabel = id
    End Select
End Function

' Phase R-2 F-1: resolve a header field's current value from the loaded
' knowledge data (edit/view); register mode returns blank.
Private Function ResolveHeaderFieldValue(ByVal id As String) As String
    Dim key As String
    Select Case LCase$(Trim$(id))
        Case "knowledgeid"
            ResolveHeaderFieldValue = m_knowledgeId
            Exit Function
        Case "createdat":  key = "CreatedAt"
        Case "updatedat":  key = "UpdatedAt"
        Case Else:         key = id
    End Select
    If Not m_knowledgeData Is Nothing Then
        If m_knowledgeData.Exists(key) Then ResolveHeaderFieldValue = CStr(m_knowledgeData(key))
    End If
End Function

' Phase R-2 F-1: render readonly header fields below the format row.
' Returns the y position after the last header field (= subheader top).
Private Function AddHeaderFields(ByVal designer As Object, ByVal yStart As Long) As Long
    Dim s As String
    s = Trim$(m_headerFields)
    If Len(s) = 0 Then
        AddHeaderFields = yStart
        Exit Function
    End If
    Dim ids() As String
    ids = Split(s, ",")
    Dim y As Long
    y = yStart
    Dim k As Long
    For k = LBound(ids) To UBound(ids)
        Dim id As String
        id = Trim$(ids(k))
        If Len(id) > 0 Then
            Dim idxStr As String
            idxStr = Format$(k + 1, "000")
            Dim lbl As Object
            Set lbl = designer.Controls.Add(PROGID_LABEL, "hdrlbl_" & idxStr, True)
            ' Phase R-3-ﾏ�-3 邵ｦ荳ｦ縺ｳ: header field 繧� label 荳翫�ｻdata 蜈ｨ蟷�荳九�
            lbl.caption = ResolveHeaderFieldLabel(id)
            lbl.top = y
            lbl.left = m_margin
            lbl.Width = m_formWidth - m_margin * 2
            lbl.Height = VLABEL_H
            ApplyBaseFont lbl

            Dim ctl As Object
            Set ctl = designer.Controls.Add(PROGID_TEXTBOX, "hdrctl_" & idxStr, True)
            ctl.Text = ResolveHeaderFieldValue(id)
            Dim dataTopHF As Long, dataHHF As Long
            dataTopHF = y + VLABEL_H + VLABEL_GAP
            dataHHF = m_rowHeightSingle - VLABEL_H - VLABEL_GAP - VROW_SPACING
            ctl.top = dataTopHF
            ctl.left = m_margin
            ctl.Width = m_formWidth - m_margin * 2
            ctl.Height = dataHHF
            ' knowledgeId is auto-assigned but editable; view/preview locks it.
            ctl.Locked = (m_mode = "view" Or m_mode = "preview")
            ApplyBaseFont ctl

            Dim hfExtra As Long
            hfExtra = 0
            ' Phase R-2 F-2: per-header-field help line (data 陦檎峩荳�)縲�
            If Not m_headerHelp Is Nothing Then
                If m_headerHelp.Exists(id) Then
                    hfExtra = RenderHelpLine(designer, "hdrhelp_" & idxStr, dataTopHF + dataHHF + 2, _
                            m_margin, m_formWidth - m_margin * 2, CStr(m_headerHelp(id)))
                End If
            End If
            y = y + m_rowHeightSingle + hfExtra
        End If
    Next k
    AddHeaderFields = y
End Function

' Apply 繝｡繧､繝ｪ繧ｪ base font (Phase R-1-c, spec �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ1)
Private Sub ApplyBaseFont(ByVal ctl As Object)
    On Error Resume Next
    ctl.Font.Name = m_fontName
    ctl.Font.Size = m_fontSize
    On Error GoTo 0
End Sub

' Subheader "�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽi�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽb�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽW�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽﾌ難ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽe" row (between header and dynamic fields, per mock)
' V3 fix (2026-05-29): mock navy title bar at form top.
' OS title bar (VBA-uncontrollable) stays; this paints a navy band
' just below it inside client area with the form's JP title.
Private Sub AddTitleBar(ByVal designer As Object)
    On Error Resume Next
    Dim tb As Object
    Set tb = designer.Controls.Add(PROGID_LABEL, "lblTitleBar", True)
    tb.caption = " " & ResolveFormTitle()
    tb.top = 0
    tb.left = 0
    tb.Width = m_formWidth
    tb.Height = TITLE_BAR_HEIGHT
    tb.BackColor = TITLE_BAR_COLOR
    tb.ForeColor = TITLE_BAR_FORE_COLOR
    tb.TextAlign = 1
    tb.Font.Name = m_fontName
    tb.Font.Size = TITLE_BAR_FONT_SIZE
    tb.Font.Bold = True
    On Error GoTo 0
End Sub

Private Sub AddSubheaderRow(ByVal designer As Object, ByVal y As Long)
    On Error Resume Next
    Dim lbl As Object
    Set lbl = designer.Controls.Add(PROGID_LABEL, "lblSubheader", True)
    ' �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽi�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽb�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽW�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽﾌ難ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽe
    lbl.caption = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H306E) & ChrW(&H5185) & ChrW(&H5BB9)
    lbl.top = y + 4
    lbl.left = m_margin
    lbl.Width = m_formWidth - m_margin * 2
    lbl.Height = m_subheaderHeight - 4
    lbl.Font.Name = m_fontName
    lbl.Font.Size = m_subheaderFontSize
    lbl.Font.Bold = True
    lbl.ForeColor = SUBHEADER_COLOR
    On Error GoTo 0
End Sub

Private Function AddFieldRow(ByVal designer As Object, _
                              ByVal sec As ClsStanzaSection, _
                              ByVal y As Long, _
                              ByVal knowledgeData As Object) As Long
    Dim fieldName As String
    Dim fieldType As String
    Dim curVal As String
    fieldName = sec.GetValue("FieldName")
    fieldType = sec.GetValue("FieldType")
    curVal = ""
    If Not knowledgeData Is Nothing Then
        If knowledgeData.Exists(fieldName) Then curVal = CStr(knowledgeData(fieldName))
    End If

    ' Phase O-2: stable index-based control names to avoid SafeName collisions
    ' for multi-byte JP field names. m_fieldNamesByIdx maps ctlName -> fieldName.
    Dim idx As Long
    idx = m_fieldNamesByIdx.count + 1
    Dim idxStr As String
    idxStr = Format$(idx, "000")
    Dim ctlName As String
    ctlName = "ctl_" & idxStr

    ' Label
    Dim lbl As Object
    Set lbl = designer.Controls.Add(PROGID_LABEL, "lbl_" & idxStr, True)
    ' Phase R-3-ﾏ�-3 邵ｦ荳ｦ縺ｳ: label 縺ｯ荳翫�ｮ陦後�ｻ蜈ｨ蟷�(badge 蛻�繧呈而髯､)縲�
    lbl.caption = fieldName
    lbl.top = y
    lbl.left = m_margin
    lbl.Width = m_formWidth - m_margin * 2 - m_badgeWidth - 4
    lbl.Height = VLABEL_H
    ApplyBaseFont lbl

    ' Phase R-1-c: �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽK�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ{ badge for required fields (FieldRequired = TRUE / true)
    Dim isReq As Boolean
    Dim reqStr As String
    ' Phase R-1-j: format files use "Required"; fall back to legacy "FieldRequired".
    reqStr = LCase$(Trim$(sec.GetValue("Required")))
    If Len(reqStr) = 0 Then reqStr = LCase$(Trim$(sec.GetValue("FieldRequired")))
    isReq = (reqStr = "true" Or reqStr = "1" Or reqStr = "yes" Or reqStr = "tr")
    If isReq Then
        Dim badge As Object
        Set badge = designer.Controls.Add(PROGID_LABEL, "bdg_" & idxStr, True)
        badge.caption = ChrW(&H5FC5) & ChrW(&H9808)   ' �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽK�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ{
        ' 邵ｦ荳ｦ縺ｳ: badge 縺ｯ label 陦後�ｮ蜿ｳ遶ｯ縺ｫ驟咲ｽｮ
        badge.top = y + (VLABEL_H - m_badgeHeight) \ 2
        badge.left = m_formWidth - m_margin - m_badgeWidth
        badge.Width = m_badgeWidth
        badge.Height = m_badgeHeight
        badge.TextAlign = 2   ' fmTextAlignCenter
        badge.BackColor = BADGE_BACK_COLOR
        badge.ForeColor = BADGE_TEXT_COLOR
        On Error Resume Next
        badge.Font.Name = m_fontName
        badge.Font.Size = 9
        badge.Font.Bold = True
        On Error GoTo 0
    End If

    Dim rowH As Long
    rowH = m_rowHeightSingle
    Dim ctl As Object
    m_fieldNamesByIdx(ctlName) = fieldName

    Select Case fieldType
        Case ChrW(&H5358) & ChrW(&H4E00) & ChrW(&H884C)  ' �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽP�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽs
            Set ctl = designer.Controls.Add(PROGID_TEXTBOX, ctlName, True)
            ctl.MultiLine = False
            ctl.Text = curVal
        Case ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C)  ' �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽs
            Set ctl = designer.Controls.Add(PROGID_TEXTBOX, ctlName, True)
            ctl.MultiLine = True
            ctl.WordWrap = True
            ctl.EnterKeyBehavior = True
            ' Phase R-1-f: scroll bar mode from [USERFORM].multiLineScrollBars
            Select Case LCase$(m_multiLineScrollBars)
                Case "vertical": ctl.ScrollBars = 2   ' fmScrollBarsVertical
                Case "horizontal": ctl.ScrollBars = 1   ' fmScrollBarsHorizontal
                Case "both": ctl.ScrollBars = 3
                Case "none": ctl.ScrollBars = 0
                Case Else: ctl.ScrollBars = 2
            End Select
            ctl.Text = curVal
            ' Long multi-line for �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽﾆ手順/�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ/�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ/�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽﾚ搾ｿｽ/�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽe; others short
            If IsLongMultilineField(fieldName) Then
                rowH = m_rowHeightMultiLong
            Else
                rowH = m_rowHeightMulti
            End If
        Case ChrW(&H65E5) & ChrW(&H6642)  ' �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ - mark for Change-event validation
            Set ctl = designer.Controls.Add(PROGID_TEXTBOX, ctlName, True)
            ctl.MultiLine = False
            ctl.Text = curVal
            m_dateFieldIndices(idxStr) = True
        Case ChrW(&H9078) & ChrW(&H629E)  ' �ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽI�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽE�ｿｽ
            Set ctl = designer.Controls.Add(PROGID_COMBOBOX, ctlName, True)
            ctl.Style = 2   ' fmStyleDropDownList
            Dim opts As String
            opts = sec.GetValue("DropdownOptions")
            If Len(opts) > 0 Then
                Dim parts() As String
                parts = Split(opts, "|")
                Dim k As Long
                For k = LBound(parts) To UBound(parts)
                    parts(k) = Trim$(parts(k))
                Next k
                m_comboItemsByCtl(ctlName) = parts
            End If
            If Len(curVal) > 0 Then m_comboInitialByCtl(ctlName) = curVal
        Case Else
            Set ctl = designer.Controls.Add(PROGID_TEXTBOX, ctlName, True)
            ctl.MultiLine = False
            ctl.Text = curVal
    End Select

    ' Phase R-3-ﾏ�-3 邵ｦ荳ｦ縺ｳ: data 縺ｯ label 陦後�ｮ荳九�ｻ蜈ｨ蟷�縲�
    ' dataHeight = 陦後ヴ繝�繝�(rowH) - labelZone(VLABEL_H+VLABEL_GAP) - VROW_SPACING縲�
    Dim dataTop As Long, dataH As Long
    dataTop = y + VLABEL_H + VLABEL_GAP
    dataH = rowH - VLABEL_H - VLABEL_GAP - VROW_SPACING
    ctl.top = dataTop
    ctl.left = m_margin
    ctl.Width = m_formWidth - m_margin * 2
    ctl.Height = dataH
    ApplyBaseFont ctl

    ' Apply mode-based locking (view 縺ｨ preview 縺ｯ陦ｨ遉ｺ蟆ら畑 = readonly)
    If m_mode = "view" Or m_mode = "preview" Then
        On Error Resume Next
        ctl.Locked = True
        ctl.BackColor = RGB(240, 240, 240)
        On Error GoTo 0
    End If

    ' Phase R-2 F-4: placeholder grey text in empty TextBoxes (not view mode,
    ' not ComboBox). Stored so InjectFormCode can emit Enter/Exit handlers and
    ' the persistence layer can treat a still-placeholder field as empty.
    ' Phase R-3-ﾏ�-2: preview (M-04) 縺ｯ螳溘ョ繝ｼ繧ｿ繧呈戟縺溘★ placeholder(險伜�･萓�)縺ｮ縺ｿ陦ｨ遉ｺ縲�
    ' readonly 縺ｮ縺溘ａ focus-clear 繝上Φ繝峨Λ縺ｯ逕滓�舌○縺� static 縺ｫ蜃ｺ縺吶�
    If TypeName(ctl) = "TextBox" And m_mode <> "view" And Len(curVal) = 0 Then
        Dim ph As String
        ph = Trim$(sec.GetValue("fieldPlaceholder"))
        If Len(ph) > 0 Then
            On Error Resume Next
            ctl.Text = ph
            ctl.ForeColor = PLACEHOLDER_COLOR
            On Error GoTo 0
            If m_mode <> "preview" Then
                ctl.Tag = PLACEHOLDER_TAG
                m_placeholderByCtl(ctlName) = ph
            End If
        End If
    End If

    ' Phase R-2 F-2: optional per-field help line under the control. Grows the
    ' row so following fields do not overlap.
    Dim fHelp As String
    fHelp = Trim$(sec.GetValue("fieldHelp"))
    Dim extraH As Long
    extraH = RenderHelpLine(designer, "fldhelp_" & idxStr, dataTop + dataH + 2, m_margin, _
                            m_formWidth - m_margin * 2, fHelp)

    AddFieldRow = rowH + extraH
End Function

' Replace characters not valid in VBA identifiers with underscore.
Private Function SafeName(ByVal s As String) As String
    Dim i As Long
    Dim out As String
    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        Dim code As Long
        code = AscW(ch)
        If (code >= 65 And code <= 90) Or _
           (code >= 97 And code <= 122) Or _
           (code >= 48 And code <= 57) Then
            out = out & ch
        Else
            out = out & "_"
        End If
    Next i
    If Len(out) = 0 Then out = "f"
    SafeName = out
End Function

Private Sub AddButtonBar(ByVal designer As Object, ByVal y As Long)
    Dim btnTop As Long
    btnTop = y + 12
    Dim rightEdge As Long
    rightEdge = m_formWidth - m_margin
    Dim labels() As String
    Dim names() As String
    Dim kinds() As String   ' Phase R-1-c: button kind for coloring (p=primary, s=secondary, d=destructive)

    Select Case m_mode
        Case "register"
            names = Split("btnClear|btnRegister", "|")
            labels = Split(ChrW(&H30AF) & ChrW(&H30EA) & ChrW(&H30A2) & "|" & ChrW(&H767B) & ChrW(&H9332), "|")
            kinds = Split("s|p", "|")
        Case "edit"
            names = Split("btnDelete|btnUpdate", "|")
            labels = Split(ChrW(&H524A) & ChrW(&H9664) & "|" & ChrW(&H66F4) & ChrW(&H65B0), "|")
            kinds = Split("d|p", "|")
        Case "view"
            names = Split("btnEdit|btnDelete|btnClose", "|")
            labels = Split(ChrW(&H7DE8) & ChrW(&H96C6) & "|" & ChrW(&H524A) & ChrW(&H9664) & "|" & ChrW(&H9589) & ChrW(&H3058) & ChrW(&H308B), "|")
            kinds = Split("s|d|p", "|")
        Case "preview"
            ' Phase R-3-ﾏ�-2: M-04 繝励Ξ繝薙Η繝ｼ縺ｯ縲碁哩縺倥ｋ縲榊腰迢ｬ (M-09 縺ｨ蜷� btnClose)
            names = Split("btnClose", "|")
            labels = Split(ChrW(&H9589) & ChrW(&H3058) & ChrW(&H308B), "|")
            kinds = Split("p", "|")
        Case Else
            Exit Sub
    End Select

    ' Phase R-1-i: button help text per mock. Help label below each button.
    Dim helpTexts() As String
    Select Case m_mode
        Case "register"
            ' 繧ｯ繝ｪ繧｢ / 逋ｻ骭ｲ
            helpTexts = Split(ChrW(&H5165) & ChrW(&H529B) & ChrW(&H6B04) & ChrW(&H3092) & ChrW(&H7A7A) & ChrW(&H306B) & ChrW(&H623B) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & "|" & ChrW(&H5185) & ChrW(&H5BB9) & ChrW(&H3092) & ChrW(&H4FDD) & ChrW(&H5B58) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H767B) & ChrW(&H9332) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059), "|")
        Case "edit"
            ' 蜑企勁 / 譖ｴ譁ｰ
            helpTexts = Split(ChrW(&H78BA) & ChrW(&H8A8D) & ChrW(&H306E) & ChrW(&H3046) & ChrW(&H3048) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & "|" & ChrW(&H5185) & ChrW(&H5BB9) & ChrW(&H3092) & ChrW(&H4E0A) & ChrW(&H66F8) & ChrW(&H304D) & ChrW(&H4FDD) & ChrW(&H5B58) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059), "|")
        Case "view"
            ' 邱ｨ髮・/ 蜑企勁 / 髢峨§繧・
            helpTexts = Split("|" & "|" & ChrW(&H691C) & ChrW(&H7D22) & ChrW(&H753B) & ChrW(&H9762) & ChrW(&H306B) & ChrW(&H623B) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H3059), "|")
        Case "preview"
            ' 繝励Ξ繝薙Η繝ｼ繧帝哩縺倥∪縺�
            helpTexts = Split(ChrW(&H30D7) & ChrW(&H30EC) & ChrW(&H30D3) & ChrW(&H30E5) & ChrW(&H30FC) & ChrW(&H3092) & ChrW(&H9589) & ChrW(&H3058) & ChrW(&H307E) & ChrW(&H3059), "|")
        Case Else
            helpTexts = Split("|", "|")
    End Select

    Dim n As Long
    ' R-3-e (2026-05-28): config-driven button filter. [USERFORM] buttons (csv)
    ' keeps only the listed button ids from the mode-default set (no hardcode).
    If Len(Trim$(m_buttons)) > 0 Then
        Dim wantB() As String
        wantB = Split(m_buttons, ",")
        Dim fNames() As String, fLabels() As String, fKinds() As String, fHelp() As String
        ReDim fNames(0 To UBound(names))
        ReDim fLabels(0 To UBound(names))
        ReDim fKinds(0 To UBound(names))
        ReDim fHelp(0 To UBound(names))
        Dim fc As Long
        fc = 0
        Dim bi As Long, wj As Long
        For bi = 0 To UBound(names)
            Dim keepB As Boolean
            keepB = False
            For wj = 0 To UBound(wantB)
                If Trim$(wantB(wj)) = names(bi) Then keepB = True
            Next wj
            If keepB Then
                fNames(fc) = names(bi)
                fLabels(fc) = labels(bi)
                fKinds(fc) = kinds(bi)
                If bi <= UBound(helpTexts) Then fHelp(fc) = helpTexts(bi) Else fHelp(fc) = ""
                fc = fc + 1
            End If
        Next bi
        If fc > 0 Then
            ReDim Preserve fNames(0 To fc - 1)
            ReDim Preserve fLabels(0 To fc - 1)
            ReDim Preserve fKinds(0 To fc - 1)
            ReDim Preserve fHelp(0 To fc - 1)
            names = fNames
            labels = fLabels
            kinds = fKinds
            helpTexts = fHelp
        End If
    End If
    n = UBound(names) - LBound(names) + 1
    Dim i As Long
    For i = 0 To n - 1
        Dim btn As Object
        Set btn = designer.Controls.Add(PROGID_BUTTON, names(i), True)
        btn.caption = labels(i)
        btn.Width = m_buttonWidth
        btn.Height = m_buttonHeight
        btn.top = btnTop
        Dim slot As Long
        slot = (n - 1 - i)
        btn.left = rightEdge - m_buttonWidth - slot * (m_buttonWidth + m_buttonGap)
        On Error Resume Next
        btn.Font.Name = m_fontName
        btn.Font.Size = m_fontSize
        btn.Font.Bold = True
        Select Case kinds(i)
            Case "p"
                btn.BackColor = PRIMARY_BACK_COLOR
                btn.ForeColor = PRIMARY_FORE_COLOR
            Case "d"
                btn.BackColor = DESTRUCTIVE_BACK_COLOR
                btn.ForeColor = PRIMARY_FORE_COLOR
            Case Else
                btn.BackColor = RGB(240, 240, 240)
                btn.ForeColor = RGB(60, 60, 60)
        End Select
        ' V5 fix (2026-05-30) BUG-2: Default/Cancel 譏守､ｺ set縲Ｑrimary(p) 繧� Enter 縺ｧ逋ｺ轣ｫ縺吶ｋ
        ' Default 縺ｫ縲，lear/Close 繧� Esc 縺ｧ逋ｺ轣ｫ縺吶ｋ Cancel 縺ｫ縲ゅ％繧後〒 Tab 縺ｧ field 遘ｻ蜍穂ｸｭ縺ｮ
        ' 隱､ Enter 縺後後け繝ｪ繧｢縲阪ｒ逋ｺ轣ｫ縺励※ field 蜈ｨ豸亥､ｱ縺吶ｋ UX 莠区腐繧帝亟縺舌�
        On Error Resume Next
        Select Case kinds(i)
            Case "p"
                btn.Default = True
                btn.Cancel = False
            Case Else
                btn.Default = False
                ' btnClear / btnClose 繧� Esc 繧ｭ繝ｼ縺ｧ逋ｺ轣ｫ縺輔○繧�
                If names(i) = "btnClear" Or names(i) = "btnClose" Then
                    btn.Cancel = True
                Else
                    btn.Cancel = False
                End If
        End Select
        On Error GoTo 0

        ' R-3-g1A (2026-05-28): per-button help (mock 貅匁侠縲∵｡�c 邨ｱ蜷郁｣懷勧譁�繧呈彫蝗�)縲�
        ' buttonWidth(113) 蟷�蜀�縺ｫ蜿弱ａ蜿ｳ縺ｿ縺阪ｌ/驥阪↑繧翫ｒ蝗樣∩縲�
        If i <= UBound(helpTexts) Then
            If Len(helpTexts(i)) > 0 Then
                Dim bhelp As Object
                On Error Resume Next
                Set bhelp = designer.Controls.Add(PROGID_LABEL, "lbl_help_" & names(i), True)
                bhelp.caption = helpTexts(i)
                bhelp.top = btnTop + m_buttonHeight + 2
                bhelp.left = btn.left
                bhelp.Width = m_buttonWidth
                bhelp.Height = HELP_LINE_HEIGHT
                bhelp.WordWrap = False
                bhelp.TextAlign = 2
                bhelp.Font.Name = m_fontName
                bhelp.Font.Size = HELP_FONT_SIZE
                bhelp.ForeColor = HELP_TEXT_COLOR
                On Error GoTo 0
            End If
        End If
    Next i

End Sub

' Inject simple Click handlers into the form's code module so the form is
' self-contained. m_returnId is set via Application.Run callback (the form
' calls a Public module-level Sub to communicate back to this class).
Private Sub InjectFormCode(ByVal vbc As Object)
    On Error GoTo ErrHandler
    Dim cm As Object
    Set cm = vbc.CodeModule
    Dim s As String
    s = "Option Explicit" & vbCrLf & vbCrLf
    s = s & "Private Sub btnRegister_Click()" & vbCrLf
    s = s & "    Application.Run ""modUserFormCallback.OnRegister""" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnUpdate_Click()" & vbCrLf
    s = s & "    Application.Run ""modUserFormCallback.OnUpdate""" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    Application.Run ""modUserFormCallback.OnDelete""" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnEdit_Click()" & vbCrLf
    s = s & "    Application.Run ""modUserFormCallback.OnEdit""" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnClear_Click()" & vbCrLf
    s = s & "    Application.Run ""modUserFormCallback.OnClear""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnLoad_Click()" & vbCrLf
    s = s & "    Application.Run ""modUserFormCallback.OnLoad""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    ' V4 fix (2026-05-29) #M-05: in register mode the first empty DropDownList
    ' ComboBox halts the frScroll initial paint, leaving later fields blank.
    ' UserForm_Activate fires even under modal Show, so scroll the frame to the
    ' bottom then back to top with Repaint to force every field to paint.
    ' (edit/view already paint fine; ScrollTop returns to 0 so they look identical.)
    s = s & vbCrLf
    s = s & "Private Sub UserForm_Activate()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim frR As Object" & vbCrLf
    s = s & "    Set frR = Me.Controls(""frScroll"")" & vbCrLf
    s = s & "    If Not frR Is Nothing Then" & vbCrLf
    s = s & "        frR.ScrollTop = frR.ScrollHeight" & vbCrLf
    s = s & "        frR.Repaint" & vbCrLf
    s = s & "        DoEvents" & vbCrLf
    s = s & "        frR.ScrollTop = 0" & vbCrLf
    s = s & "        frR.Repaint" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Me.Repaint" & vbCrLf
    s = s & "    DoEvents" & vbCrLf
    s = s & "End Sub" & vbCrLf
    ' Phase R-2 F-3: format dropdown change -> re-render with new format's fields.
    ' Only injected in dropdown mode (the cboFormatId control exists).
    ' Phase R-3-ﾏ�-5: format 陦後′辟｡縺� mode 縺ｧ縺ｯ cboFormatId 縺悟ｭ伜惠縺励↑縺�縺溘ａ inject 縺励↑縺�縲�
    If m_formatSelectorType = "dropdown" And m_formatRowEnabled Then
        s = s & vbCrLf
        s = s & "Private Sub cboFormatId_Change()" & vbCrLf
        s = s & "    Application.Run ""modUserFormCallback.OnFormatChange""" & vbCrLf
        s = s & "End Sub" & vbCrLf
    End If
    ' Phase O-2: per-field Change handlers for ???? (date) validation.
    ' For each idx in m_dateFieldIndices we emit a Change handler that
    ' (a) accepts empty values, (b) IsDate-valid values keep BackColor=white,
    ' (c) invalid values paint BackColor=RGB(255,192,192) (light red).
    If Not (m_dateFieldIndices Is Nothing) Then
        Dim k As Variant
        For Each k In m_dateFieldIndices.Keys
            Dim cn As String
            cn = "ctl_" & CStr(k)
            s = s & vbCrLf
            s = s & "Private Sub " & cn & "_Change()" & vbCrLf
            s = s & "    Dim v As String" & vbCrLf
            s = s & "    On Error Resume Next" & vbCrLf
            ' Phase R-2 F-4: a still-showing placeholder is treated as empty.
            s = s & "    If Me." & cn & ".Tag = """ & PLACEHOLDER_TAG & """ Then" & vbCrLf
            s = s & "        Me." & cn & ".BackColor = vbWhite" & vbCrLf
            s = s & "        Exit Sub" & vbCrLf
            s = s & "    End If" & vbCrLf
            s = s & "    v = Trim$(CStr(Me." & cn & ".Text))" & vbCrLf
            s = s & "    If Len(v) = 0 Then" & vbCrLf
            s = s & "        Me." & cn & ".BackColor = vbWhite" & vbCrLf
            s = s & "    ElseIf IsDate(v) Then" & vbCrLf
            s = s & "        Me." & cn & ".BackColor = vbWhite" & vbCrLf
            s = s & "    Else" & vbCrLf
            s = s & "        Me." & cn & ".BackColor = RGB(255, 192, 192)" & vbCrLf
            s = s & "    End If" & vbCrLf
            s = s & "End Sub" & vbCrLf
        Next k
    End If

    ' Phase R-2 F-4: placeholder Enter/Exit handlers (clear grey hint on focus,
    ' restore it on blur if still empty).
    If Not (m_placeholderByCtl Is Nothing) Then
        Dim pk As Variant
        For Each pk In m_placeholderByCtl.Keys
            Dim pn As String
            pn = CStr(pk)
            Dim phTxt As String
            phTxt = Replace(CStr(m_placeholderByCtl(pn)), """", """""")
            s = s & vbCrLf
            s = s & "Private Sub " & pn & "_Enter()" & vbCrLf
            s = s & "    On Error Resume Next" & vbCrLf
            s = s & "    If Me." & pn & ".Tag = """ & PLACEHOLDER_TAG & """ Then" & vbCrLf
            s = s & "        Me." & pn & ".Text = """"" & vbCrLf
            s = s & "        Me." & pn & ".ForeColor = vbWindowText" & vbCrLf
            s = s & "        Me." & pn & ".Tag = """"" & vbCrLf
            s = s & "    End If" & vbCrLf
            s = s & "End Sub" & vbCrLf
            s = s & "Private Sub " & pn & "_Exit(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
            s = s & "    On Error Resume Next" & vbCrLf
            s = s & "    If Len(Trim$(CStr(Me." & pn & ".Text))) = 0 Then" & vbCrLf
            s = s & "        Me." & pn & ".Text = """ & phTxt & """" & vbCrLf
            s = s & "        Me." & pn & ".ForeColor = RGB(128, 128, 128)" & vbCrLf
            s = s & "        Me." & pn & ".Tag = """ & PLACEHOLDER_TAG & """" & vbCrLf
            s = s & "    End If" & vbCrLf
            s = s & "End Sub" & vbCrLf
        Next pk
    End If
    cm.AddFromString s
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] InjectFormCode: " & Err.Number & " " & Err.Description
End Sub

' Phase R-1-i matrix test harness: load config (default + sheet + format) and
' return current m_* as semicolon-separated key=value string. Does NOT build
' the UserForm or call .Show, so it's safely callable from PowerShell COM.
Public Function MeasureFormConfig(ByVal xlsmName As String, ByVal mode As String, _
                                   ByVal knowledgeId As String, ByVal formatId As String) As String
    m_xlsmName = xlsmName
    m_mode = LCase$(mode)
    m_knowledgeId = knowledgeId
    m_formatId = formatId
    InitFormConfig
    ApplyUserformStanza
    ApplyUserformStanzaFromFormat
    Dim sb As String
    sb = ""
    sb = sb & "formWidth=" & m_formWidth & ";"
    sb = sb & "formHeight=" & m_formHeight & ";"
    sb = sb & "headerHeight=" & m_headerHeight & ";"
    sb = sb & "labelWidth=" & m_labelWidth & ";"
    sb = sb & "margin=" & m_margin & ";"
    sb = sb & "rowHeightSingle=" & m_rowHeightSingle & ";"
    sb = sb & "rowHeightMulti=" & m_rowHeightMulti & ";"
    sb = sb & "rowHeightMultiLong=" & m_rowHeightMultiLong & ";"
    sb = sb & "subheaderHeight=" & m_subheaderHeight & ";"
    sb = sb & "buttonBarHeight=" & m_buttonBarHeight & ";"
    sb = sb & "buttonWidth=" & m_buttonWidth & ";"
    sb = sb & "buttonHeight=" & m_buttonHeight & ";"
    sb = sb & "buttonGap=" & m_buttonGap & ";"
    sb = sb & "bottomMargin=" & m_bottomMargin & ";"
    sb = sb & "badgeWidth=" & m_badgeWidth & ";"
    sb = sb & "badgeHeight=" & m_badgeHeight & ";"
    sb = sb & "fontName=" & m_fontName & ";"
    sb = sb & "fontSize=" & m_fontSize & ";"
    sb = sb & "subheaderFontSize=" & m_subheaderFontSize & ";"
    sb = sb & "multiLineScrollBars=" & m_multiLineScrollBars & ";"
    sb = sb & "caption=" & m_captionOverride & ";"
    sb = sb & "backColor=" & m_backColor
    MeasureFormConfig = sb
End Function

' Phase R-1-j modeless show. Builds the form then calls Show vbModeless.
' Returns immediately so PowerShell can capture a screenshot. Caller must
' invoke CloseModelessForm afterward.
Public Function ShowFormModeless(ByVal xlsmName As String, ByVal mode As String, _
                                  ByVal knowledgeId As String, ByVal formatId As String) As String
    On Error GoTo ErrHandler
    m_xlsmName = xlsmName
    m_mode = LCase$(mode)
    m_knowledgeId = knowledgeId
    m_formatId = formatId
    m_returnId = ""
    m_dumpToFile = ""
    m_modelessRequest = True
    modUserFormCallback.SetRenderer Me
    Dim knowledgeData As Object
    Set knowledgeData = Nothing
    BuildAndShow knowledgeData
    ShowFormModeless = "OK:" & m_dynFormName
    Exit Function
ErrHandler:
    ShowFormModeless = "ERR:" & Err.Number & ":" & Err.Description
End Function

Public Sub CloseModelessForm()
    On Error Resume Next
    If Not m_modelessFormInstance Is Nothing Then
        Unload m_modelessFormInstance
    End If
    If Not m_modelessVbc Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove m_modelessVbc
    End If
    Set m_modelessFormInstance = Nothing
    Set m_modelessVbc = Nothing
    m_modelessRequest = False
End Sub

' Phase R-1-j path verify: build the form like BuildAndShow but skip Show,
' dump designer.Controls + resolved m_* values + fmtFields to a file, then
' remove the VBComponent. PowerShell parses the file to verify the
' stanza -> loader -> control path.
Public Function TestR1j_DumpForm(ByVal xlsmName As String, ByVal mode As String, _
                                  ByVal knowledgeId As String, ByVal formatId As String, _
                                  ByVal dumpPath As String) As String
    On Error GoTo ErrHandler
    m_xlsmName = xlsmName
    m_mode = LCase$(mode)
    m_knowledgeId = knowledgeId
    m_formatId = formatId
    m_returnId = ""
    m_dumpToFile = dumpPath
    modUserFormCallback.SetRenderer Me
    Dim knowledgeData As Object
    Set knowledgeData = Nothing
    BuildAndShow knowledgeData
    TestR1j_DumpForm = "OK:" & dumpPath
    Exit Function
ErrHandler:
    TestR1j_DumpForm = "ERR:" & Err.Number & ":" & Err.Description
End Function

' Dump designer.Controls + resolved layout to a UTF-8 text file.
Private Sub DumpFormToFile(ByVal vbc As Object, ByVal designer As Object, _
                            ByVal fmtFields As Collection, ByVal formH As Long)
    On Error GoTo ErrHandler
    Dim s As String
    Dim fc As Long
    Dim cc As Long
    Dim i As Long, j As Long
    Dim fsec As ClsStanzaSection
    Dim fn As String, ft As String, fr As String
    Dim ctl As Object
    Dim nm As String, lft As String, tp As String, wd As String, ht As String
    Dim fname As String, fsize As String, extra As String, cls As String
    Dim cap As String, tx As String

    s = ""
    s = s & "[FORM]" & vbCrLf
    s = s & "width=" & vbc.Properties("Width") & vbCrLf
    s = s & "height=" & vbc.Properties("Height") & vbCrLf
    s = s & "caption=" & vbc.Properties("Caption") & vbCrLf
    s = s & "computedH=" & formH & vbCrLf

    s = s & "[RESOLVED]" & vbCrLf
    s = s & "formWidth=" & m_formWidth & vbCrLf
    s = s & "formHeight=" & m_formHeight & vbCrLf
    s = s & "headerHeight=" & m_headerHeight & vbCrLf
    s = s & "labelWidth=" & m_labelWidth & vbCrLf
    s = s & "margin=" & m_margin & vbCrLf
    s = s & "rowHeightSingle=" & m_rowHeightSingle & vbCrLf
    s = s & "rowHeightMulti=" & m_rowHeightMulti & vbCrLf
    s = s & "rowHeightMultiLong=" & m_rowHeightMultiLong & vbCrLf
    s = s & "subheaderHeight=" & m_subheaderHeight & vbCrLf
    s = s & "buttonBarHeight=" & m_buttonBarHeight & vbCrLf
    s = s & "buttonWidth=" & m_buttonWidth & vbCrLf
    s = s & "buttonHeight=" & m_buttonHeight & vbCrLf
    s = s & "buttonGap=" & m_buttonGap & vbCrLf
    s = s & "bottomMargin=" & m_bottomMargin & vbCrLf
    s = s & "badgeWidth=" & m_badgeWidth & vbCrLf
    s = s & "badgeHeight=" & m_badgeHeight & vbCrLf
    s = s & "fontName=" & m_fontName & vbCrLf
    s = s & "fontSize=" & m_fontSize & vbCrLf
    s = s & "subheaderFontSize=" & m_subheaderFontSize & vbCrLf
    s = s & "multiLineScrollBars=" & m_multiLineScrollBars & vbCrLf
    s = s & "captionOverride=" & m_captionOverride & vbCrLf
    s = s & "backColor=" & m_backColor & vbCrLf
    s = s & "mode=" & m_mode & vbCrLf
    s = s & "formatId=" & m_formatId & vbCrLf

    s = s & "[FMTFIELDS]" & vbCrLf
    fc = 0
    If Not fmtFields Is Nothing Then fc = fmtFields.Count
    s = s & "count=" & fc & vbCrLf
    For i = 1 To fc
        Set fsec = fmtFields.Item(i)
        fn = fsec.GetValue("FieldName")
        ft = fsec.GetValue("FieldType")
        fr = fsec.GetValue("Required")
        If Len(fr) = 0 Then fr = fsec.GetValue("FieldRequired")
        s = s & i & "|" & fn & "|" & ft & "|" & fr & vbCrLf
    Next i

    ' R-3-ﾏ�-4: designer.Controls 縺ｯ frame 蟄舌ｂ flatten 縺励※蜷ｫ繧(frame-relative 蠎ｧ讓�)縺溘ａ縲�
    ' frame 蟄舌�ｯ [CONTROLS] 縺九ｉ髯､螟悶＠ [FRAMECONTROLS] 縺ｫ縺ｮ縺ｿ蜃ｺ縺吶�
    Dim frChildren As Object
    Set frChildren = CreateObject("Scripting.Dictionary")
    If Not m_scrollFrame Is Nothing Then
        Dim fcx As Object
        For Each fcx In m_scrollFrame.Controls
            frChildren(fcx.Name) = True
        Next fcx
    End If
    s = s & "[CONTROLS]" & vbCrLf
    Dim hdrLines As String
    Dim emitted As Long
    hdrLines = "": emitted = 0
    cc = designer.Controls.Count
    For j = 0 To cc - 1
        Set ctl = designer.Controls(j)
        nm = "?": lft = "?": tp = "?": wd = "?": ht = "?"
        fname = "?": fsize = "?": extra = "": cls = "?"
        On Error Resume Next
        nm = ctl.Name
        cls = TypeName(ctl)
        lft = CStr(ctl.Left)
        tp = CStr(ctl.Top)
        wd = CStr(ctl.Width)
        ht = CStr(ctl.Height)
        fname = CStr(ctl.Font.Name)
        fsize = CStr(ctl.Font.Size)
        cap = ""
        cap = CStr(ctl.Caption)
        If Len(cap) > 0 Then
            extra = "caption=" & cap
        Else
            tx = ""
            tx = CStr(ctl.Text)
            extra = "text=" & tx
        End If
        On Error GoTo ErrHandler
        If Not frChildren.Exists(nm) Then
            emitted = emitted + 1
            hdrLines = hdrLines & emitted & "|" & nm & "|" & cls & "|" & lft & "|" & tp & "|" & wd & "|" & ht & "|" & fname & "|" & fsize & "|" & extra & vbCrLf
        End If
    Next j
    s = s & "count=" & emitted & vbCrLf & hdrLines

    ' Phase R-3-ﾏ�-4: frScroll frame geometry + 縺昴�ｮ蟄� controls (frame-relative 蠎ｧ讓�)
    s = s & "[FRAME]" & vbCrLf
    If m_scrollFrame Is Nothing Then
        s = s & "present=0" & vbCrLf
    Else
        s = s & "present=1" & vbCrLf
        s = s & "name=" & m_scrollFrame.Name & vbCrLf
        s = s & "top=" & m_scrollFrame.top & vbCrLf
        s = s & "left=" & m_scrollFrame.left & vbCrLf
        s = s & "width=" & m_scrollFrame.Width & vbCrLf
        s = s & "height=" & m_scrollFrame.Height & vbCrLf
        s = s & "scrollHeight=" & m_scrollHeightPx & vbCrLf
        s = s & "headerHeight=" & m_headerHeightPx & vbCrLf
        s = s & "[FRAMECONTROLS]" & vbCrLf
        Dim fcc As Long
        fcc = m_scrollFrame.Controls.Count
        s = s & "count=" & fcc & vbCrLf
        Dim jf As Long
        For jf = 0 To fcc - 1
            Dim fctl As Object
            Set fctl = m_scrollFrame.Controls(jf)
            Dim fnm As String, flft As String, ftp As String, fwd As String, fht As String, fcls As String, fextra As String
            fnm = "?": flft = "?": ftp = "?": fwd = "?": fht = "?": fcls = "?": fextra = ""
            On Error Resume Next
            fnm = fctl.Name
            fcls = TypeName(fctl)
            flft = CStr(fctl.Left)
            ftp = CStr(fctl.Top)
            fwd = CStr(fctl.Width)
            fht = CStr(fctl.Height)
            Dim fcap As String
            fcap = "": fcap = CStr(fctl.Caption)
            If Len(fcap) > 0 Then
                fextra = "caption=" & fcap
            Else
                Dim ftx As String
                ftx = "": ftx = CStr(fctl.Text)
                fextra = "text=" & ftx
            End If
            On Error GoTo ErrHandler
            s = s & (jf + 1) & "|" & fnm & "|" & fcls & "|" & flft & "|" & ftp & "|" & fwd & "|" & fht & "|||" & fextra & vbCrLf
        Next jf
    End If

    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = "UTF-8"
    ado.Open
    ado.WriteText s
    ado.SaveToFile m_dumpToFile, 2
    ado.Close
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] DumpFormToFile: " & Err.Number & " " & Err.Description
End Sub

' --- Callback target setters (called by modUserFormCallback) ---
Public Sub SetReturnId(ByVal v As String)
    m_returnId = v
End Sub

' Phase R-2 F-3: format dropdown changed. Only re-render when the new format
' actually differs from the current one (guards against the Change event that
' fires when the initial value is set programmatically).
' R-3-a: btnLoad -> set the knowledge number, flag a reload; ShowForm loop refetches.
Public Sub RequestLoad(ByVal knowledgeNo As String)
    m_knowledgeId = Trim$(knowledgeNo)
    m_loadRequested = True
    m_reformatRequested = True
End Sub

Public Sub RequestReformat(ByVal newFormatId As String)
    If Len(newFormatId) = 0 Then Exit Sub
    If newFormatId = m_formatId Then Exit Sub
    m_formatId = newFormatId
    m_reformatRequested = True
End Sub

' Phase P E2E only: pre-seed m_formatId before ShowForm so the dynamic
' form has fields even in register mode (production code reads formatId
' from a sheet selector that is not yet wired).
Public Sub TestSeedFormatId(ByVal fmtId As String)
    m_formatId = fmtId
End Sub
Public Function GetDynFormName() As String
    GetDynFormName = m_dynFormName
End Function
' Phase O-2: expose ctlName -> fieldName mapping for the callback bridge.
Public Function GetFieldNameForCtl(ByVal ctlName As String) As String
    If m_fieldNamesByIdx Is Nothing Then Exit Function
    If m_fieldNamesByIdx.Exists(ctlName) Then
        GetFieldNameForCtl = CStr(m_fieldNamesByIdx(ctlName))
    End If
End Function
Public Function GetMode() As String
    GetMode = m_mode
End Function
Public Function GetKnowledgeId() As String
    GetKnowledgeId = m_knowledgeId
End Function
Public Function GetFormatId() As String
    GetFormatId = m_formatId
End Function

' --- IScreenRenderer stubs (kept for compat) ---
Private Sub IScreenRenderer_BindSheet(ByVal sheetName As String)
End Sub
Private Sub IScreenRenderer_ClearScreen()
End Sub
Private Sub IScreenRenderer_ApplyFromStanza(ByVal xlsmName As String, ByVal screenId As String)
End Sub
Private Sub IScreenRenderer_ShowSheet()
End Sub
Private Sub IScreenRenderer_HideSheet()
End Sub
Private Sub IScreenRenderer_ActivateSheet()
End Sub
Private Sub IScreenRenderer_ProtectSheet(ByVal level As String)
End Sub
Private Sub IScreenRenderer_UnprotectSheet()
End Sub
```
