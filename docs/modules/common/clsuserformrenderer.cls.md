---
title: clsUserFormRenderer.cls
description: clsUserFormRenderer.cls のソースコード（コピペ用）
---

# clsUserFormRenderer.cls

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: クラスモジュール
**更新日**: 2026-06-12 08:51 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsUserFormRenderer.cls`
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
'   formWidth = 540 (案A訂正 mock 720px=540pt), label width = 120, control width = formWidth - 140
'   headerHeight = 48, rowHeight ?P??s/????/?I?? = 24, rowHeight ?????s = 60
'   buttonBarHeight = 40, bottomMargin = 16
' ================================================================
Implements IScreenRenderer
Option Explicit

' --- Layout DEFAULTS (§A4 案A訂正後・mock 準拠 540pt/720px; overridable via [USERFORM] stanza) ---
' §A4 実装 SSOT: formWidth=540(pt; mock 720px 準拠 案A 2026-05-28), labelWidth=120, headerHeight=48, rowHeight(単一行/日時/選択)=24,
' rowHeight(複数行)=60, margin=10, buttonBarHeight=40, button 高さ=24/間隔=8, bottomMargin=16.
' All values can be overridden per-screen via ui_seed/<role>/M-XX.txt [USERFORM] section.
Private Const DEFAULT_FORM_WIDTH As Long = 486
Private Const DEFAULT_FORM_HEIGHT As Long = 0                 ' 0 = auto-compute from rows
Private Const DEFAULT_HEADER_HEIGHT As Long = 48
Private Const DEFAULT_LABEL_WIDTH As Long = 107
Private Const DEFAULT_MARGIN As Long = 18
Private Const DEFAULT_BADGE_WIDTH As Long = 29
Private Const DEFAULT_BADGE_HEIGHT As Long = 16
Private Const DEFAULT_BADGE_GAP As Long = 3
' Phase R-3-χ-3 (2026-05-28): 縦並び layout (label 行上 + data 行全幅下)。
' row pitch(total) = labelZone(18) + dataHeight + rowSpacing(6)。
'   single 48 = 18 + 24 + 6 / multi 114 = 18 + 90 + 6 / multiLong 129 = 18 + 105 + 6
Private Const DEFAULT_ROW_HEIGHT_SINGLE As Long = 48
Private Const DEFAULT_ROW_HEIGHT_MULTI As Long = 114
Private Const DEFAULT_ROW_HEIGHT_MULTI_LONG As Long = 129
' 縦並び layout 定数 (label control 高 / label-data 間 gap / data 下 spacing)。
' label zone = VLABEL_H + VLABEL_GAP = 18。dataHeight = rowPitch - 18 - VROW_SPACING。
Private Const VLABEL_H As Long = 16
Private Const VLABEL_GAP As Long = 2
Private Const VROW_SPACING As Long = 6
Private Const DEFAULT_SUBHEADER_HEIGHT As Long = 28
Private Const DEFAULT_BUTTON_BAR_HEIGHT As Long = 48
Private Const DEFAULT_BUTTON_WIDTH As Long = 113
Private Const DEFAULT_BUTTON_HEIGHT As Long = 26
Private Const DEFAULT_BUTTON_GAP As Long = 11
Private Const DEFAULT_BOTTOM_MARGIN As Long = 16
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
Private m_formatRowEnabled As Boolean      ' Phase R-3-χ-5: format 行表示(default mode=register のみ、ui_seed formatRowEnabled で上書き可)
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
' Phase R-3-χ-4 (2026-05-28): 案C/案A scroll 化 ? 固定 header + frScroll Frame。
Private Const PROGID_FRAME As String = "Forms.Frame.1"
' V5 fix (2026-05-30) BUG-1: 1080px 画面で button bar が画面外に出る regression を防ぐため
' SCROLL_FORM_CAP を 900→720 に下げ (72pt/inch × 10in = 720pt = ~960px、Excel chrome + taskbar 控除後の
' 安全圏)。超過分は frScroll 内スクロールで吸収する。
' V4 fix (2026-05-29) #3 履歴: M-09「原因」隠れ回避で 900pt に拡張したが、1080px screen で
' button bar 画面外事故が出たため再縮小。長 field は frame scroll で見せる。
Private Const SCROLL_FORM_CAP As Long = 720   ' form inside 高の上限(pt)。超過分は frame 内スクロール。

' --- Module-level state (single-instance UserForm session) ---
Private m_returnId As String          ' return value from form: knowledgeId / "DELETED" / ""
Private m_xlsmName As String          ' xlsm name (kanri / touroku / etc)
Private m_mode As String              ' "register" / "edit" / "view"
Private m_formatId As String          ' selected format id
Private m_knowledgeId As String       ' current knowledge id
Private m_readOnlyFormat As Boolean   ' format selection locked (edit/view)
Private m_fieldCount As Long
Private m_dynFormName As String

' Phase R-3-χ-4: scroll 化 ? 固定 header の下に置く frScroll Frame と寸法。
Private m_scrollFrame As Object       ' frScroll Frame (field 群の container)
Private m_headerHeightPx As Long      ' 固定 header 高(format selector + button bar)
Private m_scrollHeightPx As Long      ' frame 内 content total 高 (= Frame.ScrollHeight)

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
Private m_fieldRequiredByCtl As Object     ' [BUG-B16] Dictionary: ctlName -> Boolean (required)
Private m_dateFieldIndices As Object       ' Dictionary: idxStr -> True for ???? fields
' Phase P fix (2026-05-27): ComboBox items added on Designer.Controls.Add
' do not persist to the runtime UserForm instance returned by
' VBA.UserForms.Add. Stash items per ctlName during AddFieldRow, then
' populate the live instance after UserForms.Add (before Show).
Private m_comboItemsByCtl As Object       ' Dictionary: ctlName -> Array of items
Private m_comboInitialByCtl As Object     ' Dictionary: ctlName -> initial selected value
Private m_displayToFmtId As Object        ' 2026-06-11 fix: Dictionary: FormatName(display) -> FormatID
Private m_placeholderByCtl As Object      ' Phase R-2 F-4: ctlName -> placeholder text

' --- Public API (Phase N-3, ADR-0073 ??3.1) ---
' xlsmName: "???" / "?o?^?C??" / "????" (used for log scope only; form is workbook-local)
' mode    : "register" | "edit" | "view" | "preview"
' knowledgeId: edit/view only (register/preview pass "")
' readOnlyFormat: TRUE pins format selection (edit/view/preview). FALSE allows change (register).
' formatId: preview (M-04) only ? fixed format to preview; data is not loaded.
Public Function ShowForm( _
        ByVal xlsmName As String, _
        ByVal mode As String, _
        ByVal knowledgeId As String, _
        Optional ByVal readOnlyFormat As Boolean = False, _
        Optional ByVal formatId As String = "" _
) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0772] clsUserFormRenderer.ShowForm ENTER"  ' [ADR-0100]

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
    m_mode = LCase(mode)
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
            ElseIf knowledgeData.Exists("formatId") Then
                m_formatId = CStr(knowledgeData("formatId"))
            End If
        End If
    Else
        Set knowledgeData = Nothing
    End If

    ' Phase R-3-χ-2: preview mode (M-04) ? fixed format from caller, no data load.
    ' M-09 と同 interface (表示専用 popup) で、差は「実データを load しない」点のみ。
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
                    If knowledgeData.Exists("FormatID") Then
                        m_formatId = CStr(knowledgeData("FormatID"))
                    ElseIf knowledgeData.Exists("formatId") Then
                        m_formatId = CStr(knowledgeData("formatId"))
                    End If
                End If
            End If
        End If
        BuildAndShow knowledgeData
    Loop While m_reformatRequested

    ShowForm = m_returnId
    GoTo Cleanup

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0773] clsUserFormRenderer.ShowForm EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0774] clsUserFormRenderer.InitFormConfig ENTER"  ' [ADR-0100]
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
    ' Phase R-3-χ-5: format 行は既定で register のみ表示(M-04/M-06/M-09 は本体から除外)。
    m_formatRowEnabled = (m_mode = "register")
    m_formatHelp = ""
    Set m_headerHelp = CreateObject("Scripting.Dictionary")
    m_knowledgeNoRow = "off"
    m_knowledgeNoLabel = ""
    m_loadButtonLabel = ""
    m_knowledgeNoHelp = ""
    m_headerYOffset = 0
    m_buttons = ""
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0775] clsUserFormRenderer.InitFormConfig EXIT-OK"  ' [ADR-0100]
End Sub

' Parse an RRGGBB hex string to a VBA color Long (&HBBGGRR). Returns -1 if
' the string is empty or unparseable (caller leaves the OS default).
Private Function ParseHexColor(ByVal rrggbb As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0776] clsUserFormRenderer.ParseHexColor ENTER"  ' [ADR-0100]
    On Error GoTo Fail
    Dim s As String
    s = Trim(rrggbb)
    If Left(s, 1) = "#" Then s = Mid(s, 2)
    If Len(s) <> 6 Then GoTo Fail
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid(s, 1, 2))
    g = CLng("&H" & Mid(s, 3, 2))
    b = CLng("&H" & Mid(s, 5, 2))
    ParseHexColor = RGB(r, g, b)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0777] clsUserFormRenderer.ParseHexColor EXIT-OK"  ' [ADR-0100]
    Exit Function
Fail:
    ParseHexColor = -1
End Function

' Map (xlsmName, mode) -> screenId for ui_seed lookup (Phase R-1-f).
Private Function ResolveScreenId() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0778] clsUserFormRenderer.ResolveScreenId ENTER"  ' [ADR-0100]
    Select Case m_mode
        Case "register": ResolveScreenId = "M-05"
        Case "edit":     ResolveScreenId = "M-06"
        Case "view":     ResolveScreenId = "M-09"
        Case "preview":  ResolveScreenId = "M-04"
        Case Else:       ResolveScreenId = ""
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0779] clsUserFormRenderer.ResolveScreenId EXIT-OK"  ' [ADR-0100]
End Function

' Phase R-1-f: load ui_seed/<role>/<screenId>.txt, find [USERFORM] section, override m_*.
Private Sub ApplyUserformStanza()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0780] clsUserFormRenderer.ApplyUserformStanza ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0781] clsUserFormRenderer.ApplyUserformStanza EXIT-OK"  ' [ADR-0100]
End Sub

' Phase R-1-h: format-level [USERFORM] override (highest priority, 3-stage fallback).
' format > sheet > default. Each key independent (per-key fallback).
Private Sub ApplyUserformStanzaFromFormat()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0782] clsUserFormRenderer.ApplyUserformStanzaFromFormat ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0783] clsUserFormRenderer.ApplyUserformStanzaFromFormat EXIT-OK"  ' [ADR-0100]
End Sub

' Per-key override (only set if key present + parseable). Phase R-1-f.
Private Sub ApplyConfigFromStanza(ByVal sec As ClsStanzaSection)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0784] clsUserFormRenderer.ApplyConfigFromStanza ENTER"  ' [ADR-0100]
    Dim v As String
    v = Trim(sec.GetValue("formWidth"))
    If Len(v) > 0 And IsNumeric(v) Then m_formWidth = CLng(v)
    v = Trim(sec.GetValue("formHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_formHeight = CLng(v)
    v = Trim(sec.GetValue("headerHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_headerHeight = CLng(v)
    v = Trim(sec.GetValue("labelWidth"))
    If Len(v) > 0 And IsNumeric(v) Then m_labelWidth = CLng(v)
    v = Trim(sec.GetValue("margin"))
    If Len(v) > 0 And IsNumeric(v) Then m_margin = CLng(v)
    v = Trim(sec.GetValue("rowHeightSingle"))
    If Len(v) > 0 And IsNumeric(v) Then m_rowHeightSingle = CLng(v)
    v = Trim(sec.GetValue("rowHeightMulti"))
    If Len(v) > 0 And IsNumeric(v) Then m_rowHeightMulti = CLng(v)
    v = Trim(sec.GetValue("rowHeightMultiLong"))
    If Len(v) > 0 And IsNumeric(v) Then m_rowHeightMultiLong = CLng(v)
    v = Trim(sec.GetValue("subheaderHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_subheaderHeight = CLng(v)
    v = Trim(sec.GetValue("buttonBarHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_buttonBarHeight = CLng(v)
    v = Trim(sec.GetValue("buttonWidth"))
    If Len(v) > 0 And IsNumeric(v) Then m_buttonWidth = CLng(v)
    v = Trim(sec.GetValue("buttonHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_buttonHeight = CLng(v)
    v = Trim(sec.GetValue("buttonGap"))
    If Len(v) > 0 And IsNumeric(v) Then m_buttonGap = CLng(v)
    v = Trim(sec.GetValue("bottomMargin"))
    If Len(v) > 0 And IsNumeric(v) Then m_bottomMargin = CLng(v)
    v = Trim(sec.GetValue("badgeWidth"))
    If Len(v) > 0 And IsNumeric(v) Then m_badgeWidth = CLng(v)
    v = Trim(sec.GetValue("badgeHeight"))
    If Len(v) > 0 And IsNumeric(v) Then m_badgeHeight = CLng(v)
    v = Trim(sec.GetValue("fontSize"))
    If Len(v) > 0 And IsNumeric(v) Then m_fontSize = CLng(v)
    v = Trim(sec.GetValue("subheaderFontSize"))
    If Len(v) > 0 And IsNumeric(v) Then m_subheaderFontSize = CLng(v)
    v = Trim(sec.GetValue("fontName"))
    If Len(v) > 0 Then m_fontName = v
    v = Trim(sec.GetValue("multiLineScrollBars"))
    If Len(v) > 0 Then m_multiLineScrollBars = LCase(v)
    v = Trim(sec.GetValue("caption"))
    If Len(v) > 0 Then m_captionOverride = v
    v = Trim(sec.GetValue("backColor"))
    If Len(v) > 0 Then m_backColor = v
    v = Trim(sec.GetValue("headerFields"))
    If Len(v) > 0 Then m_headerFields = v
    v = Trim(sec.GetValue("formatSelectorType"))
    If Len(v) > 0 Then m_formatSelectorType = LCase(v)
    ' Phase R-3-χ-5: format 行の表示可否を ui_seed で明示上書き(M-04/06/09=false)。
    v = LCase(Trim(sec.GetValue("formatRowEnabled")))
    If Len(v) > 0 Then m_formatRowEnabled = (v = "true" Or v = "1" Or v = "yes")
    v = Trim(sec.GetValue("formatHelp"))
    If Len(v) > 0 Then m_formatHelp = v
    v = Trim(sec.GetValue("knowledgeNoRow"))
    If Len(v) > 0 Then m_knowledgeNoRow = LCase(v)
    v = Trim(sec.GetValue("knowledgeNoLabel"))
    If Len(v) > 0 Then m_knowledgeNoLabel = v
    v = Trim(sec.GetValue("loadButtonLabel"))
    If Len(v) > 0 Then m_loadButtonLabel = v
    v = Trim(sec.GetValue("knowledgeNoHelp"))
    If Len(v) > 0 Then m_knowledgeNoHelp = v
    v = Trim(sec.GetValue("buttons"))
    If Len(v) > 0 Then m_buttons = v
    ' Phase R-2 F-2: per-header-field help lines (headerHelp_<id>).
    If Len(Trim(m_headerFields)) > 0 Then
        Dim hids() As String
        hids = Split(m_headerFields, ",")
        Dim hk As Long
        For hk = LBound(hids) To UBound(hids)
            Dim hid As String
            hid = Trim(hids(hk))
            If Len(hid) > 0 Then
                Dim hv As String
                hv = Trim(sec.GetValue("headerHelp_" & hid))
                If Len(hv) > 0 Then m_headerHelp(hid) = hv
            End If
        Next hk
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0785] clsUserFormRenderer.ApplyConfigFromStanza EXIT-OK"  ' [ADR-0100]
End Sub

' --- Internal: build form, populate, modal show, then dispose ---
Private Sub BuildAndShow(ByVal knowledgeData As Object)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0786] clsUserFormRenderer.BuildAndShow ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    ' Phase R-1-f/h: 3-stage fallback config load. default ?Esheet ?Eformat.
    InitFormConfig
    ApplyUserformStanza             ' sheet-level override
    ApplyUserformStanzaFromFormat   ' format-level override (highest priority)

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    If vbProj Is Nothing Then
        Debug.Print "[ERR] VBProject not available"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0787] clsUserFormRenderer.BuildAndShow EXIT-OK"  ' [ADR-0100]
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
    Set m_fieldRequiredByCtl = CreateObject("Scripting.Dictionary")
    Set m_dateFieldIndices = CreateObject("Scripting.Dictionary")
    Set m_comboItemsByCtl = CreateObject("Scripting.Dictionary")
    Set m_comboInitialByCtl = CreateObject("Scripting.Dictionary")
    Set m_displayToFmtId = CreateObject("Scripting.Dictionary")
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

    ' chrome 実測のため Width/Height を暫定設定
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

    ' scroll area Frame (header の下)
    Dim fr As Object
    Set fr = designer.Controls.Add(PROGID_FRAME, "frScroll", True)
    fr.caption = ""
    fr.top = m_headerHeightPx
    fr.left = 0
    fr.Width = m_formWidth
    ' V4 fix (2026-05-29) #1/#5 (revised): Frame の sunken/etched 枠 + 灰色背景を消す。
    ' V4-iter1 で BorderStyle=fmBorderStyleNone を追加したら M-05 register が
    ' 1 field しか描画されない regression が出たため (frame の inset+clipping 仕様変化)、
    ' BorderStyle は default (0=fmBorderStyleNone は変えず) + SpecialEffect=0(flat)
    ' + BackColor=白 で「枠見た目だけ」抑える方向に振る。
    On Error Resume Next
    fr.SpecialEffect = 0       ' fmSpecialEffectFlat (etched 枠を消す)
    Dim bcFrame As Long
    bcFrame = ParseHexColor(m_backColor)
    If bcFrame >= 0 Then fr.BackColor = bcFrame
    On Error GoTo ErrHandler
    Set m_scrollFrame = fr

    ' scroll content (frame-relative 座標で frScroll 内に生成)
    Dim y As Long
    Dim kOff As Long
    kOff = AddKnowledgeNoRow(fr)                        ' ナレッジ番号行(edit/view); M-05=0
    y = m_margin + kOff
    y = AddHeaderFields(fr, y)                          ' 予定番号 等 header fields
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

    ' 固定 form inside 高: 内容が収まればそのまま、超えたら cap で frame 内スクロール
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
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0788] clsUserFormRenderer.BuildAndShow EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ' Phase P fix: get the live form instance, populate ComboBox items at
    ' runtime (Designer.Controls.Add items don't survive), then Show.
    LogToSheet "BuildAndShow", "step 11 about to Show formName=" & m_dynFormName, "LOG-UF-SHOW-PRE"
    Dim ufInstance As Object
    Set ufInstance = VBA.UserForms.Add(m_dynFormName)
    PopulateComboBoxesOnInstance ufInstance

    ' V4 fix (2026-05-29) #M-05: register モードで frScroll 内の field 群が
    ' 1 件しか paint されない regression が発生 (子 agent dump で 11 件全構築済
    ' 確認)。frame の ScrollTop を 0 に reset + Repaint nudge で初期描画を強制。
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
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0789] clsUserFormRenderer.BuildAndShow EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ufInstance.Show vbModal
    LogToSheet "BuildAndShow", "step 12 closed returnId=" & m_returnId, "LOG-UF-SHOW-POST"

    ' Dispose dynamic form
    On Error Resume Next
    vbProj.VBComponents.Remove vbc
    On Error GoTo ErrHandler

    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0790] clsUserFormRenderer.BuildAndShow EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0791] clsUserFormRenderer.BuildAndShow EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] BuildAndShow: " & Err.Number & " " & Err.Description
    On Error Resume Next
    LogToSheet "BuildAndShow", "Show failed errNum=" & Err.Number & " desc=" & Err.Description & " formName=" & m_dynFormName, "LOG-UF-SHOW-ERR"
    On Error GoTo 0
End Sub

' Phase R-3-χ-4: find a control by name on the live form, searching direct
' controls then the frScroll frame's children (one level). field 系は frScroll 内。
Private Function FindCtlOnForm(ByVal uf As Object, ByVal ctlName As String) As Object
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0792] clsUserFormRenderer.FindCtlOnForm ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim c As Object
    Set c = uf.Controls(ctlName)
    If Not c Is Nothing Then
        Set FindCtlOnForm = c
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0793] clsUserFormRenderer.FindCtlOnForm EXIT-OK"  ' [ADR-0100]
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
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0794] clsUserFormRenderer.FindCtlOnForm EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        End If
    Next f
End Function

' Phase P fix: walk the live UserForm instance and add ComboBox items per
' the stashed m_comboItemsByCtl map. Then set initial selected value if any.
Private Sub PopulateComboBoxesOnInstance(ByVal uf As Object)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0795] clsUserFormRenderer.PopulateComboBoxesOnInstance ENTER"  ' [ADR-0100]
    On Error Resume Next
    If m_comboItemsByCtl Is Nothing Then Exit Sub
    Dim ctlName As Variant
    For Each ctlName In m_comboItemsByCtl.Keys
        Dim items As Variant
        items = m_comboItemsByCtl(CStr(ctlName))
        Dim cb As Object
        Set cb = FindCtlOnForm(uf, CStr(ctlName))   ' R-3-χ-4: frScroll frame 内も探索
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0796] clsUserFormRenderer.PopulateComboBoxesOnInstance EXIT-OK"  ' [ADR-0100]
End Sub

' Diagnostic logger - write to the active workbook's LOG sheet via clsLogger.
Private Sub LogToSheet(ByVal funcName As String, ByVal msg As String, ByVal logId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0797] clsUserFormRenderer.LogToSheet ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets("LOG")
    If Not lg Is Nothing Then
        lg.LogInfo "clsUserFormRenderer", funcName, msg, "", logId
    End If
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0798] clsUserFormRenderer.LogToSheet EXIT-OK"  ' [ADR-0100]
End Sub

Private Function ResolveFormTitle() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0799] clsUserFormRenderer.ResolveFormTitle ENTER"  ' [ADR-0100]
    ' Phase R-1-f: stanza caption override wins over mode-derived JP title.
    If Len(m_captionOverride) > 0 Then
        ResolveFormTitle = m_captionOverride
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0800] clsUserFormRenderer.ResolveFormTitle EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Select Case m_mode
        Case "register": ResolveFormTitle = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H767B) & ChrW(&H9332)
        Case "edit":     ResolveFormTitle = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H4FEE) & ChrW(&H6B63)
        Case "view":     ResolveFormTitle = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H8868) & ChrW(&H793A)
        Case "preview"
            ' プレビュー (V4 fix 2026-05-29: drop ": <formatId>" suffix per親指示#4)
            ResolveFormTitle = ChrW(&H30D7) & ChrW(&H30EC) & ChrW(&H30D3) & ChrW(&H30E5) & ChrW(&H30FC)
        Case Else:       ResolveFormTitle = "UserForm"
    End Select
End Function

Private Function LoadFormatFields(ByVal fmtId As String) As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0801] clsUserFormRenderer.LoadFormatFields ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If Len(fmtId) = 0 Then
        Set LoadFormatFields = Nothing
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0802] clsUserFormRenderer.LoadFormatFields EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim secs As Collection
    Set secs = modFormatLoader.LoadFormat(fmtId)
    If secs Is Nothing Then
        Set LoadFormatFields = Nothing
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0803] clsUserFormRenderer.LoadFormatFields EXIT-OK"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0804] clsUserFormRenderer.LoadFormatFields EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0805] clsUserFormRenderer.LoadFormatFields EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set LoadFormatFields = Nothing
End Function

Private Function ComputeRowsHeight(ByVal flds As Collection) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0806] clsUserFormRenderer.ComputeRowsHeight ENTER"  ' [ADR-0100]
    ' Includes subheader row "?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?i?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?b?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?W?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E????E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?e" + per-field row heights.
    Dim h As Long
    h = m_subheaderHeight
    If flds Is Nothing Then
        ComputeRowsHeight = h + m_rowHeightSingle
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0807] clsUserFormRenderer.ComputeRowsHeight EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To flds.count
        Set sec = flds.Item(i)
        If sec.GetValue("FieldType") = ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C) Then  ' ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?s
            ' Long multi-line for known "long" fields (?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??菇, ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?, ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E? ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?); others use short
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

' [USER-REQ 2026-06-09] Hardcoded heuristic removed (was: tall row for fields
' named 作業手順 / 事象 / 原因 / 詳細 / 内容). Format spec says Rows= is the
' single source of truth for multi-line height. If a format file wants a tall
' multi-line, it must specify Rows= explicitly.
Private Function IsLongMultilineField(ByVal fieldName As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0808] clsUserFormRenderer.IsLongMultilineField ENTER"  ' [ADR-0100]
    IsLongMultilineField = False
End Function

' Phase R-2 F-2: render a grey 9pt help line at (leftX, y) of given width.
' Returns the vertical space consumed (HELP_LINE_HEIGHT if text non-empty,
' else 0 so callers that skip empty help do not change row heights).
Private Function RenderHelpLine(ByVal designer As Object, ByVal ctlName As String, _
                                 ByVal y As Long, ByVal leftX As Long, ByVal w As Long, _
                                 ByVal text As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0810] clsUserFormRenderer.RenderHelpLine ENTER"  ' [ADR-0100]
    If Len(Trim(text)) = 0 Then
        RenderHelpLine = 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0811] clsUserFormRenderer.RenderHelpLine EXIT-OK"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0812] clsUserFormRenderer.AddKnowledgeNoRow ENTER"  ' [ADR-0100]
    If m_knowledgeNoRow <> "editable" And m_knowledgeNoRow <> "readonly" Then
        AddKnowledgeNoRow = 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0813] clsUserFormRenderer.AddKnowledgeNoRow EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    ' Phase R-3-χ-3 縦並び: label 行上・data 行(全幅)下。
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0814] clsUserFormRenderer.AddHeaderRow ENTER"  ' [ADR-0100]
    Dim lbl As Object
    Set lbl = designer.Controls.Add(PROGID_LABEL, "lblFormat", True)
    lbl.caption = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8)
    lbl.top = m_margin + m_headerYOffset
    lbl.left = m_margin
    lbl.Width = m_formWidth - m_margin * 2
    lbl.Height = VLABEL_H
    ApplyBaseFont lbl

    ' Phase R-3-χ-3 縦並び: format selector も label 上・data 全幅下。
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
        ' [BUG-B15 2026-06-11] items are display names (FormatName); the
        ' initial value must be the display name too, or the DropDownList
        ' rejects the raw FormatID and the combo shows blank after re-render.
        If Len(m_formatId) > 0 Then m_comboInitialByCtl("cboFormatId") = ResolveFormatIdToDisplay(m_formatId)
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

    ' Phase R-2 F-2 / R-3-χ-3: help line under the format selector (data 行直下)。
    Dim helpH As Long
    helpH = RenderHelpLine(designer, "lblFormatHelp", dataTopH + dataHH + 2, m_margin, ctlW, m_formatHelp)
    ' R-3-g1A: return the actual format-block bottom (control + formatHelp + gap)
    ' so the header-field row never overlaps the format help line (D-7 overlap fix).
    Dim hdrBottom As Long
    hdrBottom = dataTopH + dataHH
    If helpH > 0 Then hdrBottom = hdrBottom + 2 + helpH
    AddHeaderRow = hdrBottom + 6
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0815] clsUserFormRenderer.AddHeaderRow EXIT-OK"  ' [ADR-0100]
End Function

' Phase R-2 F-3: list available format ids (basenames of formats/*.txt).
' Returns a string array (possibly empty).
' 2026-06-11 fix: resolve a dropdown display name back to its FormatID.
' Unknown names (or direct IDs typed in textbox mode) pass through unchanged.
Public Function ResolveDisplayToFormatId(ByVal disp As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0816b] clsUserFormRenderer.ResolveDisplayToFormatId ENTER"  ' [ADR-0100]
    ResolveDisplayToFormatId = disp
    If m_displayToFmtId Is Nothing Then Exit Function
    If m_displayToFmtId.Exists(disp) Then ResolveDisplayToFormatId = CStr(m_displayToFmtId(disp))
End Function

' [BUG-B15 2026-06-11] Reverse of ResolveDisplayToFormatId: FormatID ->
' display name (FormatName). Pass-through when unknown.
Public Function ResolveFormatIdToDisplay(ByVal fmtId As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0816c] clsUserFormRenderer.ResolveFormatIdToDisplay ENTER"  ' [ADR-0100]
    ResolveFormatIdToDisplay = fmtId
    If m_displayToFmtId Is Nothing Then Exit Function
    Dim k As Variant
    For Each k In m_displayToFmtId.Keys
        If CStr(m_displayToFmtId(k)) = fmtId Then
            ResolveFormatIdToDisplay = CStr(k)
            Exit Function
        End If
    Next k
End Function

Private Function ListFormatIds() As Variant
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0816] clsUserFormRenderer.ListFormatIds ENTER"  ' [ADR-0100]
    ' [USER-REQ 2026-06-09] Show FormatName in dropdown (not FormatID).
    ' Disabled formats are filtered out (LoadFormatList already excludes them).
    On Error GoTo Fail
    Dim col As Collection
    Set col = modFormatLoader.LoadFormatList()
    If col Is Nothing Then GoTo Fail
    If col.count = 0 Then
        ListFormatIds = Array()
        Exit Function
    End If
    Dim arr() As String
    ReDim arr(0 To col.count - 1)
    Dim i As Long
    For i = 1 To col.count
        Dim ent As Object
        Set ent = col.Item(i)
        Dim nm As String
        nm = ""
        On Error Resume Next
        If ent.Exists("Description") Then nm = CStr(ent("Description"))
        If Len(nm) = 0 And ent.Exists("FormatID") Then nm = CStr(ent("FormatID"))
        ' 2026-06-11 fix: the dropdown shows FormatName but downstream
        ' (OnFormatChange / PersistFromActiveForm) needs the FormatID.
        ' Keep a display->ID reverse map (last-one-wins on duplicate names).
        If ent.Exists("FormatID") Then
            If m_displayToFmtId Is Nothing Then Set m_displayToFmtId = CreateObject("Scripting.Dictionary")
            m_displayToFmtId(nm) = CStr(ent("FormatID"))
        End If
        On Error GoTo Fail
        arr(i - 1) = nm
    Next i
    ListFormatIds = arr
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0818] clsUserFormRenderer.ListFormatIds EXIT-OK"  ' [ADR-0100]
    Exit Function
Fail:
    ListFormatIds = Array()
End Function

' Phase R-2 F-1: number of configured header fields.
Private Function HeaderFieldCount() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0819] clsUserFormRenderer.HeaderFieldCount ENTER"  ' [ADR-0100]
    Dim s As String
    s = Trim(m_headerFields)
    If Len(s) = 0 Then
        HeaderFieldCount = 0
    Else
        HeaderFieldCount = UBound(Split(s, ",")) - LBound(Split(s, ",")) + 1
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0820] clsUserFormRenderer.HeaderFieldCount EXIT-OK"  ' [ADR-0100]
End Function

' Phase R-2 F-1: map a logical header-field id to its JP label.
Private Function ResolveHeaderFieldLabel(ByVal id As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0821] clsUserFormRenderer.ResolveHeaderFieldLabel ENTER"  ' [ADR-0100]
    Select Case LCase(Trim(id))
        Case "knowledgeid": ResolveHeaderFieldLabel = ChrW(&H4E88) & ChrW(&H5B9A) & ChrW(&H756A) & ChrW(&H53F7)        ' 予定番号
        Case "createdat":   ResolveHeaderFieldLabel = ChrW(&H767B) & ChrW(&H9332) & ChrW(&H65E5) & ChrW(&H4ED8)        ' 登録日時
        Case "updatedat":   ResolveHeaderFieldLabel = ChrW(&H66F4) & ChrW(&H65B0) & ChrW(&H65E5) & ChrW(&H4ED8)        ' 更新日時
        Case Else:          ResolveHeaderFieldLabel = id
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0822] clsUserFormRenderer.ResolveHeaderFieldLabel EXIT-OK"  ' [ADR-0100]
End Function

' Phase R-2 F-1: resolve a header field's current value from the loaded
' knowledge data (edit/view); register mode returns blank.
Private Function ResolveHeaderFieldValue(ByVal id As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0823] clsUserFormRenderer.ResolveHeaderFieldValue ENTER"  ' [ADR-0100]
    Dim key As String
    Select Case LCase(Trim(id))
        Case "knowledgeid"
            ResolveHeaderFieldValue = m_knowledgeId
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0824] clsUserFormRenderer.ResolveHeaderFieldValue EXIT-OK"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0825] clsUserFormRenderer.AddHeaderFields ENTER"  ' [ADR-0100]
    Dim s As String
    s = Trim(m_headerFields)
    If Len(s) = 0 Then
        AddHeaderFields = yStart
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0826] clsUserFormRenderer.AddHeaderFields EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim ids() As String
    ids = Split(s, ",")
    Dim y As Long
    y = yStart
    Dim k As Long
    For k = LBound(ids) To UBound(ids)
        Dim id As String
        id = Trim(ids(k))
        If Len(id) > 0 Then
            Dim idxStr As String
            idxStr = Format$(k + 1, "000")
            Dim lbl As Object
            Set lbl = designer.Controls.Add(PROGID_LABEL, "hdrlbl_" & idxStr, True)
            ' Phase R-3-χ-3 縦並び: header field も label 上・data 全幅下。
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
            ctl.Locked = (m_mode = "view")  ' [USER-REQ 2026-06-09] preview is trial-input per spec
            ApplyBaseFont ctl

            Dim hfExtra As Long
            hfExtra = 0
            ' Phase R-2 F-2: per-header-field help line (data 行直下)。
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

' Apply メイリオ base font (Phase R-1-c, spec ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?1)
Private Sub ApplyBaseFont(ByVal ctl As Object)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0827] clsUserFormRenderer.ApplyBaseFont ENTER"  ' [ADR-0100]
    On Error Resume Next
    ctl.Font.Name = m_fontName
    ctl.Font.Size = m_fontSize
    On Error GoTo 0
End Sub

' Subheader "?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?i?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?b?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?W?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E????E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?e" row (between header and dynamic fields, per mock)
' V3 fix (2026-05-29): mock navy title bar at form top.
' OS title bar (VBA-uncontrollable) stays; this paints a navy band
' just below it inside client area with the form's JP title.
Private Sub AddTitleBar(ByVal designer As Object)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0828] clsUserFormRenderer.AddTitleBar ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0829] clsUserFormRenderer.AddTitleBar EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub AddSubheaderRow(ByVal designer As Object, ByVal y As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0830] clsUserFormRenderer.AddSubheaderRow ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim lbl As Object
    Set lbl = designer.Controls.Add(PROGID_LABEL, "lblSubheader", True)
    ' ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?i?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?b?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?W?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E????E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?e
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0831] clsUserFormRenderer.AddSubheaderRow EXIT-OK"  ' [ADR-0100]
End Sub

Private Function AddFieldRow(ByVal designer As Object, _
                              ByVal sec As ClsStanzaSection, _
                              ByVal y As Long, _
                              ByVal knowledgeData As Object) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0832] clsUserFormRenderer.AddFieldRow ENTER"  ' [ADR-0100]
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
    ' Phase R-3-χ-3 縦並び: label は上の行・全幅(badge 分を控除)。
    lbl.caption = fieldName
    lbl.top = y
    lbl.left = m_margin
    lbl.Width = m_formWidth - m_margin * 2 - m_badgeWidth - 4
    lbl.Height = VLABEL_H
    ApplyBaseFont lbl

    ' Phase R-1-c: ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?K?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?{ badge for required fields (FieldRequired = TRUE / true)
    Dim isReq As Boolean
    Dim reqStr As String
    ' Phase R-1-j: format files use "Required"; fall back to legacy "FieldRequired".
    reqStr = LCase(Trim(sec.GetValue("Required")))
    If Len(reqStr) = 0 Then reqStr = LCase(Trim(sec.GetValue("FieldRequired")))
    isReq = (reqStr = "true" Or reqStr = "1" Or reqStr = "yes" Or reqStr = "tr")
    If isReq Then
        Dim badge As Object
        Set badge = designer.Controls.Add(PROGID_LABEL, "bdg_" & idxStr, True)
        badge.caption = ChrW(&H5FC5) & ChrW(&H9808)   ' ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?K?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?{
        ' 縦並び: badge は label 行の右端に配置
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
    m_fieldRequiredByCtl(ctlName) = isReq   ' [BUG-B16] remember required flag for runtime validation

    Select Case fieldType
        Case ChrW(&H5358) & ChrW(&H4E00) & ChrW(&H884C)  ' ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?P?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?s
            Set ctl = designer.Controls.Add(PROGID_TEXTBOX, ctlName, True)
            ctl.MultiLine = False
            ctl.Text = curVal
        Case ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C)  ' ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?s
            Set ctl = designer.Controls.Add(PROGID_TEXTBOX, ctlName, True)
            ctl.MultiLine = True
            ctl.WordWrap = True
            ctl.EnterKeyBehavior = True
            ' [USERINPUT 2026-06-06] per-field Scroll value from format stanza
            Dim scrollVal As String
            scrollVal = LCase(Trim(sec.GetValue("Scroll")))
            If scrollVal = "true" Or scrollVal = "1" Or scrollVal = "yes" Then
                ctl.ScrollBars = 2   ' vertical
            ElseIf scrollVal = "false" Or scrollVal = "0" Or scrollVal = "no" Then
                ctl.ScrollBars = 0   ' none
            Else
                ' fallback: global stanza setting
                Select Case LCase(m_multiLineScrollBars)
                    Case "vertical": ctl.ScrollBars = 2
                    Case "horizontal": ctl.ScrollBars = 1
                    Case "both": ctl.ScrollBars = 3
                    Case "none": ctl.ScrollBars = 0
                    Case Else: ctl.ScrollBars = 2
                End Select
            End If
            ctl.Text = curVal
            ' Long multi-line for ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??菇/?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?/?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?/?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E???/?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?e; others short
            ' [USER-REQ 2026-06-09] No hardcoded fake defaults. Read Rows from
            ' the format file as-is. If the format does not specify Rows, render
            ' at single-line height. Format files that need a tall textbox must
            ' explicitly set Rows= in the [FIELD] stanza.
            ' Also accept FieldLineCount as the spec-documented alias.
            Dim rowsN As Long
            rowsN = 0
            On Error Resume Next
            rowsN = CLng(Val(Trim(sec.GetValue("Rows"))))
            If rowsN < 1 Then rowsN = CLng(Val(Trim(sec.GetValue("FieldLineCount"))))
            On Error GoTo 0
            If rowsN >= 1 Then
                ' [USER-REQ 2026-06-09] Rows=1 should render as 1 visible row, not 2.
                ' Previously: single + (rowsN * 20) made Rows=N render as N+1 rows.
                rowH = m_rowHeightSingle + ((rowsN - 1) * 20)
            Else
                rowH = m_rowHeightSingle  ' Rows unspecified -> 1-line height
            End If
        Case ChrW(&H65E5) & ChrW(&H4ED8), ChrW(&H65E5) & ChrW(&H6642)  ' [Q2.a 2026-06-11] nichiji = alias of date ' ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E? - mark for Change-event validation
            Set ctl = designer.Controls.Add(PROGID_TEXTBOX, ctlName, True)
            ctl.MultiLine = False
            ctl.Text = curVal
            m_dateFieldIndices(idxStr) = True
        Case ChrW(&H9078) & ChrW(&H629E)  ' ?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?I?E?E?E?E?E?E?E?E?E?E?E?E?E?E?E??E?E?E?E?E?E?E?E?E?E?E?E?E?E?E?
            Set ctl = designer.Controls.Add(PROGID_COMBOBOX, ctlName, True)
            ctl.Style = 2   ' fmStyleDropDownList
            Dim opts As String
            opts = sec.GetValue("DropdownOptions")
            If Len(opts) > 0 Then
                ' 2026-06-11 fix: accept comma-separated values too (M-03 spec input);
                ' canonical file separator is pipe.
                opts = Replace(opts, ",", "|")
                Dim parts() As String
                parts = Split(opts, "|")
                Dim k As Long
                For k = LBound(parts) To UBound(parts)
                    parts(k) = Trim(parts(k))
                Next k
                m_comboItemsByCtl(ctlName) = parts
            End If
            If Len(curVal) > 0 Then m_comboInitialByCtl(ctlName) = curVal
        Case Else
            Set ctl = designer.Controls.Add(PROGID_TEXTBOX, ctlName, True)
            ctl.MultiLine = False
            ctl.Text = curVal
    End Select

    ' Phase R-3-χ-3 縦並び: data は label 行の下・全幅。
    ' dataHeight = 行ピッチ(rowH) - labelZone(VLABEL_H+VLABEL_GAP) - VROW_SPACING。
    Dim dataTop As Long, dataH As Long
    dataTop = y + VLABEL_H + VLABEL_GAP
    dataH = rowH - VLABEL_H - VLABEL_GAP - VROW_SPACING
    ctl.top = dataTop
    ctl.left = m_margin
    ctl.Width = m_formWidth - m_margin * 2
    ctl.Height = dataH
    ApplyBaseFont ctl

    ' Apply mode-based locking (view と preview は表示専用 = readonly)
    If m_mode = "view" Then  ' [USER-REQ 2026-06-09] preview is trial-input
        On Error Resume Next
        ctl.Locked = True
        ctl.BackColor = RGB(240, 240, 240)
        On Error GoTo 0
    End If

    ' Phase R-2 F-4: placeholder grey text in empty TextBoxes (not view mode,
    ' not ComboBox). Stored so InjectFormCode can emit Enter/Exit handlers and
    ' the persistence layer can treat a still-placeholder field as empty.
    ' Phase R-3-χ-2: preview (M-04) は実データを持たず placeholder(記入例)のみ表示。
    ' readonly のため focus-clear ハンドラは生成せず static に出す。
    If TypeName(ctl) = "TextBox" And m_mode <> "view" And Len(curVal) = 0 Then
        Dim ph As String
        ph = Trim(sec.GetValue("fieldPlaceholder"))
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
    fHelp = Trim(sec.GetValue("fieldHelp"))
    Dim extraH As Long
    extraH = RenderHelpLine(designer, "fldhelp_" & idxStr, dataTop + dataH + 2, m_margin, _
                            m_formWidth - m_margin * 2, fHelp)

    AddFieldRow = rowH + extraH
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0833] clsUserFormRenderer.AddFieldRow EXIT-OK"  ' [ADR-0100]
End Function

' Replace characters not valid in VBA identifiers with underscore.
Private Function SafeName(ByVal s As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0834] clsUserFormRenderer.SafeName ENTER"  ' [ADR-0100]
    Dim i As Long
    Dim out As String
    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid(s, i, 1)
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0835] clsUserFormRenderer.SafeName EXIT-OK"  ' [ADR-0100]
End Function

Private Sub AddButtonBar(ByVal designer As Object, ByVal y As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0836] clsUserFormRenderer.AddButtonBar ENTER"  ' [ADR-0100]
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
            ' Phase R-3-χ-2: M-04 プレビューは「閉じる」単独 (M-09 と同 btnClose)
            names = Split("btnClose", "|")
            labels = Split(ChrW(&H9589) & ChrW(&H3058) & ChrW(&H308B), "|")
            kinds = Split("p", "|")
        Case Else
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0837] clsUserFormRenderer.AddButtonBar EXIT-OK"  ' [ADR-0100]
            Exit Sub
    End Select

    ' Phase R-1-i: button help text per mock. Help label below each button.
    Dim helpTexts() As String
    Select Case m_mode
        Case "register"
            ' クリア / 登録
            helpTexts = Split(ChrW(&H5165) & ChrW(&H529B) & ChrW(&H6B04) & ChrW(&H3092) & ChrW(&H7A7A) & ChrW(&H306B) & ChrW(&H623B) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & "|" & ChrW(&H5185) & ChrW(&H5BB9) & ChrW(&H3092) & ChrW(&H4FDD) & ChrW(&H5B58) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H767B) & ChrW(&H9332) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059), "|")
        Case "edit"
            ' 削除 / 更新
            helpTexts = Split(ChrW(&H78BA) & ChrW(&H8A8D) & ChrW(&H306E) & ChrW(&H3046) & ChrW(&H3048) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & "|" & ChrW(&H5185) & ChrW(&H5BB9) & ChrW(&H3092) & ChrW(&H4E0A) & ChrW(&H66F8) & ChrW(&H304D) & ChrW(&H4FDD) & ChrW(&H5B58) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059), "|")
        Case "view"
            ' 編雁E/ 削除 / 閉じめE
            helpTexts = Split("|" & "|" & ChrW(&H691C) & ChrW(&H7D22) & ChrW(&H753B) & ChrW(&H9762) & ChrW(&H306B) & ChrW(&H623B) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H3059), "|")
        Case "preview"
            ' プレビューを閉じます
            helpTexts = Split(ChrW(&H30D7) & ChrW(&H30EC) & ChrW(&H30D3) & ChrW(&H30E5) & ChrW(&H30FC) & ChrW(&H3092) & ChrW(&H9589) & ChrW(&H3058) & ChrW(&H307E) & ChrW(&H3059), "|")
        Case Else
            helpTexts = Split("|", "|")
    End Select

    Dim n As Long
    ' R-3-e (2026-05-28): config-driven button filter. [USERFORM] buttons (csv)
    ' keeps only the listed button ids from the mode-default set (no hardcode).
    If Len(Trim(m_buttons)) > 0 Then
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
                If Trim(wantB(wj)) = names(bi) Then keepB = True
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
    ' [B26 2026-06-11] maximize toggle on every form (user request).
    ' Added after the [USERFORM] buttons filter so seeds cannot drop it.
    names = Split("btnMaximize|" & Join(names, "|"), "|")
    labels = Split(ChrW(&H6700) & ChrW(&H5927) & ChrW(&H5316) & "|" & Join(labels, "|"), "|")
    kinds = Split("s|" & Join(kinds, "|"), "|")
    helpTexts = Split(ChrW(&H753B) & ChrW(&H9762) & ChrW(&H3092) & ChrW(&H6700) & ChrW(&H5927) & ChrW(&H5316) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & "|" & Join(helpTexts, "|"), "|")
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
        ' V5 fix (2026-05-30) BUG-2: Default/Cancel 明示 set。primary(p) を Enter で発火する
        ' Default に、Clear/Close を Esc で発火する Cancel に。これで Tab で field 移動中の
        ' 誤 Enter が「クリア」を発火して field 全消失する UX 事故を防ぐ。
        On Error Resume Next
        Select Case kinds(i)
            Case "p"
                btn.Default = True
                btn.Cancel = False
            Case Else
                btn.Default = False
                ' btnClear / btnClose を Esc キーで発火させる
                If names(i) = "btnClear" Or names(i) = "btnClose" Then
                    btn.Cancel = True
                Else
                    btn.Cancel = False
                End If
        End Select
        On Error GoTo 0

        ' R-3-g1A (2026-05-28): per-button help (mock 準拠、案c 統合補助文を撤回)。
        ' buttonWidth(113) 幅内に収め右みきれ/重なりを回避。
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0838] clsUserFormRenderer.InjectFormCode ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim cm As Object
    Set cm = vbc.CodeModule
    Dim s As String
    s = "Option Explicit" & vbCrLf & vbCrLf
    ' [BUG-B12 2026-06-11] spec M-05: show a completion message and KEEP the
    ' form open after register (was: silent Unload Me). OnRegisterV2 shows
    ' the message itself (headless-guarded); failures also keep the form.
    s = s & "Private Sub btnRegister_Click()" & vbCrLf
    s = s & "    Application.Run ""modUserFormCallback.OnRegisterV2""" & vbCrLf
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
    ' [B26 2026-06-11] maximize/restore toggle. Resizes the form to the
    ' Excel application bounds, grows frScroll, and shifts the bottom
    ' button bar (buttons + help labels) by the delta.
    s = s & vbCrLf
    s = s & "Private Sub btnMaximize_Click()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Static origW As Double, origH As Double, isMax As Boolean" & vbCrLf
    s = s & "    Dim dW As Double, dH As Double, thr As Double" & vbCrLf
    s = s & "    Dim frS As Object" & vbCrLf
    s = s & "    Set frS = Me.Controls(""frScroll"")" & vbCrLf
    s = s & "    If frS Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    thr = frS.Top + frS.Height - 5" & vbCrLf
    s = s & "    If Not isMax Then" & vbCrLf
    s = s & "        origW = Me.Width: origH = Me.Height" & vbCrLf
    s = s & "        dW = Application.Width - Me.Width: dH = Application.Height - Me.Height" & vbCrLf
    s = s & "        If dW < 0 Then dW = 0" & vbCrLf
    s = s & "        If dH < 0 Then dH = 0" & vbCrLf
    s = s & "        Me.Move Application.Left, Application.Top, origW + dW, origH + dH" & vbCrLf
    s = s & "        isMax = True" & vbCrLf
    s = s & "        Me.Controls(""btnMaximize"").Caption = " & Chr(34) & ChrW(&H5143) & ChrW(&H306B) & ChrW(&H623B) & ChrW(&H3059) & Chr(34) & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        dW = origW - Me.Width: dH = origH - Me.Height" & vbCrLf
    s = s & "        Me.Move Application.Left + (Application.Width - origW) / 2, Application.Top + (Application.Height - origH) / 2, origW, origH" & vbCrLf
    s = s & "        isMax = False" & vbCrLf
    s = s & "        Me.Controls(""btnMaximize"").Caption = " & Chr(34) & ChrW(&H6700) & ChrW(&H5927) & ChrW(&H5316) & Chr(34) & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    frS.Width = frS.Width + dW: frS.Height = frS.Height + dH" & vbCrLf
    ' [B26v2 2026-06-12] widen the title bar and field controls too (user report)
    s = s & "    Dim tbar As Object" & vbCrLf
    s = s & "    Set tbar = Nothing" & vbCrLf
    s = s & "    Set tbar = Me.Controls(""lblTitleBar"")" & vbCrLf
    s = s & "    If Not tbar Is Nothing Then tbar.Width = tbar.Width + dW" & vbCrLf
    s = s & "    Dim c2 As Object" & vbCrLf
    s = s & "    For Each c2 In frS.Controls" & vbCrLf
    s = s & "        If TypeName(c2) = ""TextBox"" Or TypeName(c2) = ""ComboBox"" Then c2.Width = c2.Width + dW" & vbCrLf
    s = s & "    Next c2" & vbCrLf
    s = s & "    Dim c As Object" & vbCrLf
    s = s & "    For Each c In Me.Controls" & vbCrLf
    s = s & "        If Not (c Is frS) Then" & vbCrLf
    s = s & "            Dim isTop As Boolean" & vbCrLf
    s = s & "            isTop = False" & vbCrLf
    s = s & "            isTop = (c.Parent Is Me)" & vbCrLf
    s = s & "            If isTop And c.Top >= thr Then" & vbCrLf
    s = s & "                c.Top = c.Top + dH: c.Left = c.Left + dW" & vbCrLf
    s = s & "            ElseIf isTop And (TypeName(c) = ""TextBox"" Or TypeName(c) = ""ComboBox"") Then" & vbCrLf
    s = s & "                c.Width = c.Width + dW  '' [B26v3] header-area fields" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "    Me.Repaint" & vbCrLf
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
    ' Phase R-3-χ-5: format 行が無い mode では cboFormatId が存在しないため inject しない。
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
            s = s & "    v = Trim(CStr(Me." & cn & ".Text))" & vbCrLf
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
            s = s & "    If Len(Trim(CStr(Me." & pn & ".Text))) = 0 Then" & vbCrLf
            s = s & "        Me." & pn & ".Text = """ & phTxt & """" & vbCrLf
            s = s & "        Me." & pn & ".ForeColor = RGB(128, 128, 128)" & vbCrLf
            s = s & "        Me." & pn & ".Tag = """ & PLACEHOLDER_TAG & """" & vbCrLf
            s = s & "    End If" & vbCrLf
            s = s & "End Sub" & vbCrLf
        Next pk
    End If
    cm.AddFromString s
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0839] clsUserFormRenderer.InjectFormCode EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0840] clsUserFormRenderer.InjectFormCode EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] InjectFormCode: " & Err.Number & " " & Err.Description
End Sub

' Phase R-1-i matrix test harness: load config (default + sheet + format) and
' return current m_* as semicolon-separated key=value string. Does NOT build
' the UserForm or call .Show, so it's safely callable from PowerShell COM.
Public Function MeasureFormConfig(ByVal xlsmName As String, ByVal mode As String, _
                                   ByVal knowledgeId As String, ByVal formatId As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0841] clsUserFormRenderer.MeasureFormConfig ENTER"  ' [ADR-0100]
    m_xlsmName = xlsmName
    m_mode = LCase(mode)
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0842] clsUserFormRenderer.MeasureFormConfig EXIT-OK"  ' [ADR-0100]
End Function

' Phase R-1-j modeless show. Builds the form then calls Show vbModeless.
' Returns immediately so PowerShell can capture a screenshot. Caller must
' invoke CloseModelessForm afterward.
Public Function ShowFormModeless(ByVal xlsmName As String, ByVal mode As String, _
                                  ByVal knowledgeId As String, ByVal formatId As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0843] clsUserFormRenderer.ShowFormModeless ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    m_xlsmName = xlsmName
    m_mode = LCase(mode)
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0844] clsUserFormRenderer.ShowFormModeless EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0845] clsUserFormRenderer.ShowFormModeless EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ShowFormModeless = "ERR:" & Err.Number & ":" & Err.Description
End Function

Public Sub CloseModelessForm()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0846] clsUserFormRenderer.CloseModelessForm ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0847] clsUserFormRenderer.CloseModelessForm EXIT-OK"  ' [ADR-0100]
End Sub

' Phase R-1-j path verify: build the form like BuildAndShow but skip Show,
' dump designer.Controls + resolved m_* values + fmtFields to a file, then
' remove the VBComponent. PowerShell parses the file to verify the
' stanza -> loader -> control path.
Public Function TestR1j_DumpForm(ByVal xlsmName As String, ByVal mode As String, _
                                  ByVal knowledgeId As String, ByVal formatId As String, _
                                  ByVal dumpPath As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0848] clsUserFormRenderer.TestR1j_DumpForm ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    m_xlsmName = xlsmName
    m_mode = LCase(mode)
    m_knowledgeId = knowledgeId
    m_formatId = formatId
    m_returnId = ""
    m_dumpToFile = dumpPath
    modUserFormCallback.SetRenderer Me
    Dim knowledgeData As Object
    Set knowledgeData = Nothing
    BuildAndShow knowledgeData
    TestR1j_DumpForm = "OK:" & dumpPath
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0849] clsUserFormRenderer.TestR1j_DumpForm EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0850] clsUserFormRenderer.TestR1j_DumpForm EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    TestR1j_DumpForm = "ERR:" & Err.Number & ":" & Err.Description
End Function

' Dump designer.Controls + resolved layout to a UTF-8 text file.
Private Sub DumpFormToFile(ByVal vbc As Object, ByVal designer As Object, _
                            ByVal fmtFields As Collection, ByVal formH As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0851] clsUserFormRenderer.DumpFormToFile ENTER"  ' [ADR-0100]
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

    ' R-3-χ-4: designer.Controls は frame 子も flatten して含む(frame-relative 座標)ため、
    ' frame 子は [CONTROLS] から除外し [FRAMECONTROLS] にのみ出す。
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

    ' Phase R-3-χ-4: frScroll frame geometry + その子 controls (frame-relative 座標)
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0852] clsUserFormRenderer.DumpFormToFile EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0853] clsUserFormRenderer.DumpFormToFile EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] DumpFormToFile: " & Err.Number & " " & Err.Description
End Sub

' --- Callback target setters (called by modUserFormCallback) ---
Public Sub SetReturnId(ByVal v As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0854] clsUserFormRenderer.SetReturnId ENTER"  ' [ADR-0100]
    m_returnId = v
End Sub

' Phase R-2 F-3: format dropdown changed. Only re-render when the new format
' actually differs from the current one (guards against the Change event that
' fires when the initial value is set programmatically).
' R-3-a: btnLoad -> set the knowledge number, flag a reload; ShowForm loop refetches.
Public Sub RequestLoad(ByVal knowledgeNo As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0855] clsUserFormRenderer.RequestLoad ENTER"  ' [ADR-0100]
    m_knowledgeId = Trim(knowledgeNo)
    m_loadRequested = True
    m_reformatRequested = True
End Sub

Public Sub RequestReformat(ByVal newFormatId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0856] clsUserFormRenderer.RequestReformat ENTER"  ' [ADR-0100]
    If Len(newFormatId) = 0 Then Exit Sub
    If newFormatId = m_formatId Then Exit Sub
    m_formatId = newFormatId
    m_reformatRequested = True
End Sub

' Phase P E2E only: pre-seed m_formatId before ShowForm so the dynamic
' form has fields even in register mode (production code reads formatId
' from a sheet selector that is not yet wired).
Public Sub TestSeedFormatId(ByVal fmtId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0857] clsUserFormRenderer.TestSeedFormatId ENTER"  ' [ADR-0100]
    m_formatId = fmtId
End Sub
Public Function GetDynFormName() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0858] clsUserFormRenderer.GetDynFormName ENTER"  ' [ADR-0100]
    GetDynFormName = m_dynFormName
End Function
' Phase O-2: expose ctlName -> fieldName mapping for the callback bridge.
' [BUG-B16 2026-06-11] True when the control maps to a required field.
Public Function IsRequiredCtl(ByVal ctlName As String) As Boolean
    IsRequiredCtl = False
    If m_fieldRequiredByCtl Is Nothing Then Exit Function
    If m_fieldRequiredByCtl.Exists(ctlName) Then IsRequiredCtl = CBool(m_fieldRequiredByCtl(ctlName))
End Function

Public Function GetFieldNameForCtl(ByVal ctlName As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0859] clsUserFormRenderer.GetFieldNameForCtl ENTER"  ' [ADR-0100]
    If m_fieldNamesByIdx Is Nothing Then Exit Function
    If m_fieldNamesByIdx.Exists(ctlName) Then
        GetFieldNameForCtl = CStr(m_fieldNamesByIdx(ctlName))
    End If
End Function
Public Function GetMode() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0860] clsUserFormRenderer.GetMode ENTER"  ' [ADR-0100]
    GetMode = m_mode
End Function
Public Function GetKnowledgeId() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0861] clsUserFormRenderer.GetKnowledgeId ENTER"  ' [ADR-0100]
    GetKnowledgeId = m_knowledgeId
End Function
Public Function GetFormatId() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0862] clsUserFormRenderer.GetFormatId ENTER"  ' [ADR-0100]
    GetFormatId = m_formatId
End Function

' --- IScreenRenderer stubs (kept for compat) ---
Private Sub IScreenRenderer_BindSheet(ByVal sheetName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0863] clsUserFormRenderer.IScreenRenderer_BindSheet ENTER"  ' [ADR-0100]
End Sub
Private Sub IScreenRenderer_ClearScreen()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0864] clsUserFormRenderer.IScreenRenderer_ClearScreen ENTER"  ' [ADR-0100]
End Sub
Private Sub IScreenRenderer_ApplyFromStanza(ByVal xlsmName As String, ByVal screenId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0865] clsUserFormRenderer.IScreenRenderer_ApplyFromStanza ENTER"  ' [ADR-0100]
End Sub
Private Sub IScreenRenderer_ShowSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0866] clsUserFormRenderer.IScreenRenderer_ShowSheet ENTER"  ' [ADR-0100]
End Sub
Private Sub IScreenRenderer_HideSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0867] clsUserFormRenderer.IScreenRenderer_HideSheet ENTER"  ' [ADR-0100]
End Sub
Private Sub IScreenRenderer_ActivateSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0868] clsUserFormRenderer.IScreenRenderer_ActivateSheet ENTER"  ' [ADR-0100]
End Sub
Private Sub IScreenRenderer_ProtectSheet(ByVal level As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0869] clsUserFormRenderer.IScreenRenderer_ProtectSheet ENTER"  ' [ADR-0100]
End Sub
Private Sub IScreenRenderer_UnprotectSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0870] clsUserFormRenderer.IScreenRenderer_UnprotectSheet ENTER"  ' [ADR-0100]
End Sub

' ADR-0006/0090/0094 JP literal removal:
Private Property Get DEFAULT_FONT_NAME() As String
    DEFAULT_FONT_NAME = ChrW(&H30E1) & ChrW(&H30A4) & ChrW(&H30EA) & ChrW(&H30AA)
End Property

' [B32 2026-06-12] date-type lookup for register/update validation
Public Function IsDateCtl(ByVal ctlName As String) As Boolean
    IsDateCtl = False
    On Error Resume Next
    If m_dateFieldIndices Is Nothing Then Exit Function
    Dim kk As Variant
    For Each kk In m_dateFieldIndices.Keys
    If "ctl_" & CStr(kk) = ctlName Then
        IsDateCtl = True
        Exit Function
    End If
    Next kk
End Function
```
