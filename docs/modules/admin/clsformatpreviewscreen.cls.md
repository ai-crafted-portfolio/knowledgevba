---
title: clsFormatPreviewScreen.cls
description: clsFormatPreviewScreen.cls のソースコード（コピペ用）
---

# clsFormatPreviewScreen.cls

**配置先**: `管理.xlsm` 用の VBA モジュール
**種類**: クラスモジュール

---

## ファイルとして保存

メモ帳（または任意のテキストエディタ）に下のソースコード全文を貼り付け、**`clsFormatPreviewScreen.cls`** という名前で `installer\vba_modules\admin\` 配下に保存してください。文字コードは ANSI（Shift-JIS）、改行は CRLF にしてください。

---

## ソースコード


```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormatPreviewScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsFormatPreviewScreen (画面層 - M-04 フォーマットプレビュー)
' 概要:   M-04 画面の構築・再描画。spec を modScreenRender に委譲。
'         プレビューグリッドの動的再描画は [GRID_FROM_FORMAT] +
'         modUILoader.ApplyGridFromFormat 経由 (v2.2 task 7.A)。
' 依存先: IScreenRenderer, clsScreenSpec, clsFieldSpec, modScreenRender,
'         modUILoader, ClsStanzaSection
' 備考:   v2.1 (2026-05-20) で publish 投入。buildtest archive 旧版を
'         踏襲しつつ logging 規約 v1 準拠の新規 emit を追加。
'         既存 LOG-M04-SCREENCLS-SETUP-* 系は §3.3 後方互換維持で残置。
'         v2.2 task 7.A (2026-05-23): RenderPreview を追加。M-03 で
'         編集中のフォーマット (M-03!C3) が変わった／プレビュー画面へ
'         遷移した際に、画面側からプレビューグリッドを再描画する
'         (v2_ui_stanza_schema.md §3.13 / §4 注、OP-2 確定)。初回描画は
'         Setup -> RenderStandardScreen -> ApplyUiToSheet 内で
'         [GRID_FROM_FORMAT] が適用されるため、本クラスの RenderPreview
'         は「再描画」専用。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

' --- v2.2 task 7.A: M-04 プレビュー動的再描画用の定数 ---
' UI スタンザ配置 xlsm 名 (管理.xlsm = "Kanri"。modEntryFormat の
' M-03 読込と同じ命名規約。screenId は m_spec.ScreenId から取得)
Private Const UI_XLSM_NAME_LEGACY As String = "Kanri" ' deprecated, ADR-0089 (kept for ref)
Private Const SECTION_GRID_FROM_FORMAT As String = "GRID_FROM_FORMAT"
Private Const PREVIEW_CLEAR_RANGE As String = "A6:L39"
Private Function UiXlsmName() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0038] clsFormatPreviewScreen.UiXlsmName ENTER"  ' [ADR-0100]
    ' ADR-0089: CJK ui_seed dir name (U+7BA1 U+7406)
    UiXlsmName = ChrW(&H7BA1) & ChrW(&H7406)
End Function
' フォーマット駆動プレビューグリッドのセクション名
' 再描画前にクリアする動的描画域。M-04.txt の [GRID_FROM_FORMAT]
' StartCell=A6 から、最下部 [NOTE] (A40:J45) の直上 39 行まで。
' 列は [COLUMN]/[HEADER] が定義する A:L を含める。
' RenderPreview のログ ID (既存 LOG-M04-SCREENCLS-* 系を踏襲)

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0039] clsFormatPreviewScreen.Init ENTER"  ' [ADR-0100]
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0040] clsFormatPreviewScreen.Setup ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "LOG-M04-SCREENCLS-SETUP-ENTRY")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "Setup", "EXIT ok", "LOG-M04-SCREENCLS-SETUP-EXIT-OK")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0041] clsFormatPreviewScreen.Setup EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0042] clsFormatPreviewScreen.Setup EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Call modScreenRender.LogScreenError("clsFormatPreviewScreen", "Setup", stepName, Err.Number, Err.Description, "LOG-M04-SCREENCLS-SETUP-ERR")
End Sub

Public Sub Render()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0043] clsFormatPreviewScreen.Render ENTER"  ' [ADR-0100]
    Setup
End Sub

' ================================================================
' 関数名: RenderPreview
' 概要:   M-04 プレビューグリッドの動的再描画。M-03 で編集中の
'         フォーマット (M-03!C3 の FormatID) が変わった／プレビュー
'         画面へ遷移した際に呼ぶ再描画エントリ。M-04 UI スタンザの
'         [GRID_FROM_FORMAT] セクションを読み、動的描画域を一旦
'         クリアしてから modUILoader.ApplyGridFromFormat へ描画を
'         委譲する (v2_ui_stanza_schema.md §3.13 / §4 注、OP-2 確定)。
' 引数:   なし (描画対象シートは m_spec.SheetName から解決)
' 戻り値: なし
' 備考:   初回描画は Setup -> RenderStandardScreen -> ApplyUiToSheet
'         内で [GRID_FROM_FORMAT] が適用されるため、本 Sub は
'         「再描画」専用。再描画前クリアは画面 cls の責務 (Phase 3
'         実装ログ §5(2))。汎用ハンドラ ApplyGridFromFormat は描画
'         のみで旧データ行クリアを行わないため、ここでクリアする。
'         M-09 表示の clsSearchEngine.DisplayKnowledge と同型の委譲役。
' ================================================================
Public Sub RenderPreview()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0044] clsFormatPreviewScreen.RenderPreview ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0049] clsFormatPreviewScreen.RenderPreview STEP-1 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
    Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "RenderPreview", "ENTER", "LOG-M04-SCREENCLS-PREVIEW-ENTRY")

    ' 前提: Init 済 (描画対象シートの解決に m_spec が必要)
    If m_spec Is Nothing Then
        If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0050] clsFormatPreviewScreen.RenderPreview STEP-2 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
        Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "RenderPreview", "spec not initialised", "LOG-M04-SCREENCLS-PREVIEW-WARN")
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0045] clsFormatPreviewScreen.RenderPreview EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ' 描画対象シート (M-04) を解決
    stepName = "resolve sheet " & m_spec.SheetName
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(m_spec.SheetName)

    ' M-04 スタンザの [GRID_FROM_FORMAT] セクションを取得
    stepName = "FindGridFromFormatSection"
    Dim gridSec As ClsStanzaSection
    Set gridSec = FindGridFromFormatSection()
    If gridSec Is Nothing Then
        If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0051] clsFormatPreviewScreen.RenderPreview STEP-3 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
        Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "RenderPreview", "M-04 stanza has no GRID_FROM_FORMAT section", "LOG-M04-SCREENCLS-PREVIEW-WARN")
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0046] clsFormatPreviewScreen.RenderPreview EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ' 再描画前クリア (前フォーマットの余分な行を消す。ApplyGridFromFormat
    ' は描画のみで旧データ行クリアを行わないため画面 cls 側でクリアする)
    stepName = "clear preview area"
    On Error Resume Next
    ws.Range(PREVIEW_CLEAR_RANGE).ClearContents
    On Error GoTo ErrHandler

    ' [GRID_FROM_FORMAT] 描画を modUILoader へ委譲 (試打用・データ層へ保存しない)
    stepName = "ApplyGridFromFormat"
    modUILoader.ApplyGridFromFormat ws, gridSec

    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0052] clsFormatPreviewScreen.RenderPreview STEP-4 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
    Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "RenderPreview", "EXIT ok", "LOG-M04-SCREENCLS-PREVIEW-EXIT-OK")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0047] clsFormatPreviewScreen.RenderPreview EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0053] clsFormatPreviewScreen.RenderPreview STEP-5 pre modScreenRender.LogScreenError"  ' [ADR-0100]
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0048] clsFormatPreviewScreen.RenderPreview EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Call modScreenRender.LogScreenError("clsFormatPreviewScreen", "RenderPreview", stepName, Err.Number, Err.Description, "LOG-M04-SCREENCLS-PREVIEW-ERR")
End Sub

Public Function ValidateInput() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0054] clsFormatPreviewScreen.ValidateInput ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' Screen-level input precondition check (M-04 format preview).
    ' The preview screen is display-only and defines no Required input
    ' fields, so an initialised spec validates True; an uninitialised
    ' screen (m_spec Is Nothing) fails the precondition.
    If m_spec Is Nothing Then
        If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0059] clsFormatPreviewScreen.ValidateInput STEP-1 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
        Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "ValidateInput", "spec not initialised", "VALIDATE-WARN-WW-204")
        ValidateInput = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0055] clsFormatPreviewScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim fld As clsFieldSpec
    For Each fld In m_spec.Fields
        If fld.Required And Len(fld.InputAddr) > 0 Then
            If Len(Trim(GetInputValue(fld.InputAddr))) = 0 Then
                If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0060] clsFormatPreviewScreen.ValidateInput STEP-2 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
                Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "ValidateInput", "required field empty: " & fld.Label, "VALIDATE-WARN-WW-204")
                ValidateInput = False
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0056] clsFormatPreviewScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        End If
    Next fld

    ValidateInput = True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0057] clsFormatPreviewScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0061] clsFormatPreviewScreen.ValidateInput STEP-3 pre modScreenRender.LogScreenError"  ' [ADR-0100]
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0058] clsFormatPreviewScreen.ValidateInput EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Call modScreenRender.LogScreenError("clsFormatPreviewScreen", "ValidateInput", "validate", Err.Number, Err.Description, "VALIDATE-WARN-WW-204")
    ValidateInput = False
End Function

' ================================================================
' 関数名: FindGridFromFormatSection
' 概要:   M-04 の UI スタンザを読み [GRID_FROM_FORMAT] セクションを
'         返す。xlsmName は管理.xlsm = "Kanri" (modEntryFormat の
'         M-03 読込と同じ命名)。screenId は m_spec.ScreenId (= "M-04")。
' 引数:   なし
' 戻り値: ClsStanzaSection - 見つからなければ Nothing
' ================================================================
Private Function FindGridFromFormatSection() As ClsStanzaSection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0062] clsFormatPreviewScreen.FindGridFromFormatSection ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim sections As Collection
    Set sections = modUILoader.LoadUiDefinition(UiXlsmName(), m_spec.ScreenId)

    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To sections.Count
        Set sec = sections.Item(i)
        If sec.SectionName = SECTION_GRID_FROM_FORMAT Then
            Set FindGridFromFormatSection = sec
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0063] clsFormatPreviewScreen.FindGridFromFormatSection EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next i

    Set FindGridFromFormatSection = Nothing
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0064] clsFormatPreviewScreen.FindGridFromFormatSection EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0065] clsFormatPreviewScreen.FindGridFromFormatSection EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set FindGridFromFormatSection = Nothing
End Function

Private Function GetInputValue(ByVal cellAddr As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0066] clsFormatPreviewScreen.GetInputValue ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(m_spec.SheetName)
    If ws Is Nothing Then
        GetInputValue = ""
    Else
        GetInputValue = CStr(ws.Range(cellAddr).Value)
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0067] clsFormatPreviewScreen.GetInputValue EXIT-OK"  ' [ADR-0100]
End Function
```
