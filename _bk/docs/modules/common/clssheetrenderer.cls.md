---
title: clsSheetRenderer.cls
description: clsSheetRenderer.cls のソースコード（コピペ用）
---

# clsSheetRenderer.cls

**配置先**: `共通モジュール (3 ブック全て)` 用の VBA モジュール  
**種類**: クラス モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\`
- ファイル名: `clsSheetRenderer.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

---

## ソースコード

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "clsSheetRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽN・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽX: clsSheetRenderer・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽiv2.1・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽAPhase 3 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽj・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽAthin shim ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽj
' ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽT・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽv:   IScreenRenderer ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ Sheet ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ^・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽAmodUILoader.ApplyUiToSheet ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽﾉ集・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ
' Version: v2.1・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽi2026-05-16 EOD・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽAQ20/Q5 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽf・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽj
' ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽﾖ連:   ADR-0053 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ6・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽi・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽO・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ UI ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽX・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ^・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽU・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽj, IScreenRenderer.cls
' v2.1 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽv・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽX・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽV:
'   - v1 521 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽs ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ v2.1 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ 100 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽs thin shim
'   - ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽﾂ包ｿｽ Render* ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ\・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽb・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽh・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽp・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ~・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽAApplyFromStanza 1 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽs・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽW・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ
'   - modUILoader.ApplyUiToSheet・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽiQ20 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽﾖ撰ｿｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽj・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽﾄ出
' ================================================================
Implements IScreenRenderer
Option Explicit

Private m_targetSheet As Worksheet
Private m_xlsmName As String

' ----------------------------------------------------------------
' IScreenRenderer ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽi8 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ\・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽb・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽh・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽAv2.1・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽj
' ----------------------------------------------------------------

Private Sub IScreenRenderer_BindSheet(ByVal sheetName As String)
    ' Phase R-2-Fix (2026-05-28): sheets are renamed to JP display names by
    ' RenameSheetsToDisplayNames in step 4 of RunFullSetup, but ApplyUiStanzas
    ' still uses M-NN logical ids. Try logical name first, then fall back to
    ' the display name resolved from the ui_seed [SHEET] section.
    ' CRITICAL: clear m_targetSheet first so a failed lookup doesn't retain
    ' the previous iteration's sheet (silent bug discovered 2026-05-28 R-2-Fix).
    Set m_targetSheet = Nothing
    On Error Resume Next
    Set m_targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If Not m_targetSheet Is Nothing Then Exit Sub
    ' Fallback: read seed and look up SheetName from [SHEET] section
    On Error Resume Next
    Dim displayName As String
    displayName = ResolveDisplayNameFromSeed(sheetName)
    If Len(displayName) > 0 Then
        Set m_targetSheet = ThisWorkbook.Worksheets(displayName)
    End If
    On Error GoTo 0
    If m_targetSheet Is Nothing Then
        Err.Raise vbObjectError + 100, "clsSheetRenderer", _
                  "BindSheet failed: tried '" & sheetName & "' and display name fallback"
    End If
End Sub

' Phase R-2-Fix / R-3: look up display name from ui_seed [SHEET].SheetName.
' Robust version: walks every subdir of ui_dir to find <screenId>.txt.
Private Function ResolveDisplayNameFromSeed(ByVal screenId As String) As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim uiDir As String
    uiDir = modConfigHolder.GetUiDir()
    If Not fso.FolderExists(uiDir) Then Exit Function
    Dim root As Object
    Set root = fso.GetFolder(uiDir)
    Dim sub_ As Object
    For Each sub_ In root.SubFolders
        Dim p As String
        p = sub_.Path & "\" & screenId & ".txt"
        If fso.FileExists(p) Then
            Dim secs As Collection
            Set secs = modStanzaIO.ParseStanzaFile(p)
            Dim s As ClsStanzaSection
            Dim j As Long
            For j = 1 To secs.Count
                Set s = secs.Item(j)
                If s.SectionName = "SHEET" Then
                    Dim nm As String
                    nm = Trim$(s.GetValue("SheetName"))
                    ' Verify the workbook has a sheet with this display name.
                    Dim testWs As Worksheet
                    Set testWs = Nothing
                    On Error Resume Next
                    Set testWs = ThisWorkbook.Worksheets(nm)
                    On Error GoTo 0
                    If Not testWs Is Nothing Then
                        ResolveDisplayNameFromSeed = nm
                        Exit Function
                    End If
                End If
            Next j
        End If
    Next sub_
End Function

Private Sub IScreenRenderer_ClearScreen()
    If m_targetSheet Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    ' Phase R-2-Fix (2026-05-28): unprotect before clear so a previously-
    ' protected sheet doesn't silently retain content across re-installs.
    On Error Resume Next
    m_targetSheet.Unprotect Password:=""
    On Error GoTo ErrHandler
    m_targetSheet.Cells.Clear
    ' Shape ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽn・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽi・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ{・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ^・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽj・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE
    Dim shp As Shape
    For Each shp In m_targetSheet.Shapes
        shp.Delete
    Next shp
    Exit Sub
ErrHandler:
    Debug.Print "[clsSheetRenderer.ClearScreen ERROR] " & Err.Description
End Sub

' v2.1 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽj・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽF1 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽs・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ UI ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽS・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ\・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽz・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽiQ20 modUILoader.ApplyUiToSheet ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽﾄ出・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽj
Private Sub IScreenRenderer_ApplyFromStanza(ByVal xlsmName As String, ByVal screenId As String)
    On Error GoTo ErrHandler
    If m_targetSheet Is Nothing Then
        Err.Raise vbObjectError + 101, "clsSheetRenderer", _
                  "ApplyFromStanza: BindSheet ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽs"
    End If
    m_xlsmName = xlsmName
    Call modUILoader.ApplyUiToSheet(xlsmName, screenId, m_targetSheet)
    Exit Sub
ErrHandler:
    Err.Raise vbObjectError + 102, "clsSheetRenderer", _
              "ApplyFromStanza ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽs: " & xlsmName & "/" & screenId & " - " & Err.Description
End Sub

Private Sub IScreenRenderer_ShowSheet()
    If m_targetSheet Is Nothing Then Exit Sub
    m_targetSheet.Visible = xlSheetVisible
End Sub

Private Sub IScreenRenderer_HideSheet()
    If m_targetSheet Is Nothing Then Exit Sub
    m_targetSheet.Visible = xlSheetHidden  ' Q2: VeryHidden ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽﾅはなゑｿｽ
End Sub

Private Sub IScreenRenderer_ActivateSheet()
    If m_targetSheet Is Nothing Then Exit Sub
    m_targetSheet.Activate
End Sub

Private Sub IScreenRenderer_ProtectSheet(ByVal level As String)
    If m_targetSheet Is Nothing Then Exit Sub
    On Error Resume Next
    Select Case LCase(level)
        Case "light"
            ' light ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽﾛ鯉ｿｽF・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ[・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽU・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽZ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽI・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽﾍ可、・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽﾍ不・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ
            m_targetSheet.Protect Password:="", AllowFormattingCells:=True, _
                                   AllowFormattingColumns:=True, AllowFormattingRows:=True
        Case "strict"
            ' strict ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽﾛ鯉ｿｽF・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽS lock
            m_targetSheet.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True
        Case Else
            m_targetSheet.Protect Password:=""
    End Select
    ' Q7 ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽK・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ Y・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽFOn Error Resume Next ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽE・ｽ Err check
    If Err.Number <> 0 Then
        Debug.Print "[clsSheetRenderer.ProtectSheet ERROR] " & Err.Description
        Err.Clear
    End If
End Sub

Private Sub IScreenRenderer_UnprotectSheet()
    If m_targetSheet Is Nothing Then Exit Sub
    On Error Resume Next
    m_targetSheet.Unprotect Password:=""
    If Err.Number <> 0 Then
        Debug.Print "[clsSheetRenderer.UnprotectSheet ERROR] " & Err.Description
        Err.Clear
    End If
End Sub
```
