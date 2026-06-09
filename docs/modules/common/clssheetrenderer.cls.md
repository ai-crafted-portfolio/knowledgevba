---
title: clsSheetRenderer.cls
description: clsSheetRenderer.cls のソースコード（コピペ用）
---

# clsSheetRenderer.cls

**配置先**: 共通モジュール（3 ブック共通）
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsSheetRenderer.cls`
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
Attribute VB_Name = "clsSheetRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' クラス: clsSheetRenderer（v2.1、Phase 3 抽出、thin shim 化）
' 概要:   IScreenRenderer の Sheet 型実装、modUILoader.ApplyUiToSheet に集約
' Version: v2.1（2026-05-16 EOD、Q20/Q5 反映）
' 関連:   ADR-0053 第6章（外部 UI スタンザ駆動）, IScreenRenderer.cls
' v2.1 主要更新:
'   - v1 521 行 から v2.1 約 100 行 thin shim
'   - 旧 Render* メソッド廃止、ApplyFromStanza 1 行集約
'   - modUILoader.ApplyUiToSheet（Q20 関数）を呼出
' ================================================================
Implements IScreenRenderer
Option Explicit

Private m_targetSheet As Worksheet
Private m_xlsmName As String

' ----------------------------------------------------------------
' IScreenRenderer 実装メソッド（8 メソッド、v2.1）
' ----------------------------------------------------------------

Private Sub IScreenRenderer_BindSheet(ByVal sheetName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0743] clsSheetRenderer.IScreenRenderer_BindSheet ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0744] clsSheetRenderer.IScreenRenderer_BindSheet EXIT-OK"  ' [ADR-0100]
End Sub

' Phase R-2-Fix / R-3: look up display name from ui_seed [SHEET].SheetName.
' Robust version: walks every subdir of ui_dir to find <screenId>.txt.
Private Function ResolveDisplayNameFromSeed(ByVal screenId As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0745] clsSheetRenderer.ResolveDisplayNameFromSeed ENTER"  ' [ADR-0100]
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
                    nm = Trim(s.GetValue("SheetName"))
                    ' Verify the workbook has a sheet with this display name.
                    Dim testWs As Worksheet
                    Set testWs = Nothing
                    On Error Resume Next
                    Set testWs = ThisWorkbook.Worksheets(nm)
                    On Error GoTo 0
                    If Not testWs Is Nothing Then
                        ResolveDisplayNameFromSeed = nm
                        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0746] clsSheetRenderer.ResolveDisplayNameFromSeed EXIT-OK"  ' [ADR-0100]
                        Exit Function
                    End If
                End If
            Next j
        End If
    Next sub_
End Function

Private Sub IScreenRenderer_ClearScreen()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0747] clsSheetRenderer.IScreenRenderer_ClearScreen ENTER"  ' [ADR-0100]
    If m_targetSheet Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    ' Phase R-2-Fix (2026-05-28): unprotect before clear so a previously-
    ' protected sheet doesn't silently retain content across re-installs.
    On Error Resume Next
    m_targetSheet.Unprotect Password:=""
    On Error GoTo ErrHandler
    m_targetSheet.Cells.Clear
    ' Shape 系（ボタン等）を削除
    Dim shp As Shape
    For Each shp In m_targetSheet.Shapes
        shp.Delete
    Next shp
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0748] clsSheetRenderer.IScreenRenderer_ClearScreen EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0749] clsSheetRenderer.IScreenRenderer_ClearScreen EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[clsSheetRenderer.ClearScreen ERROR] " & Err.Description
End Sub

' v2.1 主旨：1 行で UI 全構築（Q20 modUILoader.ApplyUiToSheet 呼出）
Private Sub IScreenRenderer_ApplyFromStanza(ByVal xlsmName As String, ByVal screenId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0750] clsSheetRenderer.IScreenRenderer_ApplyFromStanza ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_targetSheet Is Nothing Then
        Err.Raise vbObjectError + 101, "clsSheetRenderer", _
                  "ApplyFromStanza: BindSheet " & ChrW(&H304C) & ChrW(&H672A) & ChrW(&H5B9F) & ChrW(&H884C)
    End If
    m_xlsmName = xlsmName
    Call modUILoader.ApplyUiToSheet(xlsmName, screenId, m_targetSheet)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0751] clsSheetRenderer.IScreenRenderer_ApplyFromStanza EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0752] clsSheetRenderer.IScreenRenderer_ApplyFromStanza EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Err.Raise vbObjectError + 102, "clsSheetRenderer", _
              "ApplyFromStanza " & ChrW(&H5931) & ChrW(&H6557) & ": " & xlsmName & "/" & screenId & " - " & Err.Description
End Sub

Private Sub IScreenRenderer_ShowSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0753] clsSheetRenderer.IScreenRenderer_ShowSheet ENTER"  ' [ADR-0100]
    If m_targetSheet Is Nothing Then Exit Sub
    m_targetSheet.Visible = xlSheetVisible
End Sub

Private Sub IScreenRenderer_HideSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0754] clsSheetRenderer.IScreenRenderer_HideSheet ENTER"  ' [ADR-0100]
    If m_targetSheet Is Nothing Then Exit Sub
    m_targetSheet.Visible = xlSheetHidden  ' Q2: VeryHidden ではない
End Sub

Private Sub IScreenRenderer_ActivateSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0755] clsSheetRenderer.IScreenRenderer_ActivateSheet ENTER"  ' [ADR-0100]
    If m_targetSheet Is Nothing Then Exit Sub
    m_targetSheet.Activate
End Sub

Private Sub IScreenRenderer_ProtectSheet(ByVal level As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0756] clsSheetRenderer.IScreenRenderer_ProtectSheet ENTER"  ' [ADR-0100]
    If m_targetSheet Is Nothing Then Exit Sub
    On Error Resume Next
    Select Case LCase(level)
        Case "light"
            ' light 保護：ユーザ操作はセル選択や書式変更まで許可
            m_targetSheet.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True, _
                                   AllowFormattingColumns:=True, AllowFormattingRows:=True
        Case "strict"
            ' strict 保護：全 lock
            m_targetSheet.Protect Password:="", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True
        Case Else
            m_targetSheet.Protect Password:="", UserInterfaceOnly:=True
    End Select
    ' Q7 規約 Y：On Error Resume Next 配下でも Err check は実施
    If Err.Number <> 0 Then
        Debug.Print "[clsSheetRenderer.ProtectSheet ERROR] " & Err.Description
        Err.Clear
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0757] clsSheetRenderer.IScreenRenderer_ProtectSheet EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub IScreenRenderer_UnprotectSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0758] clsSheetRenderer.IScreenRenderer_UnprotectSheet ENTER"  ' [ADR-0100]
    If m_targetSheet Is Nothing Then Exit Sub
    On Error Resume Next
    m_targetSheet.Unprotect Password:=""
    If Err.Number <> 0 Then
        Debug.Print "[clsSheetRenderer.UnprotectSheet ERROR] " & Err.Description
        Err.Clear
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0759] clsSheetRenderer.IScreenRenderer_UnprotectSheet EXIT-OK"  ' [ADR-0100]
End Sub
```
