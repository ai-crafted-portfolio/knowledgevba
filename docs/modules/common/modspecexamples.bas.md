---
title: modSpecExamples.bas
description: modSpecExamples.bas 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# modSpecExamples.bas

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`modSpecExamples.bas`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
```vb
Attribute VB_Name = "modSpecExamples"
Option Explicit

' ================================================================
' モジュール: modSpecExamples (エントリポイント層)
' 概要:       clsFormSpec の使い方デモ。
'             検索結果プレビュー用 spec を組み立てる。
'             UI から呼び出さなくても Application.Run "Macro_Show..."
'             で手動起動できる。
' 依存先:     clsFormSpec, modFormBuilder, clsSearchEngine
' 備考:       2026-05-20 ShowError/ShowWarning 浮き解消 (helper 廃止に伴い MsgBox 直呼出化)
' ================================================================

' ================================================================
' 関数名: NewSearchResultSpec
' 概要:   検索結果プレビュー用フォーム spec を組み立てる
' 引数:   knwNo   - ナレッジ番号 (フォームタイトルに表示)
'         imgPath - 画像絶対パス ("" なら fallback テキスト)
'         snippet - 表示するスニペット (タイトル + 本文抜粋)
' 戻り値: clsFormSpec
' ================================================================
Public Function NewSearchResultSpec(ByVal knwNo As String, _
                                      ByVal imgPath As String, _
                                      ByVal snippet As String) As clsFormSpec
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1372] modSpecExamples.NewSearchResultSpec ENTER"  ' [ADR-0100]
    Dim spec As clsFormSpec
    Set spec = New clsFormSpec
    spec.FormTitle = ChrW(&H691C) & ChrW(&H7D22) & ChrW(&H7D50) & ChrW(&H679C) & ChrW(&H30D7) & ChrW(&H30EC) & ChrW(&H30D3) & ChrW(&H30E5) & ChrW(&H30FC) & ": " & knwNo
    spec.Width = 600
    spec.Height = 460

    ' Title Label (上部)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1374] modSpecExamples.NewSearchResultSpec STEP-1 pre spec.AddControl"  ' [ADR-0100]
    Call spec.AddControl("Label", "lblTitle", 10, 10, 580, 20, _
                          ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ": " & knwNo, "")

    ' 画像領域 (左)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1375] modSpecExamples.NewSearchResultSpec STEP-2 pre spec.AddControl"  ' [ADR-0100]
    Call spec.AddControl("Image", "imgPreview", 10, 40, 400, 300, "", "")

    ' スニペット (右)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1376] modSpecExamples.NewSearchResultSpec STEP-3 pre spec.AddControl"  ' [ADR-0100]
    Call spec.AddControl("TextBox", "txtSnippet", 420, 40, 170, 300, _
                          snippet, "")

    ' 閉じるボタン
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1377] modSpecExamples.NewSearchResultSpec STEP-4 pre spec.AddControl"  ' [ADR-0100]
    Call spec.AddControl("Button", "btnClose", 250, 360, 80, 30, _
                          ChrW(&H9589) & ChrW(&H3058) & ChrW(&H308B), "frmCallback_searchResult_close")

    Set NewSearchResultSpec = spec
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1373] modSpecExamples.NewSearchResultSpec EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: Macro_ShowSearchResultPreview
' 概要:   現在 検索シートで選択中の行のナレッジを spec 駆動フォームで表示
' 備考:   Application.Run or Alt+F8 から手動呼出可能。
'         既存の Btn_DetailDisplay (M-09 シート遷移) と並立。
' ================================================================
Public Sub Macro_ShowSearchResultPreview()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1378] modSpecExamples.Macro_ShowSearchResultPreview ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)

    Dim selRow As Long
    selRow = ws.Application.Selection.Row
    If selRow < 15 Then
        ' 2026-05-20 xref clean: ShowWarning helper 廃止に伴い MsgBox 直呼出に変更
        ' 2026-05-26 v2.3 install hang fix: modCommon.IsHeadless() ガード追加。
        ' Macro_ShowSearchResultPreview は手動 Application.Run 用途で install
        ' 経路では呼ばれないはずだが、念のため headless ガードを入れて
        ' install_admin.bat 実行中の予期せぬハングを防ぐ。
        If Not modCommon.IsHeadless() Then
            MsgBox ChrW(&H691C) & ChrW(&H7D22) & ChrW(&H7D50) & ChrW(&H679C) & ChrW(&H306E) & ChrW(&H884C) & ChrW(&H304C) & ChrW(&H9078) & ChrW(&H629E) & ChrW(&H3055) & ChrW(&H308C) & ChrW(&H3066) & ChrW(&H3044) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & vbCrLf & _
                   ChrW(&H7D50) & ChrW(&H679C) & ChrW(&H4E00) & ChrW(&H89A7) & ChrW(&H304B) & ChrW(&H3089) & ChrW(&H884C) & ChrW(&H3092) & ChrW(&H9078) & ChrW(&H3093) & ChrW(&H3067) & ChrW(&H518D) & ChrW(&H5B9F) & ChrW(&H884C) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044), _
                   vbExclamation, ChrW(&H30D7) & ChrW(&H30EC) & ChrW(&H30D3) & ChrW(&H30E5) & ChrW(&H30FC)
        Else
            Debug.Print "[HEADLESS] Macro_ShowSearchResultPreview: " & ChrW(&H884C) & ChrW(&H672A) & ChrW(&H9078) & ChrW(&H629E) & " (MsgBox " & ChrW(&H6291) & ChrW(&H6B62) & ")"
        End If
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1379] modSpecExamples.Macro_ShowSearchResultPreview EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    Dim knwNo As String
    knwNo = CStr(ws.Cells(selRow, 2).Value)
    If knwNo = "" Then Exit Sub

    Dim title As String
    title = CStr(ws.Cells(selRow, 4).Value)
    Dim score As String
    score = CStr(ws.Cells(selRow, 9).Value)

    Dim spec As clsFormSpec
    Set spec = NewSearchResultSpec(knwNo, "", _
                                     "Title: " & title & vbCrLf & _
                                     "Score: " & score & vbCrLf & vbCrLf & _
                                     "(" & ChrW(&H753B) & ChrW(&H50CF) & ChrW(&H3068) & ChrW(&H672C) & ChrW(&H6587) & ChrW(&H306F) & " ChromaDB " & ChrW(&H9023) & ChrW(&H52D5) & ChrW(&H771F) & ChrW(&H7248) & ChrW(&H3067) & ChrW(&H5C55) & ChrW(&H958B) & ")")

    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1382] modSpecExamples.Macro_ShowSearchResultPreview STEP-1 pre BuildAndShow"  ' [ADR-0100]
    Call BuildAndShow(spec, True)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1380] modSpecExamples.Macro_ShowSearchResultPreview EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1381] modSpecExamples.Macro_ShowSearchResultPreview EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' 2026-05-20 xref clean: ShowError helper 廃止に伴い MsgBox 直呼出に変更
    ' 2026-05-26 v2.3 install hang fix: modCommon.IsHeadless() ガード追加。
    If Not modCommon.IsHeadless() Then
        MsgBox Err.Description & vbCrLf & _
               "VBA " & ChrW(&H30D7) & ChrW(&H30ED) & ChrW(&H30B8) & ChrW(&H30A7) & ChrW(&H30AF) & ChrW(&H30C8) & " " & ChrW(&H30AA) & ChrW(&H30D6) & ChrW(&H30B8) & ChrW(&H30A7) & ChrW(&H30AF) & ChrW(&H30C8) & " " & ChrW(&H30E2) & ChrW(&H30C7) & ChrW(&H30EB) & ChrW(&H4FE1) & ChrW(&H983C) & ChrW(&H304C) & " ON " & ChrW(&H304B) & ChrW(&H78BA) & ChrW(&H8A8D), _
               vbCritical, ChrW(&H30D7) & ChrW(&H30EC) & ChrW(&H30D3) & ChrW(&H30E5) & ChrW(&H30FC)
    Else
        Debug.Print "[HEADLESS] Macro_ShowSearchResultPreview ErrHandler: " & Err.Description
    End If
End Sub

' ================================================================
' 関数名: frmCallback_searchResult_close
' 概要:   spec 駆動フォームの「閉じる」ボタン Click コールバック
' 引数:   frm - UserForm インスタンス (Application.Run の第 2 引数で渡る)
' 備考:   Application.Run の引数として UserForm が渡るが、本モックでは
'         単純に Unload するだけ。
' ================================================================
Public Sub frmCallback_searchResult_close(ByVal frm As Object)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1383] modSpecExamples.frmCallback_searchResult_close ENTER"  ' [ADR-0100]
    On Error Resume Next
    Unload frm
End Sub
```