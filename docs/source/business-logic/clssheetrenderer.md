---
title: clsSheetRenderer.cls
---

# clsSheetRenderer.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | IScreenRenderer 実装。シート上にボタン / ラベル / 帯を物理配置する描画クラス |
| 行数 | 381 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsSheetRenderer` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IScreenRenderer

' ================================================================
' クラス: clsSheetRenderer（画面層 — 実装1）
' 概要:   IScreenRenderer のシート埋込型実装。
'         ws.Range / ws.Shapes を使ってセル背景・ラベル・フォーム
'         コントロールボタンを物理配置する。
' 依存先: modCommon (カラー定数), clsButtonSpec, clsFieldSpec, clsLogger
' 備考:   将来 UserForm 切替する場合は clsUserFormRenderer に差し替え可。
'         画面クラス側は IScreenRenderer 経由で呼ぶため変更不要。
'         v21 (E2E rerun) で RenderButton に AddFormControl 失敗ログを注入。
' ================================================================

Private Const MOD_NAME As String = "clsSheetRenderer"

Private Const XL_BUTTON_CONTROL As Long = 0  ' xlButtonControl と同値（headless 互換）
Private Const BTN_MIN_W As Double = 100#
Private Const BTN_MIN_H As Double = 26#
Private Const TITLE_ROW_HEIGHT As Double = 28#
Private Const SECTION_ROW_HEIGHT As Double = 20#

Private m_ws As Worksheet

' ================================================================
' 関数名: IScreenRenderer_BindSheet
' 概要:   描画対象シートを束縛する
' 引数:   sheetName - 対象シート名
' ================================================================
Private Sub IScreenRenderer_BindSheet(ByVal sheetName As String)
    Set m_ws = ThisWorkbook.Worksheets(sheetName)
End Sub

' ================================================================
' 関数名: IScreenRenderer_ClearScreen
' 概要:   シート内容（セル値・色・シェイプ）を全クリア。
'         再描画前のリセット用。
' ================================================================
Private Sub IScreenRenderer_ClearScreen()
    If m_ws Is Nothing Then Exit Sub
    On Error Resume Next
    m_ws.Cells.Clear
    Dim shp As Shape
    For Each shp In m_ws.Shapes
        shp.Delete
    Next shp
    Err.Clear
    On Error GoTo 0
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderTitle
' 概要:   画面タイトル帯を行 1 に描画する
' 引数:   screenId - "M-01" 等
'         title    - タイトル文字列
'         colorHex - 背景色 HEX（例 "#1F3864"）
' ================================================================
Private Sub IScreenRenderer_RenderTitle(ByVal screenId As String, _
                                         ByVal title As String, _
                                         ByVal colorHex As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range("A1")
    cell.Value = title
    cell.Font.Bold = True
    cell.Font.Color = RGB(255, 255, 255)
    cell.Font.Size = 14
    cell.Interior.Color = HexToRgb(colorHex)
    m_ws.Rows(1).RowHeight = TITLE_ROW_HEIGHT
    ' タイトル列 A:L を背景色塗り
    m_ws.Range("A1:L1").Interior.Color = HexToRgb(colorHex)
    m_ws.Range("A1:L1").Font.Color = RGB(255, 255, 255)
    m_ws.Range("A1:L1").Font.Bold = True
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderSection
' 概要:   セクション帯（■ ラベル）を描画する
' 引数:   sectionAddr   - 開始セル位置 ("A3" 等)
'         sectionLabel  - "■ モード/予定番号" 等
'         colorHex      - 背景色 HEX（例 "#4472C4"）
' ================================================================
Private Sub IScreenRenderer_RenderSection(ByVal sectionAddr As String, _
                                           ByVal sectionLabel As String, _
                                           ByVal colorHex As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range(sectionAddr)
    cell.Value = sectionLabel
    cell.Font.Bold = True
    cell.Font.Color = RGB(255, 255, 255)
    ' セクション帯は A〜L 列を塗る
    Dim rowNum As Long
    rowNum = cell.Row
    Dim rowRange As Range
    Set rowRange = m_ws.Range(m_ws.Cells(rowNum, 1), m_ws.Cells(rowNum, 12))
    rowRange.Interior.Color = HexToRgb(colorHex)
    rowRange.Font.Color = RGB(255, 255, 255)
    rowRange.Font.Bold = True
    m_ws.Rows(rowNum).RowHeight = SECTION_ROW_HEIGHT
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderButton
' 概要:   フォームコントロールボタンを配置する
' 引数:   btnSpec - clsButtonSpec インスタンス（CellAddr/Caption/ColorHex/BtnName 必須、
'                   HintAddr/HintText 任意）
' 備考:   既存同名ボタンは削除→再配置で idempotent。色付けはシート背景セルで擬似実現
'         （フォームコントロール本体は色付け不可のため、ボタン直下のセルを塗ることで
'          グループ識別性を確保）。
' ================================================================
Private Sub IScreenRenderer_RenderButton(ByVal btnSpec As Object)
    Dim stepName As String : stepName = "begin"

    If m_ws Is Nothing Then
        Call LogWarnSafe("RenderButton", "m_ws=Nothing (BindSheet not called?)")
        Exit Sub
    End If

    stepName = "cast spec"
    Dim spec As clsButtonSpec
    Set spec = btnSpec

    Call LogTraceSafe("RenderButton", _
                       "ENTER ws=" & m_ws.Name & " btnName=" & spec.BtnName & " addr=" & spec.CellAddr)

    ' 既存同名ボタン削除
    stepName = "DeleteShapeByName"
    Call DeleteShapeByName(spec.BtnName)

    stepName = "resolve Range " & spec.CellAddr
    Dim rng As Range
    Set rng = m_ws.Range(spec.CellAddr)

    Dim leftPt As Double
    Dim topPt As Double
    Dim widthPt As Double
    Dim heightPt As Double
    leftPt = rng.Left
    topPt = rng.Top
    widthPt = rng.Width
    heightPt = rng.Height
    If widthPt < BTN_MIN_W Then widthPt = BTN_MIN_W
    If heightPt < BTN_MIN_H Then heightPt = BTN_MIN_H

    ' ボタン直下セルに色を塗る（グループ識別性 — フォームコントロール本体は色付け不可）
    stepName = "tint cell"
    On Error Resume Next
    rng.Interior.Color = HexToRgb(spec.ColorHex)
    rng.Font.Color = RGB(255, 255, 255)
    rng.Font.Bold = True
    If Err.Number <> 0 Then
        Call LogWarnSafe("RenderButton", "cell tint failed btnName=" & spec.BtnName & _
                          " errNum=" & Err.Number & " desc=" & Err.Description)
        Err.Clear
    End If
    On Error GoTo 0

    ' フォームコントロールボタン配置（Object late binding で headless/Excel 互換確保）
    stepName = "AddFormControl"
    Dim shp As Shape
    Dim shapesObj As Object
    Set shapesObj = m_ws.Shapes
    Dim afcErrNum As Long : afcErrNum = 0
    Dim afcErrDesc As String : afcErrDesc = ""
    On Error Resume Next
    Set shp = shapesObj.AddFormControl(XL_BUTTON_CONTROL, leftPt, topPt, widthPt, heightPt)
    afcErrNum = Err.Number
    afcErrDesc = Err.Description
    Err.Clear
    On Error GoTo 0

    If shp Is Nothing Then
        Call LogErrorWithErrSafe("RenderButton", _
                                  "AddFormControl returned Nothing btnName=" & spec.BtnName & _
                                  " ws=" & m_ws.Name & " addr=" & spec.CellAddr, _
                                  afcErrNum, afcErrDesc)
        Exit Sub
    End If

    stepName = "shp.Name"
    shp.Name = spec.BtnName

    stepName = "SetButtonCaptionAndAction"
    Call SetButtonCaptionAndAction(shp, spec.Caption, spec.BtnName)

    ' ヒントテキスト（ボタン下の説明セル）
    If Len(spec.HintAddr) > 0 And Len(spec.HintText) > 0 Then
        stepName = "hint " & spec.HintAddr
        m_ws.Range(spec.HintAddr).Value = spec.HintText
        m_ws.Range(spec.HintAddr).WrapText = True
        m_ws.Range(spec.HintAddr).Font.Size = 9
    End If

    Call LogTraceSafe("RenderButton", "EXIT ok btnName=" & spec.BtnName)
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderLabel
' 概要:   ラベルセルを描画（背景色任意）
' ================================================================
Private Sub IScreenRenderer_RenderLabel(ByVal cellAddr As String, _
                                         ByVal labelText As String, _
                                         ByVal colorHex As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range(cellAddr)
    cell.Value = labelText
    If Len(colorHex) > 0 Then
        cell.Interior.Color = HexToRgb(colorHex)
    End If
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderInputField
' 概要:   入力フィールドの「ラベル + 必須マーク + 型表示 + 入力欄ハイライト」を描画。
'         データが空でもフィールドが見えるようにする「空状態 UI」の実装。
' 引数:   cellAddr  - 開始セル（未使用、fieldSpec 内のアドレスを使う）
'         fieldSpec - clsFieldSpec インスタンス
' ================================================================
Private Sub IScreenRenderer_RenderInputField(ByVal cellAddr As String, _
                                              ByVal fieldSpec As Object)
    If m_ws Is Nothing Then Exit Sub
    Dim spec As clsFieldSpec
    Set spec = fieldSpec

    ' 順序番号
    If Len(spec.OrderAddr) > 0 Then
        m_ws.Range(spec.OrderAddr).Value = spec.FieldOrder
        m_ws.Range(spec.OrderAddr).HorizontalAlignment = xlCenter
    End If

    ' 必須マーク
    If spec.Required And Len(spec.ReqMarkAddr) > 0 Then
        Call IScreenRenderer_RenderRequiredMark(spec.ReqMarkAddr)
    End If

    ' ラベル
    If Len(spec.LabelAddr) > 0 Then
        m_ws.Range(spec.LabelAddr).Value = spec.Label
        m_ws.Range(spec.LabelAddr).HorizontalAlignment = xlRight
    End If

    ' 型表示（イタリック、薄黄）
    If Len(spec.TypeAddr) > 0 Then
        m_ws.Range(spec.TypeAddr).Value = spec.TypeText
        m_ws.Range(spec.TypeAddr).Font.Italic = True
        m_ws.Range(spec.TypeAddr).Interior.Color = HexToRgb(COLOR_HINT_YELLOW)
        m_ws.Range(spec.TypeAddr).Font.Size = 9
    End If

    ' 入力欄ヒント（説明文）
    If Len(spec.InputAddr) > 0 And Len(spec.HintText) > 0 Then
        m_ws.Range(spec.InputAddr).Value = spec.HintText
        m_ws.Range(spec.InputAddr).Font.Color = RGB(128, 128, 128)
        m_ws.Range(spec.InputAddr).Font.Italic = True
    End If
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderRequiredMark
' 概要:   必須マーク（赤背景に "*"）を描画
' ================================================================
Private Sub IScreenRenderer_RenderRequiredMark(ByVal cellAddr As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range(cellAddr)
    cell.Value = "*"
    cell.Interior.Color = HexToRgb(COLOR_REQUIRED_RED)
    cell.Font.Color = RGB(255, 255, 255)
    cell.Font.Bold = True
    cell.HorizontalAlignment = xlCenter
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderHint
' 概要:   ヒント/凡例テキストを描画
' ================================================================
Private Sub IScreenRenderer_RenderHint(ByVal cellAddr As String, ByVal hintText As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range(cellAddr)
    cell.Value = hintText
    cell.Interior.Color = HexToRgb(COLOR_HINT_BAR)
    cell.Font.Italic = True
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderHeaderRow
' 概要:   一覧テーブルのヘッダ行を描画（M-02/M-07/M-08/M-10/M-12/M-14 共通）
' 引数:   startAddr     - 開始セル（"B10" 等）
'         headerLabels  - ヘッダ文字列の配列
'         colorHex      - 背景色（"#5B9BD5" 等）
' ================================================================
Private Sub IScreenRenderer_RenderHeaderRow(ByVal startAddr As String, _
                                             ByVal headerLabels As Variant, _
                                             ByVal colorHex As String)
    If m_ws Is Nothing Then Exit Sub
    If IsEmpty(headerLabels) Then Exit Sub

    Dim startCell As Range
    Set startCell = m_ws.Range(startAddr)
    Dim startRow As Long
    Dim startCol As Long
    startRow = startCell.Row
    startCol = startCell.Column

    Dim i As Long
    For i = LBound(headerLabels) To UBound(headerLabels)
        Dim cell As Range
        Set cell = m_ws.Cells(startRow, startCol + i)
        cell.Value = CStr(headerLabels(i))
        cell.Font.Bold = True
        cell.Font.Color = RGB(255, 255, 255)
        cell.Interior.Color = HexToRgb(colorHex)
        cell.HorizontalAlignment = xlCenter
    Next i
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderEmptyState
' 概要:   「データなし」状態のメッセージを描画
' ================================================================
Private Sub IScreenRenderer_RenderEmptyState(ByVal cellAddr As String, _
                                              ByVal message As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range(cellAddr)
    cell.Value = message
    cell.Font.Italic = True
    cell.Font.Color = RGB(128, 128, 128)
End Sub

' --- 内部ユーティリティ ---

' ================================================================
' 関数名: HexToRgb
' 概要:   HEX 文字列 ("#1F3864" or "1F3864") を Long の RGB 値に変換
' 引数:   hex - HEX 文字列
' 戻り値: Long - VBA RGB 値
' ================================================================
Private Function HexToRgb(ByVal hex As String) As Long
    Dim s As String
    s = hex
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)
    If Len(s) <> 6 Then
        HexToRgb = RGB(255, 255, 255)
        Exit Function
    End If
    Dim r As Long
    Dim g As Long
    Dim b As Long
    r = CLng("&H" & Mid$(s, 1, 2))
    g = CLng("&H" & Mid$(s, 3, 2))
    b = CLng("&H" & Mid$(s, 5, 2))
    HexToRgb = RGB(r, g, b)
End Function

' ================================================================
' 関数名: DeleteShapeByName
' 概要:   指定名のシェイプが存在すれば削除（idempotent）
' ================================================================
Private Sub DeleteShapeByName(ByVal shapeName As String)
    If m_ws Is Nothing Then Exit Sub
    Dim shp As Shape
    On Error Resume Next
    Set shp = m_ws.Shapes(shapeName)
    If Not shp Is Nothing Then shp.Delete
    Err.Clear
    On Error GoTo 0
En
```
