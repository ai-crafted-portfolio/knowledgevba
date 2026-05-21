---
title: modImageRender.bas
---

# modImageRender.bas

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | ナレッジに添付された画像のサムネ・詳細表示 |
| 行数 | 233 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modImageRender.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modImageRender"
Option Explicit

' v6: Shapes.AddPicture / LockAspectRatio などは MsoTriState Enum を期待するため、
'      named arg + Long Const を渡すと real Excel が compile-time に拒否する (ADR 0020)。
'      → 位置指定 + 数値リテラル直書きで MsoTriState への暗黙変換に頼る。
'      AddPicture 第 2 引数: 0=msoFalse / -1=msoTrue / 1=msoCTrue (Office 標準値)
'      第 3 引数: 同 MsoTriState

' ================================================================
' モジュール: modImageRender (ユーティリティ層)
' 概要:       ナレッジ画像 (PNG) のサムネ表示と詳細画像ペイン描画。
'             Shapes.AddPicture を使うため LoadPicture と異なり
'             Worksheet 上でも Form 上でも Shape 残留問題を起こさない。
'             サムネは検索結果一覧の H 列、詳細画像は M-09 表示シートの
'             J4:N20 領域に配置。
' 依存先:     なし (純粋ユーティリティ)
' 備考:       VBA 子プロセス禁止 (Shell/Run/Exec/CreateObject Exec 系全部
'             禁止) のため、画像変換やフェッチは行わない。
'             ファイル不在時は薄字フォールバック表示のみ。
' ================================================================

' --- サムネサイズ ---
Public Const KB_THUMB_WIDTH As Double = 60#
Public Const KB_THUMB_HEIGHT As Double = 40#

' --- 詳細画像最大サイズ ---
Private Const KB_DETAIL_MAX_WIDTH As Double = 400#
Private Const KB_DETAIL_MAX_HEIGHT As Double = 300#

' --- Shape 名 prefix (削除/再配置用) ---
Public Const PREFIX_THUMB As String = "kbThumb_"
Public Const PREFIX_DETAIL As String = "kbDetailImg_"

' --- セル padding ---
Private Const PAD_PIXELS As Double = 2#

' ================================================================
' 関数名: RenderThumb
' 概要:   検索結果一覧のサムネセル (col=8, H 列) に画像 Shape を貼付。
'         画像が存在しない場合は何も描画せず黙ってリターン (薄字テキストは
'         呼び出し側でセルに書く運用)。
' 引数:   ws         - 配置対象シート
'         row        - 配置先の行番号
'         col        - 配置先の列番号
'         imageFull  - 画像の絶対パス
'         knwNo      - ナレッジ番号 (Shape 名識別用)
' 戻り値: なし
' 備考:   Shape 名は PREFIX_THUMB & "<row>_" & knwNo 形式。
'         再描画時は ClearShapesByPrefix で全削除→再生成 (idempotent)。
' ================================================================
Public Sub RenderThumb(ByVal ws As Worksheet, ByVal targetRow As Long, _
                         ByVal targetCol As Long, ByVal imageFull As String, _
                         ByVal knwNo As String)
    On Error GoTo ErrHandler
    If imageFull = "" Then Exit Sub
    If Dir(imageFull) = "" Then Exit Sub

    Dim cell As Range
    Set cell = ws.Cells(targetRow, targetCol)

    Dim shp As Shape
    ' v8: Object 型 late binding で AddPicture を呼ぶ (ADR 0023)
    '     typed Shapes 経由だと MsoTriState 引数の数値リテラル暗黙変換が
    '     real Excel で reject される。Object 経由なら compile-time 型チェック skip。
    '     第 2 引数 LinkToFile = 0 (msoFalse), 第 3 引数 SaveWithDocument = -1 (msoTrue)
    Dim shapesObj As Object
    Set shapesObj = ws.Shapes
    Set shp = shapesObj.AddPicture(imageFull, 0, -1, _
        cell.Left + PAD_PIXELS, cell.Top + PAD_PIXELS, _
        KB_THUMB_WIDTH, KB_THUMB_HEIGHT)
    shp.Name = PREFIX_THUMB & CStr(targetRow) & "_" & knwNo

    ' 行高をサムネに合わせる (狭ければ拡大)
    If ws.Rows(targetRow).RowHeight < (KB_THUMB_HEIGHT + PAD_PIXELS * 2#) Then
        ws.Rows(targetRow).RowHeight = KB_THUMB_HEIGHT + PAD_PIXELS * 2#
    End If
    Exit Sub
ErrHandler:
    ' 画像描画失敗は致命ではない (検索結果テキストは別途書かれている)
End Sub

' ================================================================
' 関数名: RenderDetailImage
' 概要:   ナレッジ表示シートの J4:N20 領域に詳細画像を貼付。
'         領域からはみ出さないようアスペクト比を保ってリサイズ。
'         画像が無い場合は薄字で「[画像未配置: ...]」セル表示。
' 引数:   ws         - 配置対象シート
'         topRow     - ペイン左上行
'         leftCol    - ペイン左上列
'         botRow     - ペイン右下行
'         rightCol   - ペイン右下列
'         imageFull  - 画像の絶対パス
'         knwNo      - ナレッジ番号 (Shape 名識別用)
' ================================================================
Public Sub RenderDetailImage(ByVal ws As Worksheet, _
                               ByVal topRow As Long, ByVal leftCol As Long, _
                               ByVal botRow As Long, ByVal rightCol As Long, _
                               ByVal imageFull As String, _
                               ByVal knwNo As String)
    On Error GoTo ErrHandler

    Dim r1 As Range, r2 As Range
    Set r1 = ws.Cells(topRow, leftCol)
    Set r2 = ws.Cells(botRow, rightCol)
    Dim maxW As Double, maxH As Double
    maxW = (r2.Left + r2.Width) - r1.Left - PAD_PIXELS * 2#
    maxH = (r2.Top + r2.Height) - r1.Top - PAD_PIXELS * 2#
    If maxW < KB_DETAIL_MAX_WIDTH Then maxW = KB_DETAIL_MAX_WIDTH
    If maxH < KB_DETAIL_MAX_HEIGHT Then maxH = KB_DETAIL_MAX_HEIGHT

    If imageFull = "" Or Dir(imageFull) = "" Then
        ' フォールバックテキスト
        Call WriteFallbackText(ws, topRow, leftCol, botRow, rightCol, imageFull)
        Exit Sub
    End If

    Dim shp As Shape
    ' v8: Object 型 late binding で AddPicture を呼ぶ (ADR 0023)
    '     第 2 引数 LinkToFile = 0 (msoFalse), 第 3 引数 SaveWithDocument = -1 (msoTrue)
    '     Width=-1 / Height=-1 はオリジナルサイズで読込 (AddPicture 仕様)
    Dim shapesObj2 As Object
    Set shapesObj2 = ws.Shapes
    Set shp = shapesObj2.AddPicture(imageFull, 0, -1, _
        r1.Left + PAD_PIXELS, r1.Top + PAD_PIXELS, _
        -1, -1)
    shp.Name = PREFIX_DETAIL & knwNo
    ' v8: Object 型 late binding で LockAspectRatio を設定 (ADR 0023)
    '     v7 の Variant 中継でも real Excel が typed Shape 経由の MsoTriState 代入を
    '     reject する事例があるため、Shape を Object 経由にラップして compile-time
    '     型チェック自体を skip する。
    Dim shpObj As Object
    Set shpObj = shp
    shpObj.LockAspectRatio = -1  ' msoTrue (Object 経由なので数値リテラル OK)

    ' 領域内に収まるようリサイズ
    Dim ratioW As Double, ratioH As Double
    ratioW = maxW / shp.Width
    ratioH = maxH / shp.Height
    Dim ratio As Double
    If ratioW < ratioH Then
        ratio = ratioW
    Else
        ratio = ratioH
    End If
    If ratio < 1# Then
        shp.Width = shp.Width * ratio
    End If
    Exit Sub
ErrHandler:
    ' 画像描画失敗時はフォールバックテキストを書く
    On Error Resume Next
    Call WriteFallbackText(ws, topRow, leftCol, botRow, rightCol, imageFull)
End Sub

' ================================================================
' 関数名: WriteFallbackText
' 概要:   画像未配置時の薄字テキスト表示
' ================================================================
Private Sub WriteFallbackText(ByVal ws As Worksheet, _
                                ByVal topRow As Long, ByVal leftCol As Long, _
                                ByVal botRow As Long, ByVal rightCol As Long, _
                                ByVal imageFull As String)
    Dim msg As String
    If imageFull = "" Then
        msg = "[画像未配置]"
    Else
        msg = "[画像未配置: " & imageFull & "]"
    End If
    On Error Resume Next
    ws.Cells(topRow, leftCol).Value = msg
    ws.Cells(topRow, leftCol).Font.Color = RGB(160, 160, 160)
    ws.Cells(topRow, leftCol).Font.Italic = True
End Sub

' ================================================================
' 関数名: ClearShapesByPrefix
' 概要:   指定 prefix で始まる名前の Shape を全削除 (idempotent)
' 引数:   ws     - 対象シート
'         prefix - Shape 名 prefix (例: "kbThumb_")
' 戻り値: なし
' ================================================================
Public Sub ClearShapesByPrefix(ByVal ws As Worksheet, ByVal prefix As String)
    On Error GoTo ErrHandler
    Dim i As Long
    Dim plen As Long
    plen = Len(prefix)
    If plen = 0 Then Exit Sub

    For i = ws.Shapes.Count To 1 Step -1
        Dim shp As Shape
        On Error Resume Next
        Set shp = ws.Shapes(i)
        If Err.Number = 0 Then
            If Len(shp.Name) >= plen Then
                If Left(shp.Name, plen) = prefix Then
                    shp.Delete
                End If
            End If
        End If
        Err.Clear
        On Error GoTo ErrHandler
    Next i
    Exit Sub
ErrHandler:
    ' 削除失敗は致命ではない
End Sub

' ================================================================
' 関数名: HasShapeWithPrefix
' 概要:   指定 prefix の Shape が 1 つ以上存在するか判定 (テスト用)
' ================================================================
Public Function HasShapeWithPrefix(ByVal ws As Worksheet, _
                                     ByVal prefix As String) As Boolean
    On Error GoTo ErrHandler
    Dim i As Long
    Dim plen As Long
    plen = Len(prefix)
    For i = 1 To ws.Shapes.Count
        Dim shp As Shape
        Set shp = ws.Shapes(i)
        If Len(shp.Name) >= plen Then
            If Left(shp.Name, plen) = prefix Then
                HasShapeWithPrefix = True
                Exit Function
            End If
        End If
    Next i
    HasShapeWithPrefix = False
    Exit Function
ErrHandler:
    HasShapeWithPrefix = False
End Function
```
