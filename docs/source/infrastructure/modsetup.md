---
title: modSetup.bas
---

# modSetup.bas

| 項目 | 値 |
|---|---|
| 層 | インストーラ層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 14 シート + 29 ボタンを 1 回実行で自動生成 |
| 行数 | 682 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modSetup` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modSetup"
Option Explicit

' ================================================================
' モジュール: modSetup（インストーラ層）
' 概要:       新規 .xlsm に本システム用の 14 シートと 29 個の
'             フォームコントロールボタンを生成する初期セットアップ。
'             ユーザは VBE で本モジュール + 全 .bas/.cls をインポートし
'             ThisWorkbook の中身を貼り付けた後、本モジュールの
'             SetupSheetsAndButtons を 1 回実行する。
'             再実行しても重複や破壊は発生しない (idempotent)。
' 依存先:     なし (独立モジュール。ボタンの OnAction で参照する
'             マクロ名 (Btn_xxx) は他モジュールに存在することを前提とするが
'             コード中で直接呼び出すことはしないため、コンパイル依存はない)
' 備考:       実行後はブックを保存して再オープンするとシステムが起動する。
' ================================================================

' --- フォームコントロール定数 ---
Private Const XL_BUTTON_CONTROL As Long = 0  ' Excel 標準: xlButtonControl と同値 (m-3, LO 互換のため Const 化を維持)

' --- ボタン寸法ガード (Range が小さすぎる場合の最小値) ---
Private Const BTN_MIN_W As Double = 80#
Private Const BTN_MIN_H As Double = 22#

' --- ログシート ヘッダー位置 ---
Private Const LOG_HEADER_ROW As Long = 1
Private Const LOG_COL_DATETIME As Long = 1
Private Const LOG_COL_MODULE As Long = 2
Private Const LOG_COL_FUNC As Long = 3
Private Const LOG_COL_LEVEL As Long = 4
Private Const LOG_COL_MESSAGE As Long = 5

' --- 設定シート テストモード フラグ位置 ---
Private Const SETTINGS_TESTMODE_ROW As Long = 6
Private Const SETTINGS_TESTMODE_COL As Long = 3

' --- メインシート 案内テキスト位置 ---
Private Const MAIN_GUIDE_ROW As Long = 24
Private Const MAIN_GUIDE_COL As Long = 2

' --- 必要シート名 (順序通りに作成) ---
Private Const REQUIRED_SHEET_NAMES As String = _
    "メイン,フォーマット一覧,フォーマット設計,フォーマットプレビュー," & _
    "ナレッジ登録,ナレッジ修正,ナレッジ一覧,検索,ナレッジ表示," & _
    "格納先設定,設定,既存データへのフィールド反映,データファイル形式,ログ"

' --- メインシート (起動時の唯一可視シート) ---
Private Const SHEET_MAIN_NAME As String = "メイン"
Private Const SHEET_LOG_NAME As String = "ログ"
Private Const SHEET_SETTINGS_NAME As String = "設定"

' --- 既定の空シート名 (削除候補) ---
Private Const DEFAULT_SHEET_NAMES As String = "Sheet1,Sheet2,Sheet3,Sheet4"

' ================================================================
' 関数名: ArrowR
' 概要:   ">" (U+25B6) 文字を返す。
' 備考:   VBA ソースに > リテラルを直接書くと CP932 経路 (PowerShell
'         の Encoding.GetEncoding(932)、Python の cp932 codec 等) で
'         "?" に置換される環境があるため、ChrW で実行時構築する。
'         この関数経由で書けば取り込み時の文字化けを完全回避できる。
' 戻り値: String - ">"
' ================================================================
Private Function ArrowR() As String
    ArrowR = ChrW(&H25B6)
End Function

' ================================================================
' 関数名: ArrowL
' 概要:   "←" (U+2190) 文字を返す。
' 備考:   ← は CP932 で問題なく扱える文字だが、ArrowR との対称性と
'         将来の互換性確保のため同じく ChrW で構築する。
' 戻り値: String - "←"
' ================================================================
Private Function ArrowL() As String
    ArrowL = ChrW(&H2190)
End Function

' ================================================================
' 関数名: SetupSheetsAndButtons
' 概要:   初期セットアップのエントリポイント。
'         シート 14 個 + フォームコントロールボタン 29 個 + 初期表示を
'         一括で構築する。ユーザは本マクロを 1 回実行すればよい。
' 引数:   silent - True 時に完了 MsgBox を抑止 (Workbook_Open auto-init 用)。
'                  既定 False (手動実行時は完了通知を出す)。
' 戻り値: なし
' 備考:   実行後はブックを保存して再オープンするとシステムが起動する。
'         再実行しても重複や破壊は発生しない (同名ボタンは削除→再配置)。
' ================================================================
Public Sub SetupSheetsAndButtons(Optional ByVal silent As Boolean = False)
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Call CreateRequiredSheets
    Call InitializeLogSheetHeader
    Call InitializeSettingsSheet
    Call InitializeMainSheetCells
    Call InitializeSearchSheetHeaders
    Call PlaceAllButtons
    Call SetInitialVisibilityForSetup
    Call DeleteEmptyDefaultSheets

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    If silent Then Exit Sub

    MsgBox "セットアップ完了。" & vbCrLf & _
           "1. ブックを保存してください (Ctrl+S)。" & vbCrLf & _
           "2. 一度ブックを閉じて再オープンすると起動します。", _
           vbInformation, "knowledge_test_v2 セットアップ"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "セットアップに失敗しました:" & vbCrLf & Err.Description, _
           vbCritical, "knowledge_test_v2 セットアップ"
End Sub

' ================================================================
' 関数名: CreateRequiredSheets
' 概要:   REQUIRED_SHEET_NAMES に列挙したシートを順次作成する。
'         既存の同名シートは作成スキップ。
' 引数:   なし
' 戻り値: なし
' ================================================================
Private Sub CreateRequiredSheets()
    Dim names() As String
    names = Split(REQUIRED_SHEET_NAMES, ",")

    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim sheetName As String
        sheetName = Trim$(names(i))
        If Len(sheetName) > 0 Then
            If Not IsSheetExists(sheetName) Then
                Call AppendNewSheet(sheetName)
            End If
        End If
    Next i
End Sub

' ================================================================
' 関数名: AppendNewSheet
' 概要:   ブックの最後尾に新規シートを追加し名前を設定する
' 引数:   sheetName - 作成するシート名
' 戻り値: なし
' ================================================================
Private Sub AppendNewSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
End Sub

' ================================================================
' 関数名: IsSheetExists
' 概要:   指定名のワークシートが ThisWorkbook 内に存在するか判定
' 引数:   sheetName - 確認対象のシート名
' 戻り値: Boolean - 存在すれば True
' ================================================================
Private Function IsSheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    IsSheetExists = (Err.Number = 0 And Not ws Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

' ================================================================
' 関数名: InitializeLogSheetHeader
' 概要:   ログシート 1 行目にヘッダー (日時/モジュール名/関数名/
'         メッセージ種別/メッセージ内容) を書き込む。
'         clsLogger.ClearLog がヘッダー行を保持する仕様のため必須。
' 引数:   なし
' 戻り値: なし
' ================================================================
Private Sub InitializeLogSheetHeader()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG_NAME)
    ws.Cells(LOG_HEADER_ROW, LOG_COL_DATETIME).Value = "日時"
    ws.Cells(LOG_HEADER_ROW, LOG_COL_MODULE).Value = "モジュール名"
    ws.Cells(LOG_HEADER_ROW, LOG_COL_FUNC).Value = "関数名"
    ws.Cells(LOG_HEADER_ROW, LOG_COL_LEVEL).Value = "メッセージ種別"
    ws.Cells(LOG_HEADER_ROW, LOG_COL_MESSAGE).Value = "メッセージ内容"
End Sub

' ================================================================
' 関数名: InitializeSettingsSheet
' 概要:   設定シートのテストモードフラグ (C6) を "FALSE" に設定。
'         空欄でも IsTestMode は False を返すが、明示的に書くことで
'         セルの存在をユーザに見せる目的。
' 引数:   なし
' 戻り値: なし
' ================================================================
Private Sub InitializeSettingsSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS_NAME)
    ws.Cells(1, 1).Value = "設定"
    ws.Cells(SETTINGS_TESTMODE_ROW, 2).Value = "テストモード"
    ws.Cells(SETTINGS_TESTMODE_ROW, SETTINGS_TESTMODE_COL).Value = "FALSE"
End Sub

' ================================================================
' 関数名: InitializeMainSheetCells
' 概要:   メインシートのボタン用セル (B25:E26) に > ラベルを書き込み、
'         案内テキストを行 24 に書き込む。フォームコントロールボタンが
'         上に重なるが、ボタン削除時のラベル復帰用としても機能する。
' 引数:   なし
' 戻り値: なし
' ================================================================
Private Sub InitializeMainSheetCells()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN_NAME)
    ws.Cells(MAIN_GUIDE_ROW, MAIN_GUIDE_COL).Value = _
        "タスクを選択してください (ボタン押下で必要なシートのみ表示)"
    ws.Range("B25").Value = ArrowR & " 初回セットアップ"
    ws.Range("C25").Value = ArrowR & " 設定変更"
    ws.Range("D25").Value = ArrowR & " フォーマット管理"
    ws.Range("E25").Value = ArrowR & " ナレッジ登録"
    ws.Range("B26").Value = ArrowR & " 検索・確認"
    ws.Range("C26").Value = ArrowR & " ナレッジ修正"
    ws.Range("D26").Value = ArrowR & " ナレッジ削除"
    ws.Range("E26").Value = ArrowR & " 既存データ反映"
End Sub

' ================================================================
' 関数名: PlaceAllButtons
' 概要:   全 10 シートに合計 29 個のフォームコントロールボタンを配置。
'         シートごとに PlaceButtonsOnSheet を呼ぶ orchestrator。
' 引数:   なし
' 戻り値: なし
' ================================================================
Private Sub PlaceAllButtons()
    Call PlaceButtonsOnMainSheet
    Call PlaceButtonsOnFormatListSheet
    Call PlaceButtonsOnFormatDesignSheet
    Call PlaceButtonsOnKnowledgeSaveSheet
    Call PlaceButtonsOnKnowledgeEditSheet
    Call PlaceButtonsOnKnowledgeListSheet
    Call PlaceButtonsOnMigrationSheet
    Call PlaceButtonsOnSearchSheet
    Call PlaceButtonsOnKnowledgeDisplaySheet
    Call PlaceButtonsOnSettingsSheet
End Sub

' ================================================================
' 関数名: PlaceButtonsOnMainSheet
' 概要:   メイン シートに 8 個のタスク選択ボタンを配置 (B25:E26)
' ================================================================
Private Sub PlaceButtonsOnMainSheet()
    Dim addrs As Variant
    Dim caps As Variant
    Dim macros As Variant
    addrs = Array("B25", "C25", "D25", "E25", "B26", "C26", "D26", "E26")
    caps = Array(ArrowR & "初回セットアップ", ArrowR & "設定変更", _
                  ArrowR & "フォーマット管理", ArrowR & "ナレッジ登録", _
                  ArrowR & "検索・確認", ArrowR & "ナレッジ修正", _
                  ArrowR & "ナレッジ削除", ArrowR & "既存データ反映")
    macros = Array("Btn_TaskSetup", "Btn_TaskConfig", "Btn_TaskFormat", "Btn_TaskRegister", _
                    "Btn_TaskSearch", "Btn_TaskEdit", "Btn_TaskDelete", "Btn_TaskMigrate")
    Call PlaceButtonsOnSheet("メイン", addrs, caps, macros)
End Sub

' ================================================================
' 関数名: PlaceButtonsOnFormatListSheet
' 概要:   フォーマット一覧 シートに 3 個のボタンを配置 (B1, D1, F1)
' ================================================================
Private Sub PlaceButtonsOnFormatListSheet()
    Dim addrs As Variant
    Dim caps As Variant
    Dim macros As Variant
    addrs = Array("B1", "D1", "F1")
    caps = Array(ArrowR & "新規作成", ArrowR & "選択行を編集", ArrowR & "プレビュー")
    macros = Array("Btn_CreateNewFormat", "Btn_EditFormat", "Btn_ShowFormatPreview")
    Call PlaceButtonsOnSheet("フォーマット一覧", addrs, caps, macros)
End Sub

' ================================================================
' 関数名: PlaceButtonsOnFormatDesignSheet
' 概要:   フォーマット設計 シートに 4 個のボタンを配置 (B3, D3, F3, H3)
' ================================================================
Private Sub PlaceButtonsOnFormatDesignSheet()
    Dim addrs As Variant
    Dim caps As Variant
    Dim macros As Variant
    addrs = Array("B3", "D3", "F3", "H3")
    caps = Array(ArrowR & "読込", ArrowR & "保存", ArrowR & "プレビュー", ArrowL & "一覧に戻る")
    macros = Array("Btn_LoadFormat", "Btn_SaveFormat", "Btn_PreviewFormat", "Btn_BackToList")
    Call PlaceButtonsOnSheet("フォーマット設計", addrs, caps, macros)
End Sub

' ================================================================
' 関数名: PlaceButtonsOnKnowledgeSaveSheet
' 概要:   ナレッジ登録 シートに 2 個のボタンを配置 (B3, C3)
' ================================================================
Private Sub PlaceButtonsOnKnowledgeSaveSheet()
    Dim addrs As Variant
    Dim caps As Variant
    Dim macros As Variant
    addrs = Array("B3", "C3")
    caps = Array(ArrowR & "保存", ArrowR & "クリア")
    macros = Array("Btn_SaveKnowledge", "Btn_ClearForm")
    Call PlaceButtonsOnSheet("ナレッジ登録", addrs, caps, macros)
End Sub

' ================================================================
' 関数名: PlaceButtonsOnKnowledgeEditSheet
' 概要:   ナレッジ修正 シートに 2 個のボタンを配置 (B3, C3)
' ================================================================
Private Sub PlaceButtonsOnKnowledgeEditSheet()
    Dim addrs As Variant
    Dim caps As Variant
    Dim macros As Variant
    addrs = Array("B3", "C3")
    caps = Array(ArrowR & "読込", ArrowR & "上書保存")
    macros = Array("Btn_LoadKnowledge", "Btn_UpdateKnowledge")
    Call PlaceButtonsOnSheet("ナレッジ修正", addrs, caps, macros)
End Sub

' ================================================================
' 関数名: PlaceButtonsOnKnowledgeListSheet
' 概要:   ナレッジ一覧 シートに 2 個のボタンを配置 (B2, E3)
' ================================================================
Private Sub PlaceButtonsOnKnowledgeListSheet()
    Dim addrs As Variant
    Dim caps As Variant
    Dim macros As Variant
    addrs = Array("B2", "E3")
    caps = Array(ArrowR & "リロード", ArrowR & "選択行を削除")
    macros = Array("Btn_ReloadList", "Btn_DeleteKnowledge")
    Call PlaceButtonsOnSheet("ナレッジ一覧", addrs, caps, macros)
End Sub

' ================================================================
' 関数名: PlaceButtonsOnMigrationSheet
' 概要:   既存データへのフィールド反映 シートに 1 個のボタンを配置 (B4)
' ================================================================
Private Sub PlaceButtonsOnMigrationSheet()
    Dim addrs As Variant
    Dim caps As Variant
    Dim macros As Variant
    addrs = Array("B4")
    caps = Array(ArrowR & "反映実行")
    macros = Array("Btn_MigrateFields")
    Call PlaceButtonsOnSheet("既存データへのフィールド反映", addrs, caps, macros)
End Sub

' ================================================================
' 関数名: PlaceButtonsOnSearchSheet
' 概要:   検索 シートに 4 個のボタンを配置 (B12, D3, C12, H17)
' ================================================================
Private Sub PlaceButtonsOnSearchSheet()
    Dim addrs As Variant
    Dim caps As Variant
    Dim macros As Variant
    addrs = Array("B12", "D3", "C12", "H17")
    caps = Array(ArrowR & "検索", ArrowR & "表示", ArrowR & "クリア", ArrowR & "詳細")
    macros = Array("Btn_SearchKnowledge", "Btn_DirectLookup", "Btn_SearchClear", "Btn_DetailDisplay")
    Call PlaceButtonsOnSheet("検索", addrs, caps, macros)
End Sub

' ================================================================
' 関数名: PlaceButtonsOnKnowledgeDisplaySheet
' 概要:   ナレッジ表示 シートに 2 個のボタンを配置 (B9, C9)
' ================================================================
Private Sub PlaceButtonsOnKnowledgeDisplaySheet()
    Dim addrs As Variant
    Dim caps As Variant
    Dim macros As Variant
    addrs = Array("B9", "C9")
    caps = Array(ArrowL & "検索に戻る", ArrowR & "修正に遷移")
    macros = Array("Btn_BackToSearch", "Btn_GoToEdit")
    Call PlaceButtonsOnSheet("ナレッジ表示", addrs, caps, macros)
End Sub

' ================================================================
' 関数名: PlaceButtonsOnSettingsSheet
' 概要:   設定 シートに 1 個のボタンを配置 (E5)
' ================================================================
Private Sub PlaceButtonsOnSettingsSheet()
    Dim addrs As Variant
    Dim caps As Variant
    Dim macros As Variant
    addrs = Array("E5")
    caps = Array(ArrowR & "リセット")
    macros = Array("Btn_ResetLog")
    Call PlaceButtonsOnSheet("設定", addrs, caps, macros)
End Sub

' ================================================================
' 関数名: PlaceButtonsOnSheet
' 概要:   指定シートに複数のフォームコントロールボタンを配置する共通関数。
'         先に既存の同名ボタンを削除してから新規配置することで再実行
'         耐性 (idempotent) を確保。各ボタンは指定セル位置にスナップ。
' 引数:   sheetName - 配置先シート名
'         addrs     - セルアドレス配列 (例: Array("B25","C25"))
'         caps      - ボタンキャプション配列 (addrs と同じ要素数)
'         macros    - OnAction マクロ名配列 (addrs と同じ要素数。
'                     Shape.Name にも使う)
' 戻り値: なし
' ================================================================
Private Sub PlaceButtonsOnSheet(ByVal sheetName As String, _
                                  ByRef addrs As Variant, _
                                  ByRef caps As Variant, _
                                  ByRef macros As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Call DeleteShapesByNames(ws, macros)

    Dim idx As Long
    For idx = LBound(addrs) To UBound(addrs)
        Call PlaceOneButton(ws, CStr(addrs(idx)), CStr(caps(idx)), CStr(macros(idx)))
    Next idx
End Sub

' ================================================================
' 関数名: DeleteShapesByNames
' 概要:   指定シート上の Shapes から、macros 配列に含まれる名前と
'         一致するものを全て削除する。再配置時の重複防止が目的。
' 引数:   ws     - 対象ワークシート
'         macros - 削除対象とする Shape 名 (= マクロ名) の配列
' 戻り値: なし
' ================================================================
Private Sub DeleteShapesByNames(ByVal ws As Worksheet, ByRef macros As Variant)
    Dim toDelete As Collection
    Set toDelete = New Collection

    Call CollectMatchingShapeNames(ws, macros, toDelete)

    Dim i As Long
    For i = 1 To toDelete.Count
        On Error Resume Next
        ws.Shapes(CStr(toDelete.Item(i))).Delete
        Err.Clear
        On Error GoTo 0
    Next i
End Sub

' ================================================================
' 関数名: CollectMatchingShapeNames
' 概要:   ws.Shapes を走査し、macros 配列に含まれる名前と一致した
'         Shape 名を outNames に追加する (反復削除のための事前収集)
' 引数:   ws        - 対象ワークシート
'         macros    - 比較対象のマクロ名配列
'         outNames  - 出力: マッチした Shape 名の Collection (呼び出し側で初期化)
' 戻り値: なし
' ================================================================
Private Sub CollectMatchingShapeNames(ByVal ws As Worksheet, _
                                        ByRef macros As Variant, _
                                        ByRef outNames As Collection)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If IsNameInArray(shp.Name, macros) Then
            outNames.Add shp.Name
        End If
    Next shp
End Sub

' ================================================================
' 関数名: IsNameInArray
' 概要:   target が arr 配列のいずれかの要素と完全一致するかを判定
' 引数:   target - 比較対象の文字列
'         arr    - 文字列の配列 (Variant)
' 戻り値: Boolean - 一致要素があれば True
' ================================================================
Private Function IsNameInArray(ByVal target As String, ByRef arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If target = CStr(arr(i)) Then
            IsNameInArray = True
            Exit Function
        End If
    Next i
    IsNameInArray = False
End Function

' ================================================================
' 関数名: PlaceOneButton
' 概要:   1 個のフォームコントロールボタンを指定セル位置に配置し
'         キャプション・OnAction マクロ名・Shape 名を設定する。
'         セルが小さい場合は最低寸法 (BTN_MIN_W/H) で確保。
' 引数:   ws        - 対象ワークシート
'         cellAddr  - 配置先セルアドレス (例 "B25")
'         caption   - ボタン上に表示するキャプション
'         macroName - OnAction マクロ名 (Shape.Name にも使用)
' 戻り値: なし
' ================================================================
Private Sub PlaceOneButton(ByVal ws As Worksheet, _
                             ByVal cellAddr As String, _
                             ByVal btnCaption As String, _
                             ByVal macroName As String)
    Dim rng As Range
    Set rng = ws.Range(cellAddr)

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

    Dim shp As Shape
    ' v9: Object 型 late binding で AddFormControl を呼ぶ (ADR 0023 / 0024)
    '     XlFormControl Enum 引数の型一致を compile-time で skip
    Dim shapesObj As Object
    Set shapesObj = ws.Shapes
    Set shp = shapesObj.AddFormControl(XL_BUTTON_CONTROL, leftPt, topPt, widthPt, heightPt)
    shp.Name = macroName

    Call SetButtonCaptionAndAction(shp, btnCaption, macroName)
End Sub

' ================================================================
' 関数名: SetButtonCaptionAndAction
' 概要:   フォームコントロールボタンにキャプションと OnAction マクロを
'         設定する。OLEFormat.Object 経由が最も確実だが環境差を考慮し
'         TextFrame 系も保険として試みる。
' 引数:   shp       - フォームコントロールボタンの Shape
'         caption   - 表示するキャプション
'         macroName - OnAction マクロ名
' 戻り値: なし
' ================================================================
Private Sub SetButtonCaptionAndAction(ByVal shp As Shape, _
                                        ByVal btnCaption As String, _
                                        ByVal macroName As String)
    ' m-4: 2 つのフォールバックパスを統合、最終結果を取得
    Dim oleSet As Boolean
    oleSet = False

    On Error Resume Next
    shp.OLEFormat.Object.Caption = btnCaption
    shp.OLEFormat.Object.OnAction = macroName
    oleSet = (Err.Number = 0)
    Err.Clear

    ' 保険として Shape.OnAction も試す (環境差吸収)
    shp.OnAction = macroName
    Err.Clear
    On Error GoTo 0

    ' 両方失敗したら Debug.Print で警告 (ロガー依存を作らずに告知)
    If Not oleSet Then
        Debug.Print "[SetButtonCaptionAndAction] WARN: ボタン " & macroName & _
                    " のキャプション/OnAction 設定に失敗 (環境依存)"
    End If
End Sub

' ================================================================
' 関数名: SetInitialVisibilityForSetup
' 概要:   メインシートのみを表示し、その他全シートを非表示にする。
'         (起動時の Workbook_Open でも実施されるが、保存状態を整える
'          目的で本セットアップでも実行する)
' 引数:   なし
' 戻り値: なし
' ================================================================
Private Sub SetInitialVisibilityForSetup()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_MAIN_NAME Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetHidden
        End If
    Next ws
End Sub

' ================================================================
' 関数名: DeleteEmptyDefaultSheets
' 概要:   Sheet1 / Sheet2 / Sheet3 / Sheet4 のうち、データもシェイプも
'         無いものを削除する。新規 .xlsm に最初から存在する空 Sheet1
'         を整理する目的。データやシェイプがある場合は保持する。
' 引数:   なし
' 戻り値: なし
' ================================================================
Private Sub DeleteEmptyDefaultSheets()
    Dim names() As String
    names = Split(DEFAULT_SHEET_NAMES, ",")

    Dim i As Long
    For i = LBound(names) To UBound(names)
        Call DeleteSheetIfEmpty(Trim$(names(i)))
    Next i
End Sub

' ================================================================
' 関数名: DeleteSheetIfEmpty
' 概要:   指定名のシートが存在し、かつデータ・シェイプとも空である場合
'         のみ削除する。
' 引数:   sheetName - 対象シート名
' 戻り値: なし
' ================================================================
Private Sub DeleteSheetIfEmpty(ByVal sheetName As String)
    If Len(sheetName) = 0 Then Exit Sub
    If Not IsSheetExists(sheetName) Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not IsSheetEmpty(ws) Then Exit Sub

    On Error Resume Next
    ws.Visible = xlSheetVisible  ' 削除前に可視化が必要
    ws.Delete
    Err.Clear
    On Error GoTo 0
End Sub

' ================================================================
' 関数名: IsSheetEmpty
' 概要:   指定シートにデータセル・シェイプとも無いか判定
' 引数:   ws - 判定対象のワークシート
' 戻り値: Boolean - 空なら True
' ================================================================
Private Function IsSheetEmpty(ByVal ws As Worksheet) As Boolean
    Dim cellCount As Long
    cellCount = WorksheetFunction.CountA(ws.Cells)

    Dim shapeCount As Long
    shapeCount = ws.Shapes.Count

    IsSheetEmpty = (cellCount = 0 And shapeCount = 0)
End Function

' ================================================================
' 関数名: InitializeSearchSheetHeaders
' 概要:   検索シートの結果一覧 (15 行目から開始) のヘッダ行 (14 行目)
'         を 9 列構成で書き込む。image_ext rev1 で追加。
'         A14: # / B14: 番号 / C14: フォーマット / D14: タイトル
'         E14: 作成日 / F14: 更新日 / G14: 詳細 / H14: サムネ / I14: Score
' 備考:   既存テスト 82 件は row 15+ の結果値のみ参照しており、
'         row 14 のヘッダ追加は破壊的変更にならない。
' ================================================================
Private Sub InitializeSearchSheetHeaders()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)

    ws.Cells(14, 1).Value = "#"
    ws.Cells(14, 2).Value = "番号"
    ws.Cells(14, 3).Value = "フォーマット"
    ws.Cells(14, 4).Value = "タイトル"
    ws.Cells(14, 5).Value = "作成日"
    ws.Cells(14, 6).Value = "更新日"
    ws.Cells(14, 7).Value = "詳細"
    ws.Cells(14, 8).Value = "サムネ"
    ws.Cells(14, 9).Value = "Score"

    With ws.Range(ws.Cells(14, 1), ws.Cells(14, 9))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    ws.Columns(1).ColumnWidth = 4
    ws.Columns(2).ColumnWidth = 14
    ws.Columns(3).ColumnWidth = 16
    ws.Columns(4).ColumnWidth = 30
    ws.Columns(5).ColumnWidth = 12
    ws.Columns(6).ColumnWidth = 12
    ws.Columns(7).ColumnWidth = 8
    ws.Columns(8).ColumnWidth = 10
    ws.Columns(9).ColumnWidth = 8
    Exit Sub
ErrHandler:
    ' C-3: 失敗を黙殺せず可視化する。Debug.Print + シートタブ赤化 + 警告セル
    Debug.Print "[InitializeSearchSheetHeaders] ERROR: " & Err.Description
    On Error Resume Next
    Dim wsErr As Worksheet
    Set wsErr = ThisWorkbook.Worksheets(SHEET_SEARCH)
    If Not wsErr Is Nothing Then
        wsErr.Tab.Color = RGB(255, 0, 0)
        wsErr.Cells(13, 1).Value = "[警告] ヘッダー初期化失敗: " & Err.Description
    End If
End Sub
```

## 関連

- 呼び出す: `modCommon`, `modEntrySearch`, `modEntryKnowledge`, `modEntryFormat`, `modEntrySettings`
- 呼び出される: `ユーザが Alt+F8 で 1 回実行`
