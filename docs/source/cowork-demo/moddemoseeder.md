---
title: modDemoSeeder.bas
---

# modDemoSeeder.bas

| 項目 | 値 |
|---|---|
| 層 | Cowork デモ専用 |
| 種別 | 標準モジュール (.bas) |
| 役割 | Cowork デモ用ワンクリック初期化 (フォーマット / dataFolder / ナレッジ 5 件) |
| 行数 | 493 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modDemoSeeder` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modDemoSeeder"

Option Explicit



' ================================================================

' モジュール: modDemoSeeder (Cowork デモ専用 / 配布版には含まれない)

' 概要:       knowledge_test_v2_demo.xlsm 用のワンクリックデモ初期化。

'               (1) フォーマット一覧 row 3 に "DEMO-MEMO" を登録

'               (2) 設定 row 3 col C (dataFolder) を ThisWorkbook.Path & "\data"

'               (3) ナレッジ一覧 row 3 - 7 に KN-2026-0420 - 0424 の 5 件

'                   をスコア降順で見せるためのプレースホルダ行

' 依存先:     なし (modCommon の SHEET_* / SETTINGS_ROW_DATAFOLDER /

'             SETTINGS_COL_VALUE 定数を参照)

' 備考:       本モジュールは「Cowork 環境内で構築したモック」専用。

'             SetupSheetsAndButtons の後に Alt+F8 → Macro_SeedDemoData を

'             1 回実行することで、検索・画像表示・スコア降順のデモが

'             即座に動作する状態になる。

' ================================================================



' D-2 注記: フォーマット一覧 列番号は clsFormatManager.cls の Public Const FL_COL_* を参照 (重複定義を解消)

Private Const DS_FL_HEADER_ROW As Long = 2



' フォーマット設計 列番号 (FD_*)

Private Const DS_FD_ROW_FMT_ID As Long = 1

Private Const DS_FD_ROW_FMT_NAME As Long = 2

Private Const DS_FD_FIELDS_HEADER_ROW As Long = 5

Private Const DS_FD_FIELDS_START_ROW As Long = 6



' ナレッジ一覧 列番号

Private Const DS_KL_HEADER_ROW As Long = 2

Private Const DS_KL_DATA_ROW As Long = 3

Private Const DS_KL_COL_NO As Long = 1

Private Const DS_KL_COL_KNW_NO As Long = 2

Private Const DS_KL_COL_FMT_ID As Long = 3

Private Const DS_KL_COL_TITLE As Long = 4

Private Const DS_KL_COL_CREATED As Long = 5

Private Const DS_KL_COL_UPDATED As Long = 6



' ================================================================

' 関数名: Macro_SeedDemoData

' 概要:   Cowork デモ用 1-shot 初期化マクロ。

'         SetupSheetsAndButtons の後に 1 回だけ実行する。

' 引数:   なし

' 戻り値: なし

' 備考:   再実行しても安全 (上書き)。

' ================================================================

Public Sub Macro_SeedDemoData(Optional ByVal silent As Boolean = False)

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    Application.DisplayAlerts = False



    Call SeedFormatList

    Call SeedFormatDesign

    Call SeedDataFolderSetting

    Call SeedKnowledgeListPlaceholder

    Call SeedSearchSheetGuide



    Application.ScreenUpdating = True

    Application.DisplayAlerts = True



    If silent Then Exit Sub



    Dim dataPath As String

    dataPath = ThisWorkbook.Path & "\data"



    MsgBox "デモデータの初期化が完了しました。" & vbCrLf & vbCrLf & _

           "■フォーマット一覧 row 3: DEMO-MEMO 登録済み" & vbCrLf & _

           "■設定 row 3 (dataFolder): " & dataPath & vbCrLf & _

           "■ナレッジ一覧 row 3-7: KN-2026-0420 - 0424 (5 件)" & vbCrLf & _

           "■検索シート 14 行目ヘッダ: 9 列で初期化済み" & vbCrLf & vbCrLf & _

           "次に行うこと:" & vbCrLf & _

           "1. メインシートの『検索』ボタン → 検索シート" & vbCrLf & _

           "2. キーワードに 「メモリ」 等を入力 → 検索ボタン" & vbCrLf & _

           "3. 結果がスコア降順で表示され、画像列にサムネが入ります", _

           vbInformation, "knowledge_test_v2 デモ初期化"

    Exit Sub



ErrHandler:

    Application.ScreenUpdating = True

    Application.DisplayAlerts = True

    MsgBox "デモ初期化に失敗しました:" & vbCrLf & Err.Description, _

           vbCritical, "knowledge_test_v2 デモ初期化"

End Sub



' ================================================================

' 関数名: SeedFormatList

' 概要:   フォーマット一覧 row 3 に DEMO-MEMO を登録

' ================================================================

Private Sub SeedFormatList()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)



    ' ヘッダー (row 2) を念のため再書込

    ws.Cells(DS_FL_HEADER_ROW, FL_COL_NO).Value = "No"

    ws.Cells(DS_FL_HEADER_ROW, FL_COL_FMT_ID).Value = "FormatID"

    ws.Cells(DS_FL_HEADER_ROW, FL_COL_FMT_NAME).Value = "FormatName"

    ws.Cells(DS_FL_HEADER_ROW, FL_COL_ID_PATTERN).Value = "IDPattern"

    ws.Cells(DS_FL_HEADER_ROW, FL_COL_CURRENT_NUM).Value = "CurrentNum"

    ws.Cells(DS_FL_HEADER_ROW, FL_COL_NEXT_NUM).Value = "NextNum"

    ws.Cells(DS_FL_HEADER_ROW, FL_COL_VERSION).Value = "Version"

    ws.Cells(DS_FL_HEADER_ROW, FL_COL_FIELD_COUNT).Value = "FieldCount"

    ws.Cells(DS_FL_HEADER_ROW, FL_COL_KNW_COUNT).Value = "KnowledgeCount"

    ws.Cells(DS_FL_HEADER_ROW, FL_COL_CREATED).Value = "CreatedDate"

    ws.Cells(DS_FL_HEADER_ROW, FL_COL_UPDATED).Value = "UpdatedDate"



    ' データ行 (row 3) を上書き登録

    ws.Cells(FL_START_ROW, FL_COL_NO).Value = 1

    ws.Cells(FL_START_ROW, FL_COL_FMT_ID).Value = "DEMO-MEMO"

    ws.Cells(FL_START_ROW, FL_COL_FMT_NAME).Value = "メモ"

    ws.Cells(FL_START_ROW, FL_COL_ID_PATTERN).Value = "KN-yyyy-NNNN"

    ws.Cells(FL_START_ROW, FL_COL_CURRENT_NUM).Value = 424

    ws.Cells(FL_START_ROW, FL_COL_NEXT_NUM).Value = 425

    ws.Cells(FL_START_ROW, FL_COL_VERSION).Value = 1

    ws.Cells(FL_START_ROW, FL_COL_FIELD_COUNT).Value = 4

    ws.Cells(FL_START_ROW, FL_COL_KNW_COUNT).Value = 5

    ws.Cells(FL_START_ROW, FL_COL_CREATED).Value = "2026-04-20"

    ws.Cells(FL_START_ROW, FL_COL_UPDATED).Value = "2026-05-04"

End Sub



' ================================================================

' 関数名: SeedFormatDesign

' 概要:   フォーマット設計シートに DEMO-MEMO の 4 フィールドを書く

' ================================================================

Private Sub SeedFormatDesign()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)



    ws.Cells(DS_FD_ROW_FMT_ID, 1).Value = "FormatID"

    ws.Cells(DS_FD_ROW_FMT_ID, 2).Value = "DEMO-MEMO"

    ws.Cells(DS_FD_ROW_FMT_ID, 3).Value = "Version"

    ws.Cells(DS_FD_ROW_FMT_ID, 4).Value = 1

    ws.Cells(DS_FD_ROW_FMT_ID, 5).Value = "IDPattern"

    ws.Cells(DS_FD_ROW_FMT_ID, 6).Value = "KN-yyyy-NNNN"

    ws.Cells(DS_FD_ROW_FMT_NAME, 1).Value = "FormatName"

    ws.Cells(DS_FD_ROW_FMT_NAME, 2).Value = "メモ"



    ws.Cells(DS_FD_FIELDS_HEADER_ROW, 1).Value = "FieldNo"

    ws.Cells(DS_FD_FIELDS_HEADER_ROW, 2).Value = "FieldName"

    ws.Cells(DS_FD_FIELDS_HEADER_ROW, 3).Value = "Type"

    ws.Cells(DS_FD_FIELDS_HEADER_ROW, 4).Value = "Required"

    ws.Cells(DS_FD_FIELDS_HEADER_ROW, 5).Value = "Rows"



    ' 4 フィールド: タイトル / 内容 / 日付 / 担当

    Call WriteFieldRow(ws, DS_FD_FIELDS_START_ROW + 0, 1, _

                        "タイトル", FIELD_TYPE_STRING, "○", 1)

    Call WriteFieldRow(ws, DS_FD_FIELDS_START_ROW + 1, 2, _

                        "内容", FIELD_TYPE_LONGTEXT, "○", 6)

    Call WriteFieldRow(ws, DS_FD_FIELDS_START_ROW + 2, 3, _

                        "日付", FIELD_TYPE_DATE, "", 1)

    Call WriteFieldRow(ws, DS_FD_FIELDS_START_ROW + 3, 4, _

                        "担当", FIELD_TYPE_STRING, "", 1)

End Sub



Private Sub WriteFieldRow(ByVal ws As Worksheet, ByVal r As Long, _

                            ByVal no As Long, ByVal name_ As String, _

                            ByVal type_ As String, ByVal required As String, _

                            ByVal rows_ As Long)

    ws.Cells(r, 1).Value = no

    ws.Cells(r, 2).Value = name_

    ws.Cells(r, 3).Value = type_

    ws.Cells(r, 4).Value = required

    ws.Cells(r, 5).Value = rows_

End Sub



' ================================================================

' 関数名: SeedDataFolderSetting

' 概要:   設定シート row 3 col C (dataFolder) を ThisWorkbook.Path & "\data" に

' ================================================================

Private Sub SeedDataFolderSetting()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)



    ws.Cells(SETTINGS_ROW_DATAFOLDER, 2).Value = "dataFolder"

    ws.Cells(SETTINGS_ROW_DATAFOLDER, SETTINGS_COL_VALUE).NumberFormat = "@"

    ws.Cells(SETTINGS_ROW_DATAFOLDER, SETTINGS_COL_VALUE).Value = _

        ThisWorkbook.Path & "\data"



    ws.Cells(SETTINGS_ROW_CHARSET, 2).Value = "Charset"

    ws.Cells(SETTINGS_ROW_CHARSET, SETTINGS_COL_VALUE).Value = "Shift_JIS"



    ws.Cells(SETTINGS_ROW_DEBUGLEVEL, 2).Value = "DebugLevel"

    ws.Cells(SETTINGS_ROW_DEBUGLEVEL, SETTINGS_COL_VALUE).Value = "OFF"

End Sub



' ================================================================

' 関数名: SeedKnowledgeListPlaceholder

' 概要:   ナレッジ一覧 row 3-7 に 5 件の見出しを書く (検索より前に

'         可視性を上げるためのプレビュー)

' ================================================================

Private Sub SeedKnowledgeListPlaceholder()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)



    ws.Cells(DS_KL_HEADER_ROW, DS_KL_COL_NO).Value = "No"

    ws.Cells(DS_KL_HEADER_ROW, DS_KL_COL_KNW_NO).Value = "KnowledgeNo"

    ws.Cells(DS_KL_HEADER_ROW, DS_KL_COL_FMT_ID).Value = "FormatID"

    ws.Cells(DS_KL_HEADER_ROW, DS_KL_COL_TITLE).Value = "タイトル"

    ws.Cells(DS_KL_HEADER_ROW, DS_KL_COL_CREATED).Value = "CreatedDate"

    ws.Cells(DS_KL_HEADER_ROW, DS_KL_COL_UPDATED).Value = "UpdatedDate"



    Dim arr As Variant

    arr = Array( _

        Array("KN-2026-0420", "メモリ枯渇エラー対処メモ"), _

        Array("KN-2026-0421", "ChromaDB HNSW 再構築手順"), _

        Array("KN-2026-0422", "VBA ADODB.Stream の代替"), _

        Array("KN-2026-0423", "RAG 検索のスコアリング設計"), _

        Array("KN-2026-0424", "サムネ画像の自動配置メモ") _

    )



    Dim i As Long

    For i = 0 To UBound(arr)

        Dim r As Long

        r = DS_KL_DATA_ROW + i

        ws.Cells(r, DS_KL_COL_NO).Value = i + 1

        ws.Cells(r, DS_KL_COL_KNW_NO).NumberFormat = "@"

        ws.Cells(r, DS_KL_COL_KNW_NO).Value = CStr(arr(i)(0))

        ws.Cells(r, DS_KL_COL_FMT_ID).Value = "DEMO-MEMO"

        ws.Cells(r, DS_KL_COL_TITLE).Value = CStr(arr(i)(1))

        ws.Cells(r, DS_KL_COL_CREATED).Value = "2026-04-" & Format(20 + i, "00")

        ws.Cells(r, DS_KL_COL_UPDATED).Value = "2026-04-" & Format(20 + i, "00")

    Next i

End Sub



' ================================================================

' 関数名: SeedSearchSheetGuide

' 概要:   検索シート上部 (row 4 - 12) に簡易ガイドを書き、ユーザに

'         どこに何を入力すれば検索できるかを案内する

' ================================================================

Private Sub SeedSearchSheetGuide()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)



    ws.Cells(4, 2).Value = "■デモ検索ガイド (Cowork mock)"

    ws.Cells(5, 2).Value = "FormatID:"

    ws.Cells(5, 3).NumberFormat = "@"

    ws.Cells(5, 3).Value = "DEMO-MEMO"

    ws.Cells(6, 2).Value = "キーワード:"

    ws.Cells(6, 3).Value = "メモリ"

    ws.Cells(7, 2).Value = "検索モード:"

    ws.Cells(7, 3).Value = "AND"

    ws.Cells(8, 2).Value = "TargetField:"

    ws.Cells(8, 3).Value = "(全フィールド)"

    ws.Cells(9, 2).Value = "推奨操作"

    ws.Cells(10, 2).Value = "→ メインシートで「検索」ボタン → このシートで「検索実行」ボタン"

    ws.Cells(11, 2).Value = "→ 結果は 14 行目以降に Score 降順で並ぶ"

    ws.Cells(12, 2).Value = "→ 画像列 (H 列) に 60×40px サムネ"

End Sub

        
```

## 関連

- 呼び出す: `modCommon`, `clsFormatManager`
- 呼び出される: `modAutoInit`
