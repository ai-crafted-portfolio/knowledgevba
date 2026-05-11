---
title: modScreenSpecRegistry.bas
---

# modScreenSpecRegistry.bas

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | M-01〜M-14 の各画面 spec をハードコードで定義するレジストリ |
| 行数 | 556 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modScreenSpecRegistry` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modScreenSpecRegistry"
Option Explicit

' ================================================================
' モジュール: modScreenSpecRegistry（画面層 — レジストリ）
' 概要:   各画面の clsScreenSpec を構築する。spec はデータ駆動なので
'         画面修正は本モジュールの Build*Spec 関数を編集するだけで完結する。
'         （コード本体 = 各 clsXxxScreen.cls は触らない）
' 依存先: clsScreenSpec, clsSectionSpec, clsButtonSpec, clsFieldSpec, modCommon
' 備考:   v20 改修: polished mock M-01〜M-14 (設計書_v19.xlsx) 準拠。
' ================================================================

' ================================================================
' 関数名: GetScreenSpec
' 概要:   screenId に対応する画面 spec を返す（ファクトリ関数）
' 引数:   screenId - "M-01" 〜 "M-14"
' 戻り値: clsScreenSpec
' ================================================================
Public Function GetScreenSpec(ByVal screenId As String) As clsScreenSpec
    Select Case screenId
        Case "M-01": Set GetScreenSpec = BuildMainSpec()
        Case "M-02": Set GetScreenSpec = BuildFormatListSpec()
        Case "M-03": Set GetScreenSpec = BuildFormatDesignSpec()
        Case "M-04": Set GetScreenSpec = BuildFormatPreviewSpec()
        Case "M-05": Set GetScreenSpec = BuildKnowledgeRegisterSpec()
        Case "M-06": Set GetScreenSpec = BuildKnowledgeEditSpec()
        Case "M-07": Set GetScreenSpec = BuildKnowledgeListSpec()
        Case "M-08": Set GetScreenSpec = BuildSearchSpec()
        Case "M-09": Set GetScreenSpec = BuildKnowledgeViewSpec()
        Case "M-10": Set GetScreenSpec = BuildStorageSpec()
        Case "M-11": Set GetScreenSpec = BuildSystemSettingsSpec()
        Case "M-12": Set GetScreenSpec = BuildMigrationSpec()
        Case "M-13": Set GetScreenSpec = BuildFileFormatSpec()
        Case "M-14": Set GetScreenSpec = BuildLogSpec()
    End Select
End Function

' ================================================================
' 関数名: GetAllScreenIds
' 概要:   全 14 画面の ID 配列を返す
' ================================================================
Public Function GetAllScreenIds() As Variant
    GetAllScreenIds = Array("M-01", "M-02", "M-03", "M-04", "M-05", "M-06", "M-07", _
                              "M-08", "M-09", "M-10", "M-11", "M-12", "M-13", "M-14")
End Function

' ================================================================
' 関数名: BuildMainSpec
' 概要:   M-01 メイン画面の spec を構築（12 ボタン × 3 グループ）
' ================================================================
Private Function BuildMainSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-01"
    s.SheetName = SHEET_MAIN
    s.Title = "[v2] メイン (12 タスクボタンに整理)"
    s.TitleColorHex = COLOR_TITLE_DEEP_BLUE
    s.BackToMainAddr = ""

    ' グループ帯（A9 / A13 / A17）
    Call AddSec(s, "A9", "[主操作 - 青] 日常使う機能 (青ボタン)", COLOR_SECTION_BLUE)
    Call AddSec(s, "A13", "[遷移 - 緑] 管理者向け機能 (緑ボタン)", COLOR_SECTION_BLUE)
    Call AddSec(s, "A17", "[その他] 確認/初期セットアップ (緑/灰ボタン)", COLOR_SECTION_BLUE)

    ' 行 10/11: 主操作 4 ボタン + 説明
    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_TaskSearch",       ar & "検索",         "B10", COLOR_BTN_PRIMARY, "主操作", "B11", "キーワードで" & vbLf & "検索・確認")
    Call AddBtn(s, "Btn_TaskRegister",     ar & "ナレッジ登録", "E10", COLOR_BTN_PRIMARY, "主操作", "E11", "新規ナレッジを" & vbLf & "登録")
    Call AddBtn(s, "Btn_TaskModify",       ar & "ナレッジ修正", "H10", COLOR_BTN_PRIMARY, "主操作", "H11", "既存ナレッジを" & vbLf & "修正")
    Call AddBtn(s, "Btn_TaskList",         ar & "ナレッジ一覧", "K10", COLOR_BTN_PRIMARY, "主操作", "K11", "全件閲覧/削除")

    ' 行 14/15: 遷移 4 ボタン + 説明
    Call AddBtn(s, "Btn_TaskFormat",        ar & "フォーマット管理", "B14", COLOR_BTN_NAV, "遷移", "B15", "カテゴリの追加・編集")
    Call AddBtn(s, "Btn_TaskFieldReflect",  ar & "フィールド反映",   "E14", COLOR_BTN_NAV, "遷移", "E15", "フォーマット変更を" & vbLf & "全データに反映")
    Call AddBtn(s, "Btn_TaskStorage",       ar & "格納先設定",       "H14", COLOR_BTN_NAV, "遷移", "H15", "共有/BOX 等の" & vbLf & "格納先設定")
    Call AddBtn(s, "Btn_TaskSysSettings",   ar & "システム設定",     "K14", COLOR_BTN_NAV, "遷移", "K15", "データフォルダ等")

    ' 行 18/19: その他 4 ボタン + 説明
    Call AddBtn(s, "Btn_TaskLog",           ar & "ログ確認",          "B18", COLOR_BTN_NAV, "その他", "B19", "操作ログ閲覧" & vbLf & "(エクスポート可)")
    Call AddBtn(s, "Btn_TaskFileFormat",    ar & "ファイル形式",      "E18", COLOR_BTN_NAV, "その他", "E19", "ファイル形式の" & vbLf & "仕様 (リファレンス)")
    Call AddBtn(s, "Btn_TaskInitSetup",     ar & "初回セットアップ",  "H18", COLOR_BTN_NAV, "その他", "H19", "管理者の" & vbLf & "初回設定ガイド")
    Call AddBtn(s, "Btn_TaskHelpVersion",   ar & "ヘルプ",            "K18", COLOR_BTN_SUB, "その他", "K19", "バージョン情報")

    Set BuildMainSpec = s
End Function

' ================================================================
' 関数名: BuildFormatListSpec (M-02)
' ================================================================
Private Function BuildFormatListSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-02"
    s.SheetName = SHEET_FORMAT_LIST
    s.Title = "[v2] M-02 フォーマット一覧 (新規/編集/プレビュー/削除)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K6"

    Call AddSec(s, "A4", "[アクション]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A9", "[フォーマット一覧]", COLOR_SECTION_BLUE2)

    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_CreateNewFormat",   ar & "新規作成",   "B7", COLOR_BTN_PRIMARY, "アクション", "", "")
    Call AddBtn(s, "Btn_EditFormat",        ar & "編集",       "E7", COLOR_BTN_PRIMARY, "アクション", "", "")
    Call AddBtn(s, "Btn_ShowFormatPreview", ar & "プレビュー", "G7", COLOR_BTN_SUB,     "アクション", "", "")
    Call AddBtn(s, "Btn_DeleteFormat",      ar & "削除",       "I7", COLOR_BTN_DANGER,  "アクション", "", "")

    s.HeaderRowAddr = "B10"
    s.HeaderLabels = Array("No", "FormatID", "フォーマット名", "ID形式", "次の番号", "Ver", "F数", "K数", "更新日", "状態")
    s.EmptyStateAddr = "B11"
    s.EmptyStateText = "(フォーマット未登録 ― 0 件)"

    Set BuildFormatListSpec = s
End Function

' ================================================================
' 関数名: BuildFormatDesignSpec (M-03)
' ================================================================
Private Function BuildFormatDesignSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-03"
    s.SheetName = SHEET_FORMAT_DESIGN
    s.Title = "[v2] M-03 フォーマット設計 (連動の起点)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K20"

    Call AddSec(s, "A3", "[対象フォーマット]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A5", "[フィールド定義 (この順で M-04〜M-09 に並ぶ)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A16", "[フィールド追加]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A19", "[アクション]", COLOR_SECTION_BLUE2)

    ' フィールド定義テーブルのヘッダ
    s.HeaderRowAddr = "A6"
    s.HeaderLabels = Array("並び", "フィールド名", "型", "必須", "行数", "選択肢/既定値")

    ' 標準雛形フィールドラベル（データ無しでも見える）
    Call AddStandardFieldRows(s, 7)

    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_LoadFormat",     ar & "読込",       "B20", COLOR_BTN_NAV,     "アクション", "", "")
    Call AddBtn(s, "Btn_SaveFormat",     ar & "保存",       "D20", COLOR_BTN_PRIMARY, "アクション", "", "")
    Call AddBtn(s, "Btn_PreviewFormat",  ar & "プレビュー", "F20", COLOR_BTN_SUB,     "アクション", "", "")
    Call AddBtn(s, "Btn_AddField",       "+ フィールド追加", "B17", COLOR_BTN_NAV,     "追加", "", "")

    Set BuildFormatDesignSpec = s
End Function

' ================================================================
' 関数名: BuildFormatPreviewSpec (M-04)
' ================================================================
Private Function BuildFormatPreviewSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-04"
    s.SheetName = SHEET_FORMAT_PREVIEW
    s.Title = "[v2] M-04 プレビュー (M-03 で設計した内容を入力フォームとして再現)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K17"

    Call AddSec(s, "A3", "[プレビューモード]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A5", "[入力フォーム (M-03 の定義通り)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A16", "[アクション]", COLOR_SECTION_BLUE2)

    Call AddStandardFieldRows(s, 6)

    Dim ar As String: ar = ChrW(&H25B6)
    Dim al As String: al = ChrW(&H2190)
    Call AddBtn(s, "Btn_BackToFormatDesign", al & " M-03 設計に戻る", "D17", COLOR_HEADER_LIGHT, "ナビ", "", "")

    Set BuildFormatPreviewSpec = s
End Function

' ================================================================
' 関数名: BuildKnowledgeRegisterSpec (M-05)
' ================================================================
Private Function BuildKnowledgeRegisterSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-05"
    s.SheetName = SHEET_KNW_SAVE
    s.Title = "[v2] M-05 ナレッジ登録 (M-03 で定義したフォーマットに基づく)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K19"

    Call AddSec(s, "A3", "[モード/予定番号]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A5", "[フォーマット選択]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A7", "[入力フォーム (動的生成 ― M-03 定義に基づく)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A18", "[アクション]", COLOR_SECTION_BLUE2)

    Call AddStandardFieldRows(s, 8)

    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_SaveKnowledge", ar & "登録",     "B19", COLOR_BTN_PRIMARY, "アクション", "", "")
    Call AddBtn(s, "Btn_ClearForm",     ar & "クリア",   "D19", COLOR_BTN_SUB,     "アクション", "", "")

    Set BuildKnowledgeRegisterSpec = s
End Function

' ================================================================
' 関数名: BuildKnowledgeEditSpec (M-06)
' ================================================================
Private Function BuildKnowledgeEditSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-06"
    s.SheetName = SHEET_KNW_EDIT
    s.Title = "[v2] M-06 ナレッジ修正 (M-03 定義のフォーマットに既存値を表示)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K19"

    Call AddSec(s, "A3", "[モード/対象]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A5", "[フォーマット (修正不可)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A7", "[入力フォーム (既存値表示・編集可)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A18", "[アクション]", COLOR_SECTION_BLUE2)

    Call AddStandardFieldRows(s, 8)

    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_LoadKnowledge",   ar & "読込",     "B19", COLOR_BTN_NAV,     "アクション", "", "")
    Call AddBtn(s, "Btn_UpdateKnowledge", ar & "更新",     "D19", COLOR_BTN_PRIMARY, "アクション", "", "")
    Call AddBtn(s, "Btn_DeleteKnowledge", ar & "このナレッジを削除", "G19", COLOR_BTN_DANGER, "アクション", "", "")

    Set BuildKnowledgeEditSpec = s
End Function

' ================================================================
' 関数名: BuildKnowledgeListSpec (M-07)
' ================================================================
Private Function BuildKnowledgeListSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-07"
    s.SheetName = SHEET_KNW_LIST
    s.Title = "[v2] M-07 ナレッジ一覧 (修正ボタン, ページング)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K15"

    Call AddSec(s, "A2", "[絞込]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A4", "[一覧 (0 件)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A11", "[ページング]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A14", "[ナビゲーション]", COLOR_SECTION_BLUE2)

    s.HeaderRowAddr = "A5"
    s.HeaderLabels = Array("#", "ナレッジ番号", "フォーマット", "登録日", "更新日", "件名", "カテゴリ", "担当者", "表示", "修正", "削除")
    s.EmptyStateAddr = "A6"
    s.EmptyStateText = "(ナレッジ未登録 ― 0 件)"

    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_ReloadList",      ar & "絞込",     "I3", COLOR_BTN_PRIMARY, "絞込", "", "")
    Call AddBtn(s, "Btn_PageFirst",       "<<最初",        "B12", COLOR_BTN_SUB, "ページング", "", "")
    Call AddBtn(s, "Btn_PagePrev",        "<前",           "D12", COLOR_BTN_SUB, "ページング", "", "")
    Call AddBtn(s, "Btn_PageNext",        "次>",           "G12", COLOR_BTN_SUB, "ページング", "", "")
    Call AddBtn(s, "Btn_PageLast",        "最後>>",        "I12", COLOR_BTN_SUB, "ページング", "", "")

    Set BuildKnowledgeListSpec = s
End Function

' ================================================================
' 関数名: BuildSearchSpec (M-08)
' ================================================================
Private Function BuildSearchSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-08"
    s.SheetName = SHEET_SEARCH
    s.Title = "[v2] M-08 ナレッジ検索 (キーワード + フィルタ)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K13"

    Call AddSec(s, "A2", "[検索条件]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A6", "[検索結果 (0 件)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A12", "[ナビゲーション]", COLOR_SECTION_BLUE2)

    s.HeaderRowAddr = "A7"
    s.HeaderLabels = Array("#", "ナレッジ番号", "フォーマット", "登録日", "更新日", "件名/事象", "カテゴリ", "担当者", "表示", "修正", "削除")
    s.EmptyStateAddr = "A8"
    s.EmptyStateText = "(検索結果なし ― キーワード入力後に検索ボタンを押下)"

    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_SearchKnowledge", ar & "検索",       "I3", COLOR_BTN_PRIMARY, "条件", "", "")
    Call AddBtn(s, "Btn_SearchClear",     ar & "条件クリア", "I4", COLOR_BTN_SUB, "条件", "", "")

    Set BuildSearchSpec = s
End Function

' ================================================================
' 関数名: BuildKnowledgeViewSpec (M-09)
' ================================================================
Private Function BuildKnowledgeViewSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-09"
    s.SheetName = SHEET_KNW_DISPLAY
    s.Title = "[v2] M-09 ナレッジ表示 (M-03 定義のフォーマットで保存済みナレッジを全文表示)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K19"

    Call AddSec(s, "A3", "[メタ情報]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A7", "[ナレッジ内容 (M-03 定義通りの順序・行数で全文表示)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A18", "[アクション]", COLOR_SECTION_BLUE2)

    Call AddStandardFieldRows(s, 8)

    Dim ar As String: ar = ChrW(&H25B6)
    Dim al As String: al = ChrW(&H2190)
    Call AddBtn(s, "Btn_BackToSearch",  al & " 検索に戻る",  "B19", COLOR_HEADER_LIGHT, "ナビ", "", "")
    Call AddBtn(s, "Btn_GoToEdit",      ar & " 修正に遷移", "D19", COLOR_BTN_PRIMARY, "ナビ", "", "")

    Set BuildKnowledgeViewSpec = s
End Function

' ================================================================
' 関数名: BuildStorageSpec (M-10)
' ================================================================
Private Function BuildStorageSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-10"
    s.SheetName = SHEET_STORAGE
    s.Title = "[v2] M-10 格納先設定 (行追加削除, 接続テスト)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K16"

    Call AddSec(s, "A2", "[設定済み格納先一覧]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A7", "[行追加]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A9", "[接続テスト結果]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A15", "[アクション]", COLOR_SECTION_BLUE2)

    s.HeaderRowAddr = "A3"
    s.HeaderLabels = Array("#", "格納先名", "種別", "パス/URL", "デフォルト", "有効", "テスト", "編集", "削除", "↑", "↓")
    s.EmptyStateAddr = "A4"
    s.EmptyStateText = "(格納先未設定 ― 0 件 / 「+ 新規格納先を追加」で行追加)"

    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_AddStorage",     "+ 新規格納先を追加", "B8", COLOR_BTN_NAV,     "追加", "", "")
    Call AddBtn(s, "Btn_TestAllStorage", ar & "全件テスト",     "B16", COLOR_BTN_NAV,     "アクション", "", "")
    Call AddBtn(s, "Btn_SaveStorage",    ar & "保存",           "D16", COLOR_BTN_PRIMARY, "アクション", "", "")
    Call AddBtn(s, "Btn_ResetStorage",   ar & "リセット",       "F16", COLOR_BTN_SUB,     "アクション", "", "")

    Set BuildStorageSpec = s
End Function

' ================================================================
' 関数名: BuildSystemSettingsSpec (M-11)
' ================================================================
Private Function BuildSystemSettingsSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-11"
    s.SheetName = SHEET_SETTINGS
    s.Title = "[v2] M-11 設定 (保存追加, リセット分離)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K17"

    Call AddSec(s, "A2", "[一般設定]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A7", "[ログ設定 (追記+ローテート方式)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A12", "[バックアップ設定]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A16", "[アクション (保存と破壊的操作を分離)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A18", "[破壊的操作 (押下時に二重確認ダイアログ)]", COLOR_DESTROY_BAR)

    ' 一般設定ラベル
    Call AddLabelField(s, 3, "既定フォーマット", "[選択]", False)
    Call AddLabelField(s, 4, "既定担当者",       "[文字列]", False)
    Call AddLabelField(s, 5, "プレビュー時のフォント", "[選択]", False)
    Call AddLabelField(s, 6, "プレビュー時のサイズ",   "[数値]", False)
    Call AddLabelField(s, 8, "ログ保持件数", "[数値]", False)
    Call AddLabelField(s, 9, "ログ保持日数", "[数値]", False)
    Call AddLabelField(s, 10, "ローテート単位", "[選択]", False)
    Call AddLabelField(s, 11, "保存先フォルダ", "[文字列]", False)
    Call AddLabelField(s, 13, "自動バックアップ", "[ON/OFF]", False)
    Call AddLabelField(s, 14, "保存先", "[文字列]", False)
    Call AddLabelField(s, 15, "世代保持", "[数値]", False)

    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_SaveSettings",     ar & "保存",     "B17", COLOR_BTN_PRIMARY, "アクション", "", "")
    Call AddBtn(s, "Btn_CancelSettings",   ar & "取消",     "D17", COLOR_BTN_SUB,     "アクション", "", "")
    Call AddBtn(s, "Btn_ResetToDefault",   "初期値にリセット",   "B19", COLOR_BTN_DANGER, "破壊", "", "")
    Call AddBtn(s, "Btn_ResetLog",         "全ログを削除",       "E19", COLOR_BTN_DANGER, "破壊", "", "")
    Call AddBtn(s, "Btn_DeleteAllBackup",  "バックアップを全削除", "H19", COLOR_BTN_DANGER, "破壊", "", "")

    Set BuildSystemSettingsSpec = s
End Function

' ================================================================
' 関数名: BuildMigrationSpec (M-12)
' ================================================================
Private Function BuildMigrationSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-12"
    s.SheetName = SHEET_MIGRATION
    s.Title = "[v2] M-12 フィールド反映 (バックアップ, 進捗, 部分失敗)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K20"

    Call AddSec(s, "A2", "[反映対象フォーマット]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A4", "[差分プレビュー (既存ナレッジへの影響)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A10", "[バックアップオプション]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A12", "[進捗表示]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A15", "[部分失敗一覧]", COLOR_DESTROY_BAR)
    Call AddSec(s, "A19", "[アクション]", COLOR_SECTION_BLUE2)

    s.HeaderRowAddr = "A5"
    s.HeaderLabels = Array("#", "フィールド名", "現状", "変更後", "対象件数", "操作", "メモ")
    s.EmptyStateAddr = "A6"
    s.EmptyStateText = "(差分なし ― 「差分確認」ボタンで再計算)"

    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_ConfirmDiff",     ar & "差分確認",   "I3", COLOR_BTN_PRIMARY, "確認", "", "")
    Call AddBtn(s, "Btn_MigrateFields",   ar & "反映実行",   "B20", COLOR_BTN_PRIMARY, "実行", "", "")
    Call AddBtn(s, "Btn_CancelMigrate",   ar & "中断",       "D20", COLOR_BTN_DANGER,  "実行", "", "")
    Call AddBtn(s, "Btn_RestoreBackup",   "バックアップから復旧", "F20", COLOR_BTN_SUB,  "実行", "", "")

    Set BuildMigrationSpec = s
End Function

' ================================================================
' 関数名: BuildFileFormatSpec (M-13)
' ================================================================
Private Function BuildFileFormatSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-13"
    s.SheetName = SHEET_FILE_FORMAT
    s.Title = "[v2] M-13 ファイル形式設定"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K21"

    Call AddSec(s, "A2", "[出力形式]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A8", "[Markdown 出力カスタマイズ]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A14", "[Word 出力カスタマイズ]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A20", "[アクション]", COLOR_SECTION_BLUE2)

    Call AddLabelField(s, 3, "既定の出力形式", "[選択 — Word .docx]", False)
    Call AddLabelField(s, 4, "代替形式 (チェック)", "[Markdown / HTML / PDF]", False)
    Call AddLabelField(s, 5, "ファイル名テンプレート", "[$KNo$_$YYYYMMDD$_$Title$.docx]", False)
    Call AddLabelField(s, 6, "改行コード", "[CRLF / LF]", False)
    Call AddLabelField(s, 7, "文字コード", "[UTF-8 (BOM 付き)]", False)
    Call AddLabelField(s, 9, "見出し記号", "[#  (ATX)]", False)
    Call AddLabelField(s, 10, "リスト記号", "[- (ハイフン)]", False)
    Call AddLabelField(s, 11, "表区切り行", "[| --- (3 列以上は自動拡張)]", False)
    Call AddLabelField(s, 12, "コードフェンス", "[``` (3 連バッククォート)]", False)
    Call AddLabelField(s, 13, "水平区切り", "[---]", False)
    Call AddLabelField(s, 15, "ベーステンプレート", "[C:\KnowledgeMgr\templates\base.dotx]", False)
    Call AddLabelField(s, 16, "既定フォント", "[Yu Gothic UI]", False)
    Call AddLabelField(s, 17, "既定サイズ", "[10.5pt]", False)

    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_PreviewFileFormat", ar & "プレビュー", "B21", COLOR_BTN_NAV,     "アクション", "", "")
    Call AddBtn(s, "Btn_SaveFileFormat",    ar & "保存",       "D21", COLOR_BTN_PRIMARY, "アクション", "", "")
    Call AddBtn(s, "Btn_CancelFileFormat",  ar & "取消",       "F21", COLOR_BTN_SUB,     "アクション", "", "")

    Set BuildFileFormatSpec = s
End Function

' ================================================================
' 関数名: BuildLogSpec (M-14)
' ================================================================
Private Function BuildLogSpec() As clsScreenSpec
    Dim s As clsScreenSpec
    Set s = New clsScreenSpec
    s.ScreenId = "M-14"
    s.SheetName = SHEET_LOG
    s.Title = "[v2] M-14 操作ログ (追記+ローテート, エクスポート/フィルタ)"
    s.TitleColorHex = COLOR_TITLE_BLUE
    s.BackToMainAddr = "K20"

    Call AddSec(s, "A2", "[ログ管理方式]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A4", "[フィルタ]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A7", "[ログ表示 (最新)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A14", "[ファイル一覧 (ローテート済み)]", COLOR_SECTION_BLUE2)
    Call AddSec(s, "A19", "[アクション]", COLOR_SECTION_BLUE2)

    s.HeaderRowAddr = "A8"
    s.HeaderLabels = Array("#", "日時", "レベル", "ユーザ", "メッセージ", "対象")
    s.EmptyStateAddr = "A9"
    s.EmptyStateText = "(ログ未蓄積 ― 操作後に絞込ボタンを押下)"

    Dim ar As String: ar = ChrW(&H25B6)
    Call AddBtn(s, "Btn_FilterLog",     ar & "絞込",       "G5", COLOR_BTN_PRIMARY, "フィルタ", "", "")
    Call AddBtn(s, "Btn_ClearLogFilter", ar & "クリア",     "I5", COLOR_BTN_SUB,     "フィルタ", "", "")
    Call AddBtn(s, "Btn_ExportLogCSV",  ar & "CSV出力",    "K5", COLOR_BTN_NAV,     "フィルタ", "", "")
    Call AddBtn(s, "Btn_ReloadLog",     ar & "ログを再読込", "B20", COLOR_BTN_NAV,     "アクション", "", "")

    Set BuildLogSpec = s
End Function

' ================================================================
' --- 内部ヘルパー（spec 構築の繰り返しを抑制） ---
' ================================================================

Private Sub AddSec(ByVal s As clsScreenSpec, ByVal addr As String, _
                    ByVal label As String, ByVal colorHex As String)
    Dim sec As clsSectionSpec
    Set sec = New clsSectionSpec
    sec.Init addr, label, colorHex
    s.AddSection sec
End Sub

Private Sub AddBtn(ByVal s As clsScreenSpec, _
                    ByVal btnName As String, ByVal caption As String, _
                    ByVal cellAddr As String, ByVal colorHex As String, _
                    ByVal groupName As String, _
                    ByVal hintAddr As String, ByVal hintText As String)
    Dim btn As clsButtonSpec
    Set btn = New clsButtonSpec
    btn.Init btnName, caption, cellAddr, colorHex, groupName, hintAddr, hintText
    s.AddButton btn
End Sub

' ================================================================
' 関数名: AddStandardFieldRows
' 概要:   標準ナレッジフィールド（件名/発生日時/担当者/カテゴリ/優先度/事象/原因/対処内容）
'         を指定行から順に追加。データ無くてもラベルが見える「空状態 UI」を実現。
' 引数:   s         - clsScreenSpec
'         startRow  - フィールド開始行
' ================================================================
Private Sub AddStandardFieldRows(ByVal s As clsScreenSpec, ByVal startRow As Long)
    Dim names As Variant
    Dim types As Variant
    Dim reqs As Variant
    Dim rows As Variant
    Dim hints As Variant
    names = Array("件名", "発生日時", "担当者", "カテゴリ", "優先度", "事象", "原因", "対処内容")
    types = Array("単一行 1行", "日時 1行", "単一行 1行", "選択 1行", "選択 1行", "複数行 5行", "複数行 3行", "複数行 5行")
    reqs  = Array(True,    True,    True,    True,    True,    True,    True,    True)
    rows  = Array(1,       1,       1,       1,       1,       5,       3,       5)
    hints = Array("(単一行入力)", "(現在日時を既定で表示)", "(単一行入力)", "(選択してください)", "(選択 ― 高/中/低)", _
                   "(複数行入力 / Alt+Enter で改行)", "(複数行入力 / Alt+Enter で改行)", "(複数行入力 / Alt+Enter で改行)")

    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim r As Long
        r = startRow + i
        Dim fld As clsFieldSpec
        Set fld = New clsFieldSpec
        fld.Init i + 1, CStr(names(i)), CStr(types(i)), CBool(reqs(i)), CLng(rows(i)), CStr(hints(i))
        fld.SetCellAddrs "A" & r, "B" & r, "C" & r, "D" & r, "E" & r
        s.AddField fld
    Next i
End Sub

' ================================================================
' 関数名: AddLabelField
' 概要:   1 行分の単純な「項目名 [型]」ラベルを spec に追加
' ================================================================
Private Sub AddLabelField(ByVal s As clsScreenSpec, ByVal row As Long, _
                           ByVal label As String, ByVal typeText As String, _
                           ByVal required As Boolean)
    Dim fld As clsFieldSpec
    Set fld = New clsFieldSpec
    fld.Init 0, label, typeText, required, 1, ""
    fld.SetCellAddrs "", "B" & row, "C" & row, "D" & row, "E" & row
    s.AddField fld
End Sub
```
