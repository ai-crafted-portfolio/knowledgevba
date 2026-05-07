---
title: modCommon.bas
---

# modCommon.bas

| 項目 | 値 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 全モジュール共通定数 (シート名 / 列番号 / 行番号など) |
| 行数 | 101 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modCommon` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modCommon"
Option Explicit

' ================================================================
' モジュール: modCommon（ユーティリティ層）
' 概要:       全モジュールで共有する定数を定義する
' 依存先:     なし
' ================================================================

' --- シート名定数 ---
Public Const SHEET_MAIN As String = "メイン"
Public Const SHEET_FORMAT_LIST As String = "フォーマット一覧"
Public Const SHEET_FORMAT_DESIGN As String = "フォーマット設計"
Public Const SHEET_FORMAT_PREVIEW As String = "フォーマットプレビュー"
Public Const SHEET_KNW_SAVE As String = "ナレッジ登録"
Public Const SHEET_KNW_EDIT As String = "ナレッジ修正"
Public Const SHEET_KNW_LIST As String = "ナレッジ一覧"
Public Const SHEET_SEARCH As String = "検索"
Public Const SHEET_KNW_DISPLAY As String = "ナレッジ表示"
Public Const SHEET_STORAGE As String = "格納先設定"
Public Const SHEET_SETTINGS As String = "設定"
Public Const SHEET_MIGRATION As String = "既存データへのフィールド反映"
Public Const SHEET_FILE_FORMAT As String = "データファイル形式"
Public Const SHEET_LOG As String = "ログ"

' --- ファイル形式定数 ---
Public Const CHARSET_SJIS As String = "Shift_JIS"
Public Const STANZA_SEP As String = "==="

' --- デバッグレベル定数 ---
Public Const DEBUG_OFF As String = "OFF"
Public Const DEBUG_ON As String = "ON"

' --- フィールド型定数（仕様書で確定した6種） ---
Public Const FIELD_TYPE_STRING As String = "文字列"
Public Const FIELD_TYPE_LONGTEXT As String = "長文テキスト"
Public Const FIELD_TYPE_NUMBER As String = "数値"
Public Const FIELD_TYPE_DATE As String = "日付"
Public Const FIELD_TYPE_TIME As String = "時刻"
Public Const FIELD_TYPE_FILEREF As String = "ファイル参照"

' --- ログレベル定数 ---
Public Const LOG_LEVEL_ERROR As String = "エラー"
Public Const LOG_LEVEL_WARN As String = "警告"
Public Const LOG_LEVEL_INFO As String = "情報"
Public Const LOG_LEVEL_DEBUG As String = "デバッグ"

' --- ログシート列番号 ---
Public Const LOG_COL_TIMESTAMP As Long = 1
Public Const LOG_COL_MODULE As Long = 2
Public Const LOG_COL_FUNCTION As Long = 3
Public Const LOG_COL_LEVEL As Long = 4
Public Const LOG_COL_MESSAGE As Long = 5

' --- 設定シート行番号 ---
Public Const SETTINGS_ROW_DATAFOLDER As Long = 3
Public Const SETTINGS_ROW_CHARSET As Long = 4
Public Const SETTINGS_ROW_DEBUGLEVEL As Long = 5

' --- 設定シート列番号 ---
Public Const SETTINGS_COL_NAME As Long = 2
Public Const SETTINGS_COL_VALUE As Long = 3

' --- タスク名定数（8タスク） ---
Public Const TASK_SETUP As String = "初回セットアップ"
Public Const TASK_CONFIG As String = "設定変更"
Public Const TASK_FORMAT As String = "フォーマット管理"
Public Const TASK_REGISTER As String = "ナレッジ登録"
Public Const TASK_SEARCH As String = "検索・確認"
Public Const TASK_EDIT As String = "ナレッジ修正"
Public Const TASK_DELETE As String = "ナレッジ削除"
Public Const TASK_MIGRATE As String = "既存データ反映"

' --- テストモード関連定数 ---
Public Const SETTINGS_ROW_TESTMODE As Long = 6
Public Const TESTMODE_ON As String = "TRUE"
Public Const TESTMODE_OFF As String = "FALSE"

' --- テスト結果判定定数 ---
Public Const TEST_RESULT_PASS As String = "PASS"
Public Const TEST_RESULT_FAIL As String = "FAIL"
Public Const TEST_RESULT_SKIP As String = "SKIP"


' ================================================================
' 関数名: SheetExists
' 概要:   指定名のワークシートが ThisWorkbook 内に存在するか判定する共通ヘルパー。
'         M-2: modSetup.IsSheetExists と modAutoInit.SheetExists の重複定義を
'         本関数に集約し、新規モジュールはこちらを使うよう統一する。
'         既存の Private 版は内部実装維持 (互換性のため)。
' 引数:   sheetName - 確認対象のシート名
' 戻り値: Boolean - 存在すれば True
' ================================================================
Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = (Err.Number = 0 And Not ws Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function
```

## 関連

- 呼び出す: なし
- 呼び出される: `全モジュール`
