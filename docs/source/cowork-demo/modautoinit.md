---
title: modAutoInit.bas
---

# modAutoInit.bas

| 項目 | 値 |
|---|---|
| 層 | Cowork デモ専用 |
| 種別 | 標準モジュール (.bas) |
| 役割 | Cowork デモ用 auto-init (1 ティック遅延でセットアップ + シード自動実行) |
| 行数 | 135 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modAutoInit` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modAutoInit"
Option Explicit

' ================================================================
' モジュール: modAutoInit (Cowork デモ専用 / 配布版には含まれない)
' 概要:       Workbook_Open から Application.OnTime 経由で 1 ティック遅延
'             起動される auto-init エントリポイント。
'             ThisWorkbook を「OnTime を呼ぶだけ」にすることで、
'             本番 release_v2 の ThisWorkbook (auto-init 無し) との
'             差分を最小化し、入れ替えを容易にする (D-1 部分対応)。
' 依存先:     modSetup (SetupSheetsAndButtons), modDemoSeeder (Macro_SeedDemoData),
'             modCommon (SHEET_* 定数)
' 備考:       OnTime 1 ティック遅延の理由 (C-1):
'             Workbook_Open の context では Application.Run / 各種シート操作が
'             不安定になりうる (UIPI / 修復モード復帰直後の制約)。
'             OnTime で 1 ティック後に呼び直すことで、Excel が完全に
'             init を終えた状態で auto-init を実行できる。
' ================================================================

' --- メインシート上のセットアップ完了判定用ボタン名 ---
'     PlaceButtonsOnMainSheet で配置される代表ボタン
Private Const BTN_REPRESENTATIVE_MAIN As String = "Btn_TaskSetup"

' --- ナレッジ一覧シート デモデータ存在判定セル ---
Private Const KL_PROBE_ROW As Long = 3
Private Const KL_PROBE_COL As Long = 1

' ================================================================
' 関数名: RunAutoInitDeferred
' 概要:   Workbook_Open から Application.OnTime で 1 ティック遅延起動される
'         auto-init エントリポイント。冪等性確保。
' 引数:   なし
' 戻り値: なし
' 備考:   1) シート未生成 or ボタン未配置 → SetupSheetsAndButtons silent:=True
'         2) ナレッジ一覧未投入 → Macro_SeedDemoData silent:=True
'         3) メインシートをアクティブ化 (失敗は無視)
'         エラーは MsgBox 表示で握り潰さない。
' ================================================================
Public Sub RunAutoInitDeferred()
    On Error GoTo ErrHandler

    ' --- (1) シート存在 + ボタン配置 の両方を満たすか判定 ---
    If (Not SheetExists(SHEET_SEARCH)) Or (Not AllButtonsPlaced()) Then
        SetupSheetsAndButtons silent:=True
    End If

    ' --- (2) デモデータ判定 ---
    If Not HasDemoData() Then
        Macro_SeedDemoData silent:=True
    End If

    ' --- (3) メインシートをアクティブに (失敗時は無視、起動を止めない) ---
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_MAIN).Activate
    On Error GoTo ErrHandler
    Exit Sub

ErrHandler:
    ' v5: async 実行 (Application.OnTime) 経由でこの Sub に来るため、
    '     Workbook_Open の ErrHandler ではなくここで詳細を捕捉する必要がある。
    '     Source / Number / Description を全部表示し、ユーザが報告できるようにする。
    Dim errMsg As String
    errMsg = "auto-init 失敗 (RunAutoInitDeferred):" & vbCrLf & _
            "Number: " & CStr(Err.Number) & vbCrLf & _
            "Source: " & Err.Source & vbCrLf & _
            "Description: " & Err.Description & vbCrLf & vbCrLf & _
            "[Alt+F8] → SetupSheetsAndButtons を手動実行してください。" & vbCrLf & _
            "解決しない場合は VBE (Alt+F11) → デバッグ → VBAProject のコンパイル " & _
            "で赤行のモジュール名/行番号/メッセージを確認してください。"
    MsgBox errMsg, vbExclamation, "auto-init エラー"
End Sub

' ================================================================
' 関数名: SheetExists
' 概要:   指定名のシートが ThisWorkbook 内に存在するか判定
' 引数:   sheetName - 確認対象のシート名
' 戻り値: Boolean - 存在すれば True
' ================================================================
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = (Err.Number = 0 And Not ws Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

' ================================================================
' 関数名: AllButtonsPlaced
' 概要:   メインシートに代表的なセットアップ完了印 (Btn_TaskSetup) が
'         配置されているかで「ボタン配置完了」を判定する。
' 引数:   なし
' 戻り値: Boolean - 配置済なら True
' 備考:   29 個全部を確認するのは過剰。代表 1 個で「セットアップマクロが
'         走り終えたか」のヒューリスティック判定に十分。
' ================================================================
Private Function AllButtonsPlaced() As Boolean
    On Error GoTo NotFound
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    If ws Is Nothing Then
        AllButtonsPlaced = False
        Exit Function
    End If

    Dim shp As Shape
    Set shp = ws.Shapes(BTN_REPRESENTATIVE_MAIN)
    AllButtonsPlaced = (Not shp Is Nothing)
    Exit Function

NotFound:
    AllButtonsPlaced = False
End Function

' ================================================================
' 関数名: HasDemoData
' 概要:   ナレッジ一覧シート row 3 col A にデモデータが投入済みか判定
' 引数:   なし
' 戻り値: Boolean - データがあれば True
' 備考:   ナレッジ一覧シート自体が無い場合は False (再 Seed の対象)
' ================================================================
Private Function HasDemoData() As Boolean
    On Error GoTo NotFound
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    If ws Is Nothing Then
        HasDemoData = False
        Exit Function
    End If
    HasDemoData = (CStr(ws.Cells(KL_PROBE_ROW, KL_PROBE_COL).Value) <> "")
    Exit Function

NotFound:
    HasDemoData = False
End Function
```

## 関連

- 呼び出す: `modSetup`, `modDemoSeeder`, `modCommon`
- - 呼び出される: `ThisWorkbook (デモ版)`
  - 
