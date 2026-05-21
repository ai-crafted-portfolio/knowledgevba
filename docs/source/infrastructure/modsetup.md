---
title: modSetup.bas
---

# modSetup.bas

| 項目 | 内容 |
|---|---|
| 層 | インストーラ層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | 3 ブック共通のセットアップ入口。各ブックに必要なシート（LOG・各画面）を確認し、無ければ作成する |
| 行数 | 208 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modSetup.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modSetup"
' ============================================================
' modSetup (Installer layer, v2.1 / 3-xlsm)
' Role:
'   3 xlsm (管理 / 検索 / 登録修正) 共通の installer entry。
'   各 xlsm に対応する Setup_* Public Sub を提供し、
'   - 必要シート (LOG / 画面別 M-*) の存在確認 + 無ければ作成
'   - SystemSettings 系シートの初期値投入 (管理.xlsm のみ)
'   - 成功/失敗時に logId emit (SAVE-EXIT-OK-II-* / BACKTOMAIN-ERR-EE-*)
'   を実施する。
' Design SSOT:
'   - vba_module_placement_v2.md  1.1 (modSetup = 全 xlsm 配置、3 xlsm 共 setup)
'   - installer_3xlsm_v2_メイン.md  1.4 / 1.5 (Step 5-7 の xlsm 別配置)
'   - architecture_v2_3xlsm_stanza.md  7
' Logging:
'   - LogID 規約 = logging_naming_cheatsheet_v1.md  3 (II / EE 末尾可)
'   - 既存 max SEQ + 1 を採用 (audit script で確認済)
' ASCII-only inside VBA string literals (CP932 mojibake avoidance).
' ============================================================
Option Explicit

Private Const MOD_NAME As String = "modSetup"

' 3 xlsm 別の必要シート一覧 (v2.1, vba_module_placement_v2.md  1.5 画面層と整合)
' 注: メイン画面 M-01 は v2.1 で廃止 (Q39)、LOG は全 xlsm 共通必須。
Private Const SHEETS_KANRI    As String = "LOG,M-02,M-03,M-04,M-12,M-13,M-10,M-11,M-14"
Private Const SHEETS_KENSAKU  As String = "LOG,M-07,M-08,M-09"
Private Const SHEETS_TOUROKU  As String = "LOG,M-05,M-06"

' 設定系 2 画面シート (管理.xlsm 専用)。v2.1 SSOT-Q22 / R10-03: 設定編集を 2 画面に分割。
'   M-11 設定 = debugLevel を編集 / M-10 格納先設定 = パス系 4 dir を編集。
' v2.1 では config.txt が SSOT のため、本シートは初期値投入のみ (NM-2 命名統一 sheet 形)
Private Const SHEET_SYSTEM_SETTINGS As String = "M-11"
Private Const SHEET_STORAGE As String = "M-10"

' ============================================================
' Public Sub: Setup_管理
' Role: 管理.xlsm の install entry。シート構築 + SystemSettings 初期値投入。
' ============================================================
Public Sub Setup_管理()
    SetupCore "管理", SHEETS_KANRI, True
End Sub

' ============================================================
' Public Sub: Setup_検索
' Role: 検索.xlsm の install entry。シート構築のみ。
' ============================================================
Public Sub Setup_検索()
    SetupCore "検索", SHEETS_KENSAKU, False
End Sub

' ============================================================
' Public Sub: Setup_登録修正
' Role: 登録修正.xlsm の install entry。シート構築のみ。
' ============================================================
Public Sub Setup_登録修正()
    SetupCore "登録修正", SHEETS_TOUROKU, False
End Sub

' ============================================================
' Private Sub: SetupCore
' Role: 3 xlsm 共通 setup ロジック。
'       1) 必要シートを CSV から確保 (Worksheets("name") 不在なら Worksheets.Add)
'       2) 管理.xlsm のみ SystemSettings 初期値を投入
'       3) 結果を logId 付きで emit (SAVE-EXIT-OK-II-009 / BACKTOMAIN-ERR-EE-033)
' Args:
'   xlsmName       - "管理" | "検索" | "登録修正" (log emit 用)
'   sheetsCsv      - 確保すべきシート名の CSV
'   isKanriXlsm    - True の場合のみ SystemSettings 初期値投入
' ============================================================
Private Sub SetupCore(ByVal xlsmName As String, ByVal sheetsCsv As String, ByVal isKanriXlsm As Boolean)
    On Error GoTo ErrHandler
    Dim createdCount As Long
    createdCount = EnsureSheets(sheetsCsv)
    If isKanriXlsm Then
        SeedSystemSettings
        SeedStorage
    End If
    ' SAVE-EXIT-OK-II-009: setup 成功 (audit max=008 +1)
    On Error Resume Next
    Dim oLogger009 As clsLogger
    Set oLogger009 = New clsLogger
    oLogger009.Init ThisWorkbook.Worksheets("LOG")
    oLogger009.LogInfo MOD_NAME, "SetupCore", _
        "Setup " & xlsmName & " ok, created=" & createdCount, _
        xlsmName, "SAVE-EXIT-OK-II-009"
    On Error GoTo 0
    Exit Sub
ErrHandler:
    ' BACKTOMAIN-ERR-EE-033: setup 失敗 (audit max=032 +1)
    On Error Resume Next
    Dim oLogger033 As clsLogger
    Set oLogger033 = New clsLogger
    oLogger033.Init ThisWorkbook.Worksheets("LOG")
    oLogger033.LogError MOD_NAME, "SetupCore", _
        "Setup " & xlsmName & " failed: " & Err.Number & " " & Err.Description, _
        xlsmName, "BACKTOMAIN-ERR-EE-033"
    On Error GoTo 0
    Debug.Print "[ERR] modSetup.SetupCore(" & xlsmName & "): " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Private Function: EnsureSheets
' Role: CSV のシート名を順に確保。不在なら末尾に Add。
' Return: 新規追加件数
' ============================================================
Private Function EnsureSheets(ByVal sheetsCsv As String) As Long
    Dim parts() As String
    parts = Split(sheetsCsv, ",")
    Dim added As Long
    added = 0
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim nm As String
        nm = Trim(parts(i))
        If Len(nm) > 0 Then
            If Not SheetExists(nm) Then
                Dim ws As Worksheet
                Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
                ws.Name = nm
                added = added + 1
            End If
        End If
    Next i
    EnsureSheets = added
End Function

' ============================================================
' Private Function: SheetExists
' Role: シート名の存在確認 (大文字小文字 / 全角半角は Excel 既定挙動に従う)
' ============================================================
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = (Err.Number = 0 And Not ws Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

' ============================================================
' Private Sub: SeedSystemSettings
' Role: 管理.xlsm の M-11 設定シートに初期値を投入 (debugLevel のみ)。
'       既存セル値があれば上書きしない (= install を idempotent)。
'       SSOT-Q22 / R10-03: 設定編集の 2 画面分割により、M-11 は debugLevel のみ。
'       パス系 4 dir は M-10 格納先設定シート (SeedStorage) が担当する。
' Cell layout (M-11, modEntrySettings.ADDR_DEBUG_LEVEL と整合):
'   B6 = debugLevel  / C6 = (value, str)
' ============================================================
Private Sub SeedSystemSettings()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SYSTEM_SETTINGS)
    If ws Is Nothing Then Exit Sub
    SeedLabel ws, "B6", "debugLevel"
    SeedValueIfEmpty ws, "C6", SafeGetConfig("debugLevel", "INFO")
    On Error GoTo 0
End Sub

' ============================================================
' Private Sub: SeedStorage
' Role: 管理.xlsm の M-10 格納先設定シートにパス系 4 dir の初期値を投入。
'       既存セル値があれば上書きしない (= install を idempotent)。
'       SSOT-Q22 / R10-03: 設定編集の 2 画面分割により M-10 はパス系 4 dir を担当。
' Cell layout (M-10, modEntrySettings.ADDR_DATA_DIR..ADDR_BACKUP_DIR と整合):
'   B2 = data_dir    / C2 = (value)
'   B3 = format_dir  / C3 = (value)
'   B4 = ui_dir      / C4 = (value)
'   B5 = backup_dir  / C5 = (value)
' ============================================================
Private Sub SeedStorage()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STORAGE)
    If ws Is Nothing Then Exit Sub
    SeedLabel ws, "B2", "data_dir"
    SeedLabel ws, "B3", "format_dir"
    SeedLabel ws, "B4", "ui_dir"
    SeedLabel ws, "B5", "backup_dir"
    SeedValueIfEmpty ws, "C2", SafeGetConfig("data_dir", "C:\KnowledgeMgr\data\")
    SeedValueIfEmpty ws, "C3", SafeGetConfig("format_dir", "C:\KnowledgeMgr\formats\")
    SeedValueIfEmpty ws, "C4", SafeGetConfig("ui_dir", "C:\KnowledgeMgr\ui\")
    SeedValueIfEmpty ws, "C5", SafeGetConfig("backup_dir", "C:\KnowledgeMgr\data\backup\")
    On Error GoTo 0
End Sub

Private Sub SeedLabel(ByVal ws As Worksheet, ByVal addr As String, ByVal label As String)
    On Error Resume Next
    ws.Range(addr).Value = label
    On Error GoTo 0
End Sub

Private Sub SeedValueIfEmpty(ByVal ws As Worksheet, ByVal addr As String, ByVal v As String)
    On Error Resume Next
    Dim cur As String
    cur = CStr(ws.Range(addr).Value)
    If Len(cur) = 0 Then ws.Range(addr).Value = v
    On Error GoTo 0
End Sub

Private Function SafeGetConfig(ByVal key As String, ByVal defaultValue As String) As String
    On Error Resume Next
    Dim v As String
    v = modConfigHolder.GetValueOrDefault(key, defaultValue)
    If Len(v) = 0 Then v = defaultValue
    SafeGetConfig = v
    On Error GoTo 0
End Function
```
