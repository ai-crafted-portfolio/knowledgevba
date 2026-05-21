---
title: modEntrySettings.bas
---

# modEntrySettings.bas

| 項目 | 内容 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 管理.xlsm |
| 役割 | 格納先設定（M-10）・設定（M-11）画面の入口。設定値のシート反映・再読込・破棄 |
| 行数 | 384 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modEntrySettings.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modEntrySettings"
' ============================================================
' modEntrySettings (Entry layer, 管理.xlsm only / v2.1)
' Role:
'   管理.xlsm の設定系 2 画面に対する entry point Sub 群。
'   設計 SSOT (architecture §6 / §7.2 #47-48 / screen_design_v2 §2.9 /
'   placement_v2 §6) が規定する 2 画面分割 (SSOT-Q22 確定):
'     M-10 格納先設定 (clsStorageScreen)        = config.txt の
'         パス系 4 dir (data_dir / format_dir / ui_dir / backup_dir) を編集
'     M-11 設定       (clsSystemSettingsScreen)   = config.txt の
'         debugLevel を編集
'   本 module は M-10 / M-11 両画面の entry (placement_v2 §6)。
'   modConfigHolder (in-memory) の値をシートに反映 (Open) / シート編集値を
'   再ロード (Apply) / 編集破棄 (Discard) する。
'   config.txt 本体への永続化は modConfigLoader 経由の責務 (本 module は
'   memory holder に対する 反映 only、external I/O 直接呼び出しは禁止
'   = ADR-0053 §2.9 I/O 独立原則)。
'   modConfigHolder.SetConfig は dict 全置換 (merge しない) のため、
'   各 Apply Sub は自画面が所有しない key を holder から読みそろえて
'   5 key フルの dict を組み直してから SetConfig する (merge-on-write)。
' Naming (設計 SSOT 準拠):
'     M-11 設定 (debugLevel):
'       Btn_OpenSettings_v21    -- 設定画面オープン (config holder -> M-11 sheet)
'       Btn_SaveSettings_v21    -- 編集適用 (M-11 sheet -> config holder, validation 含む)
'       Btn_CancelSettings_v21  -- 編集破棄 (config holder -> M-11 sheet, idempotent)
'     M-10 格納先設定 (パス系 4 dir):
'       Btn_OpenStorage_v21     -- 格納先設定画面オープン (config holder -> M-10 sheet)
'       Btn_SaveStorage_v21     -- 編集適用 (M-10 sheet -> config holder, validation 含む)
'       Btn_CancelStorage_v21   -- 編集破棄 (config holder -> M-10 sheet, idempotent)
' Logging:
'   - SAVE-EXIT-OK-II-010 / 011 / 012 (M-11 Open / Apply / Discard 正常完了)
'   - SAVE-EXIT-OK-II-013 / 014 / 015 (M-10 Open / Apply / Discard 正常完了)
'   - VALIDATE-WARN-WW-033 (M-11 Apply の validation 失敗)
'   - VALIDATE-WARN-WW-035 (M-10 Apply の validation 失敗)
'   - BACKTOMAIN-ERR-EE-034 (M-11 Apply 中例外)
'   - BACKTOMAIN-ERR-EE-036 (M-10 Apply 中例外)
' ASCII-only inside VBA string literals (CP932 mojibake avoidance).
' ============================================================
Option Explicit

Private Const MOD_NAME As String = "modEntrySettings"

' M-11 設定 (debugLevel) 画面
Private Const SHEET_SETTINGS As String = "M-11"
' M-10 格納先設定 (パス系 4 dir) 画面
Private Const SHEET_STORAGE As String = "M-10"

' --- M-11 cell layout (debugLevel のみ) ---
Private Const ADDR_DEBUG_LEVEL As String = "C6"

' --- M-10 cell layout (パス系 4 dir) ---
' セル配置は modSetup.SeedStorage と整合。C2-C5 に 4 dir。
Private Const ADDR_DATA_DIR    As String = "C2"
Private Const ADDR_FORMAT_DIR  As String = "C3"
Private Const ADDR_UI_DIR      As String = "C4"
Private Const ADDR_BACKUP_DIR  As String = "C5"

' --- config.txt key default 値 (holder 未初期化時のフォールバック) ---
Private Const DEF_DATA_DIR    As String = "C:\KnowledgeMgr\data\"
Private Const DEF_FORMAT_DIR  As String = "C:\KnowledgeMgr\formats\"
Private Const DEF_UI_DIR      As String = "C:\KnowledgeMgr\ui\"
Private Const DEF_BACKUP_DIR  As String = "C:\KnowledgeMgr\data\backup\"
Private Const DEF_DEBUG_LEVEL As String = "INFO"

' ============================================================
' ===== M-11 設定 (debugLevel) entry =================================
' ============================================================

' ============================================================
' Public Sub: Btn_OpenSettings_v21
' Role: M-11 設定画面を表示し、modConfigHolder の debugLevel をセルに反映。
'       Workbook 起動時に config holder が SetConfig 済みであることを前提。
' ============================================================
Public Sub Btn_OpenSettings_v21()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    ws.Activate
    LoadSettingsToSheet ws
    ' SAVE-EXIT-OK-II-010
    On Error Resume Next
    Dim oLogger010 As clsLogger
    Set oLogger010 = New clsLogger
    oLogger010.Init ThisWorkbook.Worksheets("LOG")
    oLogger010.LogInfo MOD_NAME, "Btn_OpenSettings_v21", _
        "Settings opened", "", "SAVE-EXIT-OK-II-010"
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_OpenSettings_v21: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_SaveSettings_v21
' Role: M-11 シートの debugLevel を validation してから modConfigHolder に反映。
'       パス系 4 dir は M-10 が所有するため holder から読みそろえて保存。
'       本 Sub は memory holder のみ更新。config.txt 永続化は modConfigLoader 経由。
' ============================================================
Public Sub Btn_SaveSettings_v21()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)

    Dim ngLabel As String
    Dim ngAddr As String
    If Not ValidateSettingsSheet(ws, ngLabel, ngAddr) Then
        ' VALIDATE-WARN-WW-033
        On Error Resume Next
        Dim oLogger033 As clsLogger
        Set oLogger033 = New clsLogger
        oLogger033.Init ThisWorkbook.Worksheets("LOG")
        oLogger033.LogWarn MOD_NAME, "Btn_SaveSettings_v21", _
            "Validation failed at " & ngAddr & " (" & ngLabel & ")", _
            ngLabel, "VALIDATE-WARN-WW-033"
        On Error GoTo 0
        Exit Sub
    End If

    ApplySettingsToHolder ws

    ' SAVE-EXIT-OK-II-011
    On Error Resume Next
    Dim oLogger011 As clsLogger
    Set oLogger011 = New clsLogger
    oLogger011.Init ThisWorkbook.Worksheets("LOG")
    oLogger011.LogInfo MOD_NAME, "Btn_SaveSettings_v21", _
        "Settings applied to holder", "", "SAVE-EXIT-OK-II-011"
    On Error GoTo 0
    Exit Sub
ErrHandler:
    ' BACKTOMAIN-ERR-EE-034
    On Error Resume Next
    Dim oLogger034 As clsLogger
    Set oLogger034 = New clsLogger
    oLogger034.Init ThisWorkbook.Worksheets("LOG")
    oLogger034.LogError MOD_NAME, "Btn_SaveSettings_v21", _
        "Apply failed: " & Err.Number & " " & Err.Description, _
        "", "BACKTOMAIN-ERR-EE-034"
    On Error GoTo 0
    Debug.Print "[ERR] Btn_SaveSettings_v21: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_CancelSettings_v21
' Role: M-11 編集破棄。modConfigHolder の debugLevel で M-11 sheet を上書き (= Open と同等)。
' ============================================================
Public Sub Btn_CancelSettings_v21()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    LoadSettingsToSheet ws
    ' SAVE-EXIT-OK-II-012
    On Error Resume Next
    Dim oLogger012 As clsLogger
    Set oLogger012 = New clsLogger
    oLogger012.Init ThisWorkbook.Worksheets("LOG")
    oLogger012.LogInfo MOD_NAME, "Btn_CancelSettings_v21", _
        "Settings discarded (reloaded from holder)", "", "SAVE-EXIT-OK-II-012"
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_CancelSettings_v21: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' ===== M-10 格納先設定 (パス系 4 dir) entry =====================
' ============================================================

' ============================================================
' Public Sub: Btn_OpenStorage_v21
' Role: M-10 格納先設定画面を表示し、modConfigHolder のパス系 4 dir をセルに反映。
' ============================================================
Public Sub Btn_OpenStorage_v21()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STORAGE)
    ws.Activate
    LoadStorageToSheet ws
    ' SAVE-EXIT-OK-II-013
    On Error Resume Next
    Dim oLogger013 As clsLogger
    Set oLogger013 = New clsLogger
    oLogger013.Init ThisWorkbook.Worksheets("LOG")
    oLogger013.LogInfo MOD_NAME, "Btn_OpenStorage_v21", _
        "Storage settings opened", "", "SAVE-EXIT-OK-II-013"
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_OpenStorage_v21: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_SaveStorage_v21
' Role: M-10 シートのパス系 4 dir を validation してから modConfigHolder に反映。
'       debugLevel は M-11 が所有するため holder から読みそろえて保存。
' ============================================================
Public Sub Btn_SaveStorage_v21()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STORAGE)

    Dim ngLabel As String
    Dim ngAddr As String
    If Not ValidateStorageSheet(ws, ngLabel, ngAddr) Then
        ' VALIDATE-WARN-WW-035
        On Error Resume Next
        Dim oLogger035 As clsLogger
        Set oLogger035 = New clsLogger
        oLogger035.Init ThisWorkbook.Worksheets("LOG")
        oLogger035.LogWarn MOD_NAME, "Btn_SaveStorage_v21", _
            "Validation failed at " & ngAddr & " (" & ngLabel & ")", _
            ngLabel, "VALIDATE-WARN-WW-035"
        On Error GoTo 0
        Exit Sub
    End If

    ApplyStorageToHolder ws

    ' SAVE-EXIT-OK-II-014
    On Error Resume Next
    Dim oLogger014 As clsLogger
    Set oLogger014 = New clsLogger
    oLogger014.Init ThisWorkbook.Worksheets("LOG")
    oLogger014.LogInfo MOD_NAME, "Btn_SaveStorage_v21", _
        "Storage settings applied to holder", "", "SAVE-EXIT-OK-II-014"
    On Error GoTo 0
    Exit Sub
ErrHandler:
    ' BACKTOMAIN-ERR-EE-036
    On Error Resume Next
    Dim oLogger036 As clsLogger
    Set oLogger036 = New clsLogger
    oLogger036.Init ThisWorkbook.Worksheets("LOG")
    oLogger036.LogError MOD_NAME, "Btn_SaveStorage_v21", _
        "Apply failed: " & Err.Number & " " & Err.Description, _
        "", "BACKTOMAIN-ERR-EE-036"
    On Error GoTo 0
    Debug.Print "[ERR] Btn_SaveStorage_v21: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_CancelStorage_v21
' Role: M-10 編集破棄。modConfigHolder のパス系 4 dir で M-10 sheet を上書き。
' ============================================================
Public Sub Btn_CancelStorage_v21()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STORAGE)
    LoadStorageToSheet ws
    ' SAVE-EXIT-OK-II-015
    On Error Resume Next
    Dim oLogger015 As clsLogger
    Set oLogger015 = New clsLogger
    oLogger015.Init ThisWorkbook.Worksheets("LOG")
    oLogger015.LogInfo MOD_NAME, "Btn_CancelStorage_v21", _
        "Storage settings discarded (reloaded from holder)", "", "SAVE-EXIT-OK-II-015"
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_CancelStorage_v21: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' ===== M-11 private helpers (debugLevel) ====================
' ============================================================

' ============================================================
' Private Sub: LoadSettingsToSheet
' Role: modConfigHolder の debugLevel を M-11 sheet に反映。
' ============================================================
Private Sub LoadSettingsToSheet(ByVal ws As Worksheet)
    ws.Range(ADDR_DEBUG_LEVEL).Value = SafeGetCfg("debugLevel", DEF_DEBUG_LEVEL)
End Sub

' ============================================================
' Private Sub: ApplySettingsToHolder
' Role: M-11 sheet の debugLevel を modConfigHolder に反映。
'       SetConfig は dict 全置換のため、パス系 4 dir は holder 現値を
'       読みそろえて 5 key フルの dict を組み直してから SetConfig (merge-on-write)。
' ============================================================
Private Sub ApplySettingsToHolder(ByVal ws As Worksheet)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    ' パス系 4 dir = M-10 所有。holder 現値をそのまま維持。
    d("data_dir") = SafeGetCfg("data_dir", DEF_DATA_DIR)
    d("format_dir") = SafeGetCfg("format_dir", DEF_FORMAT_DIR)
    d("ui_dir") = SafeGetCfg("ui_dir", DEF_UI_DIR)
    d("backup_dir") = SafeGetCfg("backup_dir", DEF_BACKUP_DIR)
    ' debugLevel = M-11 所有。sheet 編集値を反映。
    d("debugLevel") = CStr(ws.Range(ADDR_DEBUG_LEVEL).Value)
    On Error Resume Next
    modConfigHolder.SetConfig d
    On Error GoTo 0
End Sub

' ============================================================
' Private Function: ValidateSettingsSheet
' Role: debugLevel が空でないことを検査。
' Return: True = OK, False = NG (outLabel / outAddr に NG 内容を返す)
' ============================================================
Private Function ValidateSettingsSheet(ByVal ws As Worksheet, _
                                       ByRef outLabel As String, _
                                       ByRef outAddr As String) As Boolean
    If Len(CStr(ws.Range(ADDR_DEBUG_LEVEL).Value)) = 0 Then
        outAddr = ADDR_DEBUG_LEVEL : outLabel = "debugLevel" : ValidateSettingsSheet = False : Exit Function
    End If
    ValidateSettingsSheet = True
End Function

' ============================================================
' ===== M-10 private helpers (パス系 4 dir) ====================
' ============================================================

' ============================================================
' Private Sub: LoadStorageToSheet
' Role: modConfigHolder のパス系 4 dir を M-10 sheet に反映。
' ============================================================
Private Sub LoadStorageToSheet(ByVal ws As Worksheet)
    ws.Range(ADDR_DATA_DIR).Value = SafeGetCfg("data_dir", DEF_DATA_DIR)
    ws.Range(ADDR_FORMAT_DIR).Value = SafeGetCfg("format_dir", DEF_FORMAT_DIR)
    ws.Range(ADDR_UI_DIR).Value = SafeGetCfg("ui_dir", DEF_UI_DIR)
    ws.Range(ADDR_BACKUP_DIR).Value = SafeGetCfg("backup_dir", DEF_BACKUP_DIR)
End Sub

' ============================================================
' Private Sub: ApplyStorageToHolder
' Role: M-10 sheet のパス系 4 dir を modConfigHolder に反映。
'       SetConfig は dict 全置換のため、debugLevel は holder 現値を
'       読みそろえて 5 key フルの dict を組み直してから SetConfig (merge-on-write)。
' ============================================================
Private Sub ApplyStorageToHolder(ByVal ws As Worksheet)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    ' パス系 4 dir = M-10 所有。sheet 編集値を反映。
    d("data_dir") = CStr(ws.Range(ADDR_DATA_DIR).Value)
    d("format_dir") = CStr(ws.Range(ADDR_FORMAT_DIR).Value)
    d("ui_dir") = CStr(ws.Range(ADDR_UI_DIR).Value)
    d("backup_dir") = CStr(ws.Range(ADDR_BACKUP_DIR).Value)
    ' debugLevel = M-11 所有。holder 現値をそのまま維持。
    d("debugLevel") = SafeGetCfg("debugLevel", DEF_DEBUG_LEVEL)
    On Error Resume Next
    modConfigHolder.SetConfig d
    On Error GoTo 0
End Sub

' ============================================================
' Private Function: ValidateStorageSheet
' Role: パス系 4 dir が空でないことを検査。
' Return: True = OK, False = NG (outLabel / outAddr に NG 内容を返す)
' ============================================================
Private Function ValidateStorageSheet(ByVal ws As Worksheet, _
                                      ByRef outLabel As String, _
                                      ByRef outAddr As String) As Boolean
    If Len(CStr(ws.Range(ADDR_DATA_DIR).Value)) = 0 Then
        outAddr = ADDR_DATA_DIR : outLabel = "data_dir" : ValidateStorageSheet = False : Exit Function
    End If
    If Len(CStr(ws.Range(ADDR_FORMAT_DIR).Value)) = 0 Then
        outAddr = ADDR_FORMAT_DIR : outLabel = "format_dir" : ValidateStorageSheet = False : Exit Function
    End If
    If Len(CStr(ws.Range(ADDR_UI_DIR).Value)) = 0 Then
        outAddr = ADDR_UI_DIR : outLabel = "ui_dir" : ValidateStorageSheet = False : Exit Function
    End If
    If Len(CStr(ws.Range(ADDR_BACKUP_DIR).Value)) = 0 Then
        outAddr = ADDR_BACKUP_DIR : outLabel = "backup_dir" : ValidateStorageSheet = False : Exit Function
    End If
    ValidateStorageSheet = True
End Function

' ============================================================
' ===== 共通 helper =============================================
' ============================================================

' ============================================================
' Private Function: SafeGetCfg
' Role: modConfigHolder.GetValueOrDefault の例外安全 wrapper。
' ============================================================
Private Function SafeGetCfg(ByVal key As String, ByVal defaultValue As String) As String
    On Error Resume Next
    Dim v As String
    v = modConfigHolder.GetValueOrDefault(key, defaultValue)
    If Len(v) = 0 Then v = defaultValue
    SafeGetCfg = v
    On Error GoTo 0
End Function
```
