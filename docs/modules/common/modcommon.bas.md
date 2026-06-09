---
title: modCommon.bas
description: modCommon.bas のソースコード（コピペ用）
---

# modCommon.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modCommon.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modCommon"
' ================================================================
' モジュール: modCommon（v2.1、共通定数）
' 概要:       knowledgevba v2.1 で共有する定数群
' Version:    v2.1（2026-05-17、最小実装、test ハーネス compile 用途）
' ================================================================
Option Explicit

' バックアップ保持日数（Q34）
Public Const BACKUP_RETENTION_DAYS As Long = 90

' knowledge .txt 新スタンザ形式 区切り識別子（Q47）
Public Const KNW_STANZA_PREFIX As String = "###"
Public Const KNW_STANZA_SUFFIX As String = "###"

' 検索結果最大表示件数（Q37）
Public Const SEARCH_MAX_RESULTS As Long = 100

' debugLevel 既定値（Q7）
Public Const DEFAULT_DEBUG_LEVEL As String = "ERROR"

' 起動時 ActiveSheet（ADR-0053 §2.1 表 / §9 M-2、R6-01 是正）
' SSOT: ADR-0053 §2.1。登録修正=M-05、検索=M-08、管理=M-02

' Phase R-1-j (2026-05-28): SHEET_* constants required by legacy modSheetMap /
' clsSearchEngine. Previously undefined, causing project-wide compile error.
Public Const SHEET_MAIN As String = "M-01"  ' R-3-b: M-01 メニューは v2.3 廃止(該当シート無し・据置)
Public Const SHEET_KNW_LIST As String = "M-07"  ' R-3-b: M-07 一覧は v2.3 廃止(該当シート無し・据置)
Public Const SHEET_LOG As String = "LOG"
Public Const SHEET_MIGRATION As String = "M-12"

' ログシート行（v1 から維持）
Public Const LOG_DATA_START_ROW As Long = 9

' スタンザ書式（v1 から維持）
Public Const CHARSET_SJIS As String = "Shift_JIS"
Public Const STANZA_SEP As String = "==="

' xlsm 名（v2.1）

' debugLevel enum 6 値（modConfigHolder にもあるが Public Const 整合のため本ファイルで重複定義）
Public Const DEBUG_LEVEL_OFF As Long = 0
Public Const DEBUG_LEVEL_ERROR As Long = 1
Public Const DEBUG_LEVEL_WARN As Long = 2
Public Const DEBUG_LEVEL_INFO As Long = 3
Public Const DEBUG_LEVEL_DEBUG As Long = 4
Public Const DEBUG_LEVEL_TRACE As Long = 5
' [ADR-0100][gDebugLevel] Public global, set by Workbook_Open via
' modConfigHolder.GetDebugLevel(). Fallback = DEFAULT_DEBUG_LEVEL_FALLBACK
' if config load fails. All [D-NNNN] Debug.Print lines are guarded by
' `If modCommon.gDebugLevel >= DEBUG_LEVEL_X Then ...`.
Public gDebugLevel As Long

' [BTN-GUARD-LOGGER-FIX] modBtnGuard で参照される logger holder
Public gLogger As clsLogger

' Fallback level when modConfigHolder.GetDebugLevel() fails to load.
' Set to DEBUG (4) so dev/install can still see [D-NNNN] STEP/EXIT-OK,
' but TRACE-level ENTER/EXIT is suppressed by default.
Public Const DEFAULT_DEBUG_LEVEL_FALLBACK As Long = 4    ' = DEBUG_LEVEL_DEBUG


' BUG-4 (2026-05-30): M-12 restore-backup feature constant.
' Referenced by modEntryKnowledge.Btn_RestoreBackup. ASCII only.
Public Const BACKUP_SUBFOLDER As String = "backup"

' === v2.3 headless / progress helpers (2026-05-26) ===
'
' install_admin.bat 実機テストで Setup_admin が裏で MsgBox 出してハング
' する問題への対応で、headless 判定とインストーラ進捗の追跡を
' modCommon に集約した。
' modEntryFormat.IsHeadless() は v2.3 以前から存在し、modEntrySettings
' から呼ばれているため残置。新規呼び出しは modCommon.IsHeadless() を
' 使う方針 (ThisWorkbook 管理.xlsm / modSpecExamples 等)。

' ================================================================
' Public Function: IsHeadless
' 概要: COM オートメーション配下 (Excel.Application.Visible=False
'       かつ/または Application.Interactive=False) で動作中かを
'       判定する。MsgBox 抑止ガードの条件式として使う。
' 戻り値: True = headless (MsgBox を出してはいけない)
'         False = 通常の対話実行
' 備考:
'   - Application.Interactive は COM 配下でも既定 True なので
'     Application.Visible との OR 判定にする (modEntryFormat 既存
'     IsHeadless と同じロジック)。
'   - On Error Resume Next で属性取得失敗時は False を返す
'     (= 通常実行扱い、MsgBox を出す側へフォールスルー)。
' ================================================================
Public Function IsHeadless() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0931] modCommon.IsHeadless ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim notInteractive As Boolean
    Dim notVisible As Boolean
    notInteractive = (Application.Interactive = False)
    notVisible = (Application.Visible = False)
    IsHeadless = (notInteractive Or notVisible)
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0932] modCommon.IsHeadless EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' Public Sub: AppendProgressLog
' 概要: Setup_admin / clsSetupOrchestrator 等の install 経路から
'       「どこまで進んだか」を c:\temp\setup_admin_progress.txt
'       に append で書き込むためのヘルパー。
'       Debug.Print は VBE Immediate ウィンドウにしか出ないため、
'       headless 実行で Debug.Print と等価情報をファイルに残す目的。
' 引数: msg = 進捗メッセージ 1 行 (改行は内部で付与)
' 備考:
'   - 出力先は c:\temp\setup_admin_progress.txt 固定。c:\temp が
'     存在しない場合は MkDir で自動生成 (失敗しても無視)。
'   - On Error Resume Next で書き込み失敗時は何もしない
'     (進捗ログは best-effort、失敗で本処理を止めない)。
'   - 形式は呼び出し側で "[setup] [hh:nn:ss] step N : doing X" を
'     組み立てて渡す前提。ここではタイムスタンプを付与しない。
'   - ファイルは UTF-8 ではなく Shift_JIS で書く (Open For Append
'     のデフォルトロケール = CP932 環境想定)。
' ================================================================
Public Sub AppendProgressLog(ByVal msg As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0933] modCommon.AppendProgressLog ENTER"  ' [ADR-0100]
    ' [ADR-0100][gDebugLevel] hardened: GoTo ErrHandler + fallback path.
    Dim logPath As String
    Dim fh As Integer
    On Error GoTo ErrHandler
    logPath = "C:\kvba\publish\dist_v2\setup_admin_progress.txt"
    fh = FreeFile
    Open logPath For Append As #fh
    Print #fh, msg
    Close #fh
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0934] modCommon.AppendProgressLog EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    ' Fallback when primary path is not writable (no C:\kvba\..., perm denied, etc).
    ' Try %USERPROFILE%\Documents\setup_admin_progress.txt; tag entries with [FALLBACK].
    On Error Resume Next
    Close #fh
    Dim altPath As String
    altPath = Environ$("USERPROFILE") & "\Documents\setup_admin_progress.txt"
    fh = FreeFile
    Open altPath For Append As #fh
    Print #fh, "[FALLBACK] " & msg
    Close #fh
    Err.Clear
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1832] modCommon.AppendProgressLog EXIT-OK-FALLBACK"  ' [ADR-0100]
End Sub

' ================================================================
' Public Function: ProgressTs
' 概要: AppendProgressLog 用のタイムスタンププレフィクスを生成。
'       "[setup] [hh:nn:ss] " を返す。呼び出し側で
'       AppendProgressLog ProgressTs() & "step 3 : LoadConfig done"
'       のように使う。
' ================================================================
Public Function ProgressTs() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0935] modCommon.ProgressTs ENTER"  ' [ADR-0100]
    ProgressTs = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
End Function


Public Property Get MSG_RESTORE_TITLE() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0936] modCommon.MSG_RESTORE_TITLE ENTER"  ' [ADR-0100]
    MSG_RESTORE_TITLE = ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H5FA9) & ChrW(&H65E7)
End Property

Public Property Get MSG_RESTORE_NO_BACKUP_DETAIL() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0937] modCommon.MSG_RESTORE_NO_BACKUP_DETAIL ENTER"  ' [ADR-0100]
    MSG_RESTORE_NO_BACKUP_DETAIL = ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H304C) & ChrW(&H898B) & ChrW(&H3064) & ChrW(&H304B) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093)
End Property

Public Property Get MSG_RESTORE_NO_BACKUP_ACTION() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0938] modCommon.MSG_RESTORE_NO_BACKUP_ACTION ENTER"  ' [ADR-0100]
    MSG_RESTORE_NO_BACKUP_ACTION = ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & ChrW(&H30C0) & ChrW(&H306B) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H3092) & ChrW(&H914D) & ChrW(&H7F6E) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304B) & ChrW(&H3089) & ChrW(&H518D) & ChrW(&H5EA6) & ChrW(&H30DC) & ChrW(&H30BF) & ChrW(&H30F3) & ChrW(&H3092) & ChrW(&H62BC) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044)
End Property

Public Property Get MSG_RESTORE_NOT_IMPL_DETAIL() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0939] modCommon.MSG_RESTORE_NOT_IMPL_DETAIL ENTER"  ' [ADR-0100]
    MSG_RESTORE_NOT_IMPL_DETAIL = ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H5FA9) & ChrW(&H65E7) & ChrW(&H6A5F) & ChrW(&H80FD) & ChrW(&H306F) & ChrW(&H6B21) & ChrW(&H56DE) & " Sprint " & ChrW(&H3067) & ChrW(&H5B9F) & ChrW(&H88C5) & ChrW(&H4E88) & ChrW(&H5B9A) & ChrW(&H3067) & ChrW(&H3059)
End Property

' ----------------------------------------------------------------
' ADR-0006/0090/0094 JP literal removal (Const -> Property Get).
' Each Property Get returns the same JP string via ChrW(&H...).
' Callers reference these by identifier exactly like the old Const.
' ----------------------------------------------------------------
Public Property Get STARTUP_SHEET_TOUROKU() As String
    STARTUP_SHEET_TOUROKU = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H767B) & ChrW(&H9332)
End Property

Public Property Get STARTUP_SHEET_KENSAKU() As String
    STARTUP_SHEET_KENSAKU = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H691C) & ChrW(&H7D22)
End Property

Public Property Get STARTUP_SHEET_KANRI() As String
    STARTUP_SHEET_KANRI = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H4E00) & ChrW(&H89A7)
End Property

Public Property Get SHEET_FORMAT_LIST() As String
    SHEET_FORMAT_LIST = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H4E00) & ChrW(&H89A7)
End Property

Public Property Get SHEET_FORMAT_DESIGN() As String
    SHEET_FORMAT_DESIGN = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H8A2D) & ChrW(&H8A08)
End Property

Public Property Get SHEET_FORMAT_PREVIEW() As String
    SHEET_FORMAT_PREVIEW = ChrW(&H30D7) & ChrW(&H30EC) & ChrW(&H30D3) & ChrW(&H30E5) & ChrW(&H30FC)
End Property

Public Property Get SHEET_KNW_SAVE() As String
    SHEET_KNW_SAVE = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H767B) & ChrW(&H9332)
End Property

Public Property Get SHEET_KNW_EDIT() As String
    SHEET_KNW_EDIT = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H4FEE) & ChrW(&H6B63)
End Property

Public Property Get SHEET_SEARCH() As String
    SHEET_SEARCH = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H691C) & ChrW(&H7D22)
End Property

Public Property Get SHEET_KNW_DISPLAY() As String
    SHEET_KNW_DISPLAY = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H8868) & ChrW(&H793A)
End Property

Public Property Get SHEET_STORAGE() As String
    SHEET_STORAGE = ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & ChrW(&H8A2D) & ChrW(&H5B9A)
End Property

Public Property Get SHEET_SETTINGS() As String
    SHEET_SETTINGS = ChrW(&H8A2D) & ChrW(&H5B9A)
End Property

Public Property Get SHEET_FORMAT_CHANGE_CHECK() As String
    SHEET_FORMAT_CHANGE_CHECK = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H5909) & ChrW(&H66F4) & ChrW(&H30C1) & ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF)
End Property

Public Property Get XLSM_TOUROKU() As String
    XLSM_TOUROKU = ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63)
End Property

Public Property Get XLSM_KENSAKU() As String
    XLSM_KENSAKU = ChrW(&H691C) & ChrW(&H7D22)
End Property

Public Property Get XLSM_KANRI() As String
    XLSM_KANRI = ChrW(&H7BA1) & ChrW(&H7406)
End Property
```
