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
Public Const STARTUP_SHEET_TOUROKU As String = "ナレッジ登録"  ' R-3-b JP label
Public Const STARTUP_SHEET_KENSAKU As String = "ナレッジ検索"  ' R-3-b JP label
Public Const STARTUP_SHEET_KANRI As String = "フォーマット一覧"  ' R-3-b JP label

' Phase R-1-j (2026-05-28): SHEET_* constants required by legacy modSheetMap /
' clsSearchEngine. Previously undefined, causing project-wide compile error.
Public Const SHEET_MAIN As String = "M-01"  ' R-3-b: M-01 メニューは v2.3 廃止(該当シート無し・据置)
Public Const SHEET_FORMAT_LIST As String = "フォーマット一覧"  ' R-3-b: sheet 名は JP(SSOT)。UI-def id は M-NN リテラル
Public Const SHEET_FORMAT_DESIGN As String = "フォーマット設計"  ' R-3-b JP
Public Const SHEET_FORMAT_PREVIEW As String = "プレビュー"  ' R-3-b JP
Public Const SHEET_KNW_SAVE As String = "ナレッジ登録"  ' R-3-b JP
Public Const SHEET_KNW_EDIT As String = "ナレッジ修正"  ' R-3-b JP
Public Const SHEET_KNW_LIST As String = "M-07"  ' R-3-b: M-07 一覧は v2.3 廃止(該当シート無し・据置)
Public Const SHEET_SEARCH As String = "ナレッジ検索"  ' R-3-b JP
Public Const SHEET_KNW_DISPLAY As String = "ナレッジ表示"  ' R-3-b JP
Public Const SHEET_STORAGE As String = "格納先設定"  ' R-3-b JP
Public Const SHEET_SETTINGS As String = "設定"  ' R-3-b JP
Public Const SHEET_FORMAT_CHANGE_CHECK As String = "フォーマット変更チェック"  ' R-3-b JP
Public Const SHEET_LOG As String = "LOG"
Public Const SHEET_MIGRATION As String = "M-12"

' ログシート行（v1 から維持）
Public Const LOG_DATA_START_ROW As Long = 9

' スタンザ書式（v1 から維持）
Public Const CHARSET_SJIS As String = "Shift_JIS"
Public Const STANZA_SEP As String = "==="

' xlsm 名（v2.1）
Public Const XLSM_TOUROKU As String = "登録修正"
Public Const XLSM_KENSAKU As String = "検索"
Public Const XLSM_KANRI As String = "管理"

' debugLevel enum 6 値（modConfigHolder にもあるが Public Const 整合のため本ファイルで重複定義）
Public Const DEBUG_LEVEL_OFF As Long = 0
Public Const DEBUG_LEVEL_ERROR As Long = 1
Public Const DEBUG_LEVEL_WARN As Long = 2
Public Const DEBUG_LEVEL_INFO As Long = 3
Public Const DEBUG_LEVEL_DEBUG As Long = 4
Public Const DEBUG_LEVEL_TRACE As Long = 5

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
    On Error Resume Next
    Dim notInteractive As Boolean
    Dim notVisible As Boolean
    notInteractive = (Application.Interactive = False)
    notVisible = (Application.Visible = False)
    IsHeadless = (notInteractive Or notVisible)
    On Error GoTo 0
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
    On Error Resume Next
    Dim logDir As String
    Dim logPath As String
    logDir = "C:\temp"
    logPath = logDir & "\setup_admin_progress.txt"

    ' c:\temp が無ければ作る (best-effort)
    If Len(Dir(logDir, vbDirectory)) = 0 Then
        MkDir logDir
    End If

    Dim fh As Integer
    fh = FreeFile
    Open logPath For Append As #fh
    Print #fh, msg
    Close #fh
    On Error GoTo 0
End Sub

' ================================================================
' Public Function: ProgressTs
' 概要: AppendProgressLog 用のタイムスタンププレフィクスを生成。
'       "[setup] [hh:nn:ss] " を返す。呼び出し側で
'       AppendProgressLog ProgressTs() & "step 3 : LoadConfig done"
'       のように使う。
' ================================================================
Public Function ProgressTs() As String
    ProgressTs = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
End Function

' ================================================================
' BUG-4 related (2026-05-30): M-12 restore-backup feature constants.
' Referenced by modEntryKnowledge.Btn_RestoreBackup. JP literals are
' built via ChrW() per memory rule (Edit/Write CP932 conversion bug
' avoidance) and exposed as Public Property Get because VBA Const
' cannot use a runtime function. BACKUP_SUBFOLDER is ASCII only.
' ================================================================
Public Const BACKUP_SUBFOLDER As String = "backup"

Public Property Get MSG_RESTORE_TITLE() As String
    MSG_RESTORE_TITLE = ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H5FA9) & ChrW(&H65E7)
End Property

Public Property Get MSG_RESTORE_NO_BACKUP_DETAIL() As String
    MSG_RESTORE_NO_BACKUP_DETAIL = ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H304C) & ChrW(&H898B) & ChrW(&H3064) & ChrW(&H304B) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093)
End Property

Public Property Get MSG_RESTORE_NO_BACKUP_ACTION() As String
    MSG_RESTORE_NO_BACKUP_ACTION = ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & ChrW(&H30C0) & ChrW(&H306B) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H3092) & ChrW(&H914D) & ChrW(&H7F6E) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304B) & ChrW(&H3089) & ChrW(&H518D) & ChrW(&H5EA6) & ChrW(&H30DC) & ChrW(&H30BF) & ChrW(&H30F3) & ChrW(&H3092) & ChrW(&H62BC) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044)
End Property

Public Property Get MSG_RESTORE_NOT_IMPL_DETAIL() As String
    MSG_RESTORE_NOT_IMPL_DETAIL = ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H5FA9) & ChrW(&H65E7) & ChrW(&H6A5F) & ChrW(&H80FD) & ChrW(&H306F) & ChrW(&H6B21) & ChrW(&H56DE) & " Sprint " & ChrW(&H3067) & ChrW(&H5B9F) & ChrW(&H88C5) & ChrW(&H4E88) & ChrW(&H5B9A) & ChrW(&H3067) & ChrW(&H3059)
End Property
```
