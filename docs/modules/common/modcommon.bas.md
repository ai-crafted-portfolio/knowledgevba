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
' ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽW・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ[・ｽE・ｽ・ｽE・ｽ: modCommon・ｽE・ｽiv2.1・ｽE・ｽA・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾊ定数・ｽE・ｽj
' ・ｽE・ｽT・ｽE・ｽv:       knowledgevba v2.1 ・ｽE・ｽﾅ具ｿｽ・ｽE・ｽL・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ關費ｿｽQ
' Version:    v2.1・ｽE・ｽi2026-05-17・ｽE・ｽA・ｽE・ｽﾅ擾ｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽAtest ・ｽE・ｽn・ｽE・ｽ[・ｽE・ｽl・ｽE・ｽX compile ・ｽE・ｽp・ｽE・ｽr・ｽE・ｽj
' ================================================================
Option Explicit

' ・ｽE・ｽo・ｽE・ｽb・ｽE・ｽN・ｽE・ｽA・ｽE・ｽb・ｽE・ｽv・ｽE・ｽﾛ趣ｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽiQ34・ｽE・ｽj
Public Const BACKUP_RETENTION_DAYS As Long = 90

' knowledge .txt ・ｽE・ｽV・ｽE・ｽX・ｽE・ｽ^・ｽE・ｽ・ｽE・ｽ・ｽE・ｽU・ｽE・ｽ`・ｽE・ｽ・ｽE・ｽ ・ｽE・ｽ・ｽE・ｽﾘり識・ｽE・ｽﾊ子・ｽE・ｽiQ47・ｽE・ｽj
Public Const KNW_STANZA_PREFIX As String = "###"
Public Const KNW_STANZA_SUFFIX As String = "###"

' ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾊ最托ｿｽ\・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽiQ37・ｽE・ｽj
Public Const SEARCH_MAX_RESULTS As Long = 100

' debugLevel ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽl・ｽE・ｽiQ7・ｽE・ｽj
Public Const DEFAULT_DEBUG_LEVEL As String = "ERROR"

' ・ｽE・ｽN・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ ActiveSheet・ｽE・ｽiADR-0053 ・ｽE・ｽ・ｽE・ｽ2.1 ・ｽE・ｽ\ / ・ｽE・ｽ・ｽE・ｽ9 M-2・ｽE・ｽAR6-01 ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽj
' SSOT: ADR-0053 ・ｽE・ｽ・ｽE・ｽ2.1・ｽE・ｽB・ｽE・ｽo・ｽE・ｽ^・ｽE・ｽC・ｽE・ｽ・ｽE・ｽ=M-05・ｽE・ｽA・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ=M-08・ｽE・ｽA・ｽE・ｽﾇ暦ｿｽ=M-02
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

' ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽO・ｽE・ｽV・ｽE・ｽ[・ｽE・ｽg・ｽE・ｽs・ｽE・ｽiv1 ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾛ趣ｿｽ・ｽE・ｽj
Public Const LOG_DATA_START_ROW As Long = 9

' ・ｽE・ｽX・ｽE・ｽ^・ｽE・ｽ・ｽE・ｽ・ｽE・ｽU・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽiv1 ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾛ趣ｿｽ・ｽE・ｽj
Public Const CHARSET_SJIS As String = "Shift_JIS"
Public Const STANZA_SEP As String = "==="

' xlsm ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽiv2.1・ｽE・ｽj
Public Const XLSM_TOUROKU As String = "・ｽE・ｽo・ｽE・ｽ^・ｽE・ｽC・ｽE・ｽ・ｽE・ｽ"
Public Const XLSM_KENSAKU As String = "・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ"
Public Const XLSM_KANRI As String = "・ｽE・ｽﾇ暦ｿｽ"

' debugLevel enum 6 ・ｽE・ｽl・ｽE・ｽimodConfigHolder ・ｽE・ｽﾉゑｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ驍ｪ Public Const ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾌゑｿｽ・ｽE・ｽﾟ本・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾅ重・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ`・ｽE・ｽj
Public Const DEBUG_LEVEL_OFF As Long = 0
Public Const DEBUG_LEVEL_ERROR As Long = 1
Public Const DEBUG_LEVEL_WARN As Long = 2
Public Const DEBUG_LEVEL_INFO As Long = 3
Public Const DEBUG_LEVEL_DEBUG As Long = 4
Public Const DEBUG_LEVEL_TRACE As Long = 5

' === v2.3 headless / progress helpers (2026-05-26) ===
'
' install_admin.bat ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ@・ｽE・ｽe・ｽE・ｽX・ｽE・ｽg・ｽE・ｽ・ｽE・ｽ Setup_admin ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ MsgBox ・ｽE・ｽo・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾄハ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽO
' ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾖの対会ｿｽ・ｽE・ｽﾅ、headless ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾆイ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽX・ｽE・ｽg・ｽE・ｽ[・ｽE・ｽ・ｽE・ｽ・ｽE・ｽi・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾌ追跡ゑｿｽ
' modCommon ・ｽE・ｽﾉ集・ｽE・ｽｵゑｿｽ・ｽE・ｽB
' modEntryFormat.IsHeadless() ・ｽE・ｽ・ｽE・ｽ v2.3 ・ｽE・ｽﾈ前・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ迹ｶ・ｽE・ｽﾝゑｿｽ・ｽE・ｽAmodEntrySettings
' ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾄばゑｿｽﾄゑｿｽ・ｽE・ｽ驍ｽ・ｽE・ｽﾟ残・ｽE・ｽu・ｽE・ｽB・ｽE・ｽV・ｽE・ｽK・ｽE・ｽﾄび出・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ modCommon.IsHeadless() ・ｽE・ｽ・ｽE・ｽ
' ・ｽE・ｽg・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽj (ThisWorkbook ・ｽE・ｽﾇ暦ｿｽ.xlsm / modSpecExamples ・ｽE・ｽ・ｽE・ｽ)・ｽE・ｽB

' ================================================================
' Public Function: IsHeadless
' ・ｽE・ｽT・ｽE・ｽv: COM ・ｽE・ｽI・ｽE・ｽ[・ｽE・ｽg・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ[・ｽE・ｽV・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽz・ｽE・ｽ・ｽE・ｽ (Excel.Application.Visible=False
'       ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ/・ｽE・ｽﾜゑｿｽ・ｽE・ｽ・ｽE・ｽ Application.Interactive=False) ・ｽE・ｽﾅ難ｿｽ・ｽE・ｽ・ｽE・ｽE・ｽ・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ
'       ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ閧ｷ・ｽE・ｽ・ｽE・ｽBMsgBox ・ｽE・ｽ}・ｽE・ｽ~・ｽE・ｽK・ｽE・ｽ[・ｽE・ｽh・ｽE・ｽﾌ擾ｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾆゑｿｽ・ｽE・ｽﾄ使・ｽE・ｽ・ｽE・ｽ・ｽE・ｽB
' ・ｽE・ｽﾟゑｿｽl: True = headless (MsgBox ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽo・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾄはゑｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾈゑｿｽ)
'         False = ・ｽE・ｽﾊ擾ｿｽﾌ対話・ｽE・ｽ・ｽE・ｽ・ｽE・ｽs
' ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽl:
'   - Application.Interactive ・ｽE・ｽ・ｽE・ｽ COM ・ｽE・ｽz・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾅゑｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ True ・ｽE・ｽﾈのゑｿｽ
'     Application.Visible ・ｽE・ｽﾆゑｿｽ OR ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾉゑｿｽ・ｽE・ｽ・ｽE・ｽ (modEntryFormat ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ
'     IsHeadless ・ｽE・ｽﾆ難ｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽW・ｽE・ｽb・ｽE・ｽN)・ｽE・ｽB
'   - On Error Resume Next ・ｽE・ｽﾅ托ｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ謫ｾ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽs・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ False ・ｽE・ｽ・ｽE・ｽﾔゑｿｽ
'     (= ・ｽE・ｽﾊ擾ｿｽ・ｽE・ｽ・ｽE・ｽs・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽAMsgBox ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽo・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾖフ・ｽE・ｽH・ｽE・ｽ[・ｽE・ｽ・ｽE・ｽ・ｽE・ｽX・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ[)・ｽE・ｽB
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
' ・ｽE・ｽT・ｽE・ｽv: Setup_admin / clsSetupOrchestrator ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ install ・ｽE・ｽo・ｽE・ｽH・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ
'       ・ｽE・ｽu・ｽE・ｽﾇゑｿｽ・ｽE・ｽﾜで進・ｽE・ｽｾゑｿｽ・ｽE・ｽv・ｽE・ｽ・ｽE・ｽ c:\temp\setup_admin_progress.txt
'       ・ｽE・ｽ・ｽE・ｽ append ・ｽE・ｽﾅ擾ｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾞゑｿｽ・ｽE・ｽﾟのヘ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽp・ｽE・ｽ[・ｽE・ｽB
'       Debug.Print ・ｽE・ｽ・ｽE・ｽ VBE Immediate ・ｽE・ｽE・ｽE・ｽB・ｽE・ｽ・ｽE・ｽ・ｽE・ｽh・ｽE・ｽE・ｽE・ｽﾉゑｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽo・ｽE・ｽﾈゑｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾟ、
'       headless ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽs・ｽE・ｽ・ｽE・ｽ Debug.Print ・ｽE・ｽﾆ難ｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽt・ｽE・ｽ@・ｽE・ｽC・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾉ残・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾚ的・ｽE・ｽB
' ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ: msg = ・ｽE・ｽi・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽb・ｽE・ｽZ・ｽE・ｽ[・ｽE・ｽW 1 ・ｽE・ｽs (・ｽE・ｽ・ｽE・ｽ・ｽE・ｽs・ｽE・ｽﾍ難ｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾅ付・ｽE・ｽ^)
' ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽl:
'   - ・ｽE・ｽo・ｽE・ｽﾍ撰ｿｽ・ｽE・ｽ c:\temp\setup_admin_progress.txt ・ｽE・ｽﾅ抵ｿｽBc:\temp ・ｽE・ｽ・ｽE・ｽ
'     ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾝゑｿｽ・ｽE・ｽﾈゑｿｽ・ｽE・ｽ鼾・・ｽ・ｽ・ｽE・ｽ MkDir ・ｽE・ｽﾅ趣ｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ (・ｽE・ｽ・ｽE・ｽ・ｽE・ｽs・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾄゑｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ)・ｽE・ｽB
'   - On Error Resume Next ・ｽE・ｽﾅ擾ｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾝ趣ｿｽ・ｽE・ｽs・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾍ会ｿｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾈゑｿｽ
'     (・ｽE・ｽi・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽO・ｽE・ｽ・ｽE・ｽ best-effort・ｽE・ｽA・ｽE・ｽ・ｽE・ｽ・ｽE・ｽs・ｽE・ｽﾅ本・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ~・ｽE・ｽﾟなゑｿｽ)・ｽE・ｽB
'   - ・ｽE・ｽ`・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾍ呼び出・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ "[setup] [hh:nn:ss] step N : doing X" ・ｽE・ｽ・ｽE・ｽ
'     ・ｽE・ｽg・ｽE・ｽﾝ暦ｿｽ・ｽE・ｽﾄて渡・ｽE・ｽ・ｽE・ｽ・ｽE・ｽO・ｽE・ｽ・ｽE・ｽB・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾅはタ・ｽE・ｽC・ｽE・ｽ・ｽE・ｽ・ｽE・ｽX・ｽE・ｽ^・ｽE・ｽ・ｽE・ｽ・ｽE・ｽv・ｽE・ｽ・ｽE・ｽt・ｽE・ｽ^・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾈゑｿｽ・ｽE・ｽB
'   - ・ｽE・ｽt・ｽE・ｽ@・ｽE・ｽC・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ UTF-8 ・ｽE・ｽﾅはなゑｿｽ Shift_JIS ・ｽE・ｽﾅ擾ｿｽ・ｽE・ｽ・ｽE・ｽ (Open For Append
'     ・ｽE・ｽﾌデ・ｽE・ｽt・ｽE・ｽH・ｽE・ｽ・ｽE・ｽ・ｽE・ｽg・ｽE・ｽ・ｽE・ｽ・ｽE・ｽP・ｽE・ｽ[・ｽE・ｽ・ｽE・ｽ = CP932 ・ｽE・ｽﾂ具ｿｽ・ｽE・ｽz・ｽE・ｽ・ｽE・ｽ)・ｽE・ｽB
' ================================================================
Public Sub AppendProgressLog(ByVal msg As String)
    On Error Resume Next
    Dim logDir As String
    Dim logPath As String
    logDir = "C:\temp"
    logPath = logDir & "\setup_admin_progress.txt"

    ' c:\temp ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽﾎ搾ｿｽ・ｽE・ｽ (best-effort)
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
' ・ｽE・ｽT・ｽE・ｽv: AppendProgressLog ・ｽE・ｽp・ｽE・ｽﾌタ・ｽE・ｽC・ｽE・ｽ・ｽE・ｽ・ｽE・ｽX・ｽE・ｽ^・ｽE・ｽ・ｽE・ｽ・ｽE・ｽv・ｽE・ｽv・ｽE・ｽ・ｽE・ｽ・ｽE・ｽt・ｽE・ｽB・ｽE・ｽN・ｽE・ｽX・ｽE・ｽｶ撰ｿｽ・ｽE・ｽB
'       "[setup] [hh:nn:ss] " ・ｽE・ｽ・ｽE・ｽﾔゑｿｽ・ｽE・ｽB・ｽE・ｽﾄび出・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ・ｽE・ｽ
'       AppendProgressLog ProgressTs() & "step 3 : LoadConfig done"
'       ・ｽE・ｽﾌよう・ｽE・ｽﾉ使・ｽE・ｽ・ｽE・ｽ・ｽE・ｽB
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
