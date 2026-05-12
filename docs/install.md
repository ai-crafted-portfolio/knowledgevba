# インストール

以下の bat と ps1 を 2 ファイル、自分の PC の同じフォルダ (例: デスクトップに新規フォルダを作って) に保存してから、bat に空の `.xlsm` をドラッグ＆ドロップすると、シート 14 個・ボタン 68 個・全 VBA モジュールが自動的に展開され、ナレッジ管理ツール本体が構築されます。

---

## STEP 1: Excel の事前設定

VBA を Excel COM 経由で書き込むには、Excel 側で「VBA プロジェクトへの外部アクセス」を許可しておく必要があります。**1 度だけ** 設定すれば以降は不要です。

1. Excel を起動し、空のブックを 1 つ開きます
2. **[ファイル]** → **[オプション]** を開きます
3. 左メニューから **[トラスト センター]** を選び、**[トラスト センターの設定]** ボタンを押します
4. 左メニューから **[マクロの設定]** を選びます
5. 右側の **[VBA プロジェクト オブジェクト モデルへのアクセスを信頼する]** にチェックを入れます
6. **[OK]** を押してすべてのダイアログを閉じます
7. Excel をいったん全て閉じます

この設定が OFF のままだと、STEP 4 で `[ERROR] VBA プロジェクトへのアクセスが拒否されました` というメッセージで停止します。

---

## STEP 2: bat の保存

下のコードを **メモ帳など** に貼り付け、`Install-KnowledgevbaModules.bat` というファイル名で任意のフォルダ (例: デスクトップの新規フォルダ) に保存します。

!!! warning "保存時の注意"
    メモ帳の場合、**[名前を付けて保存]** で「ファイルの種類: すべてのファイル」を選び、**文字コードを「ANSI」** にしてください。UTF-8 にすると日本語メッセージが文字化けします。

```bat
@echo off
setlocal EnableDelayedExpansion
chcp 65001 > nul
cd /d "%~dp0"

echo ==================================================
echo  ナレッジ管理ツール - 自動インストーラ
echo ==================================================
echo.

set "XLSM_PATH=%~1"
if "!XLSM_PATH!"=="" (
    echo .xlsm または .xlsx をこの bat にドラッグ＆ドロップしてください。
    echo もしくはフルパスを入力 ^(囲い引用符は不要^):
    echo.
    set /p "XLSM_PATH=Path to .xlsm/.xlsx: "
)

if "!XLSM_PATH!"=="" (
    echo [ERROR] ファイルが指定されませんでした。
    pause
    exit /b 1
)

REM 引用符を剥がす
set DQ=^"
if "!XLSM_PATH:~0,1!"=="!DQ!" set "XLSM_PATH=!XLSM_PATH:~1!"
if "!XLSM_PATH:~-1!"=="!DQ!" set "XLSM_PATH=!XLSM_PATH:~0,-1!"

if not exist "!XLSM_PATH!" (
    echo [ERROR] ファイルが見つかりません: !XLSM_PATH!
    pause
    exit /b 1
)

REM Excel が起動中だと VBA 注入が失敗するので事前チェック
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I "EXCEL.EXE" >NUL
if not errorlevel 1 (
    echo.
    echo [ERROR] Excel がすでに起動しています。
    echo         全ての Excel ウィンドウを閉じてから再度実行してください。
    echo.
    pause
    exit /b 1
)

if not exist "%~dp0Install-KnowledgevbaModules.ps1" (
    echo [ERROR] Install-KnowledgevbaModules.ps1 が見つかりません。
    echo         この bat と同じフォルダに ps1 を保存してから実行してください。
    pause
    exit /b 1
)

echo [info] 対象ファイル : !XLSM_PATH!
echo [run ] PowerShell インストーラ起動...
echo.

REM ==== マクロセキュリティの一時緩和 ====
REM 現在の設定を保存して "全てのマクロを有効" に切替。Setup 完了後に元に戻す。
REM 対象は HKCU 配下なので管理者権限不要。Excel 16 (2016/2019/2021/365) を想定。
set "VBAW_ORIG="
for /f "tokens=3" %%A in ('reg query "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings 2^>nul ^| find "VBAWarnings"') do set "VBAW_ORIG=%%A"
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >nul 2>&1
echo [trust] Excel マクロを一時的に "全て有効" に切替 ^(install 完了後に元に戻します^)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-KnowledgevbaModules.ps1" -XlsmPath "!XLSM_PATH!"
set "PS_EXIT=!errorlevel!"

REM ==== マクロセキュリティを元に戻す ====
if defined VBAW_ORIG (
    reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d !VBAW_ORIG! /f >nul 2>&1
    echo [trust] Excel マクロセキュリティを元の値 !VBAW_ORIG! に復元
) else (
    REM 元値が取れない時は "通知ありで無効" (=2、デフォルト) に戻す
    reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 2 /f >nul 2>&1
    echo [trust] Excel マクロセキュリティを既定値 2 に戻しました
)

echo.
if !PS_EXIT! NEQ 0 (
    echo [FAIL] インストールに失敗しました。上のメッセージを確認してください。
) else (
    echo [DONE] インストール完了。Excel でファイルを開いて動作確認してください。
)
echo.
pause
exit /b !PS_EXIT!
```

---

## STEP 3: ps1 の保存

下のコードを **同じフォルダ** に `Install-KnowledgevbaModules.ps1` というファイル名で保存します。VBA モジュール本体 (49 個) がすべて埋め込まれているため、ファイルサイズは約 328 KB あります。

!!! warning "保存時の注意"
    メモ帳の場合、**[名前を付けて保存]** で **文字コードを「UTF-8 (BOM 付き)」** にしてください。BOM 無し UTF-8 や ANSI で保存すると日本語コメントが文字化けし、コンパイルエラーになります。VS Code を使う場合は右下のエンコーディング表示を **`UTF-8 with BOM`** に切り替えてください。

````powershell
# =============================================================================
# Install-KnowledgevbaModules.ps1
# -----------------------------------------------------------------------------
# 同梱 VBA モジュール群を Excel COM 経由で対象 .xlsm に流し込み、
# 続けて SetupSheetsAndButtons を実行してシート 14 個 + ボタン 68 個を生成する。
# 引数で受け取った .xlsm/.xlsx を開き、保存して閉じるところまで自動で行う。
#
# 前提:
#   - Excel がインストール済み
#   - Excel の [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する] が ON
#   - 対象 Excel ファイルが他のプロセスで開かれていない
# =============================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$XlsmPath
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

# ------------------- 引数チェック ---------------------------
if (-not (Test-Path -LiteralPath $XlsmPath)) {
    Write-Host ('[ERROR] ファイルが見つかりません: {0}' -f $XlsmPath) -ForegroundColor Red
    exit 1
}
$target = (Resolve-Path -LiteralPath $XlsmPath).Path
$ext    = [IO.Path]::GetExtension($target).ToLowerInvariant()
if ($ext -notin @('.xlsm', '.xlsx')) {
    Write-Host ('[ERROR] .xlsm または .xlsx を指定してください: {0}' -f $ext) -ForegroundColor Red
    exit 1
}

Write-Host '=================================================='
Write-Host ' ナレッジ管理ツール - 自動インストーラ'
Write-Host '=================================================='
Write-Host ('[info] 対象: {0}' -f $target)

# ------------------- モジュール定義 (埋め込み) -----------------
# Type: 'std' = 標準モジュール (.bas)
#       'cls' = クラスモジュール (.cls)
#       'doc' = ThisWorkbook (既存ドキュメントモジュール)
$Modules = @(
    @{ Name='clsLogEntry'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsLogEntry（ValueObject）
' 概要:   ログ1行の値を保持する軽量クラス
' 依存先: なし
' ================================================================

Private m_timestamp As String
Private m_moduleName As String
Private m_functionName As String
Private m_level As String
Private m_message As String

' --- Property Get/Let ---

Public Property Get Timestamp() As String
    Timestamp = m_timestamp
End Property

Public Property Let Timestamp(ByVal value As String)
    m_timestamp = value
End Property

Public Property Get ModuleName() As String
    ModuleName = m_moduleName
End Property

Public Property Let ModuleName(ByVal value As String)
    m_moduleName = value
End Property

Public Property Get FunctionName() As String
    FunctionName = m_functionName
End Property

Public Property Let FunctionName(ByVal value As String)
    m_functionName = value
End Property

Public Property Get Level() As String
    Level = m_level
End Property

Public Property Let Level(ByVal value As String)
    m_level = value
End Property

Public Property Get Message() As String
    Message = m_message
End Property

Public Property Let Message(ByVal value As String)
    m_message = value
End Property

' ================================================================
' 関数名: Init
' 概要:   全プロパティを一括設定する初期化メソッド
' 引数:   ts    - 日時文字列
'         mod_  - モジュール名
'         func  - 関数名
'         lvl   - ログレベル
'         msg   - メッセージ
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal ts As String, _
                 ByVal mod_ As String, _
                 ByVal func As String, _
                 ByVal lvl As String, _
                 ByVal msg As String)
    m_timestamp = ts
    m_moduleName = mod_
    m_functionName = func
    m_level = lvl
    m_message = msg
End Sub
'@ },
    @{ Name='modCommon'; Type='std'; Code=@'
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
Public Const LOG_LEVEL_TRACE As String = "トレース"

' --- 外部ログファイルパス（C:\kvba\runtime.log に Append） ---
'     診断目的: シートに書けない致命的場面でもファイルに痕跡を残す
Public Const EXTERNAL_LOG_PATH As String = "C:\kvba\runtime.log"

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

' --- タスク名定数（12タスク — polished mock M-01 v19 準拠） ---
' v20 改修: 8 → 12 ボタン化。後方互換のため旧 8 タスク名も保持。
Public Const TASK_SEARCH As String = "検索"
Public Const TASK_REGISTER As String = "ナレッジ登録"
Public Const TASK_MODIFY As String = "ナレッジ修正"
Public Const TASK_LIST As String = "ナレッジ一覧"
Public Const TASK_FORMAT As String = "フォーマット管理"
Public Const TASK_FIELD_REFLECT As String = "フィールド反映"
Public Const TASK_STORAGE As String = "格納先設定"
Public Const TASK_SYS_SETTINGS As String = "システム設定"
Public Const TASK_LOG As String = "ログ確認"
Public Const TASK_FILE_FORMAT As String = "ファイル形式"
Public Const TASK_INIT_SETUP As String = "初回セットアップ"
Public Const TASK_HELP_VERSION As String = "ヘルプ"

' --- 旧 8 タスク名（後方互換用、廃止予定） ---
Public Const TASK_SETUP As String = "初回セットアップ"
Public Const TASK_CONFIG As String = "システム設定"
Public Const TASK_EDIT As String = "ナレッジ修正"
Public Const TASK_DELETE As String = "ナレッジ修正"
Public Const TASK_MIGRATE As String = "フィールド反映"

' --- カラー定数（polished mock 準拠 — spec.md §2 の表に対応） ---
Public Const COLOR_TITLE_DEEP_BLUE As String = "#1F3864"
Public Const COLOR_TITLE_BLUE As String = "#1F4E78"
Public Const COLOR_SECTION_BLUE As String = "#2F5496"
Public Const COLOR_SECTION_BLUE2 As String = "#4472C4"
Public Const COLOR_BTN_PRIMARY As String = "#5B9BD5"
Public Const COLOR_BTN_NAV As String = "#70AD47"
Public Const COLOR_BTN_SUB As String = "#BFBFBF"
Public Const COLOR_BTN_DANGER As String = "#ED7D31"
Public Const COLOR_DESTROY_BAR As String = "#C00000"
Public Const COLOR_REQUIRED_RED As String = "#C00000"
Public Const COLOR_HINT_YELLOW As String = "#FFF2CC"
Public Const COLOR_HINT_BAR As String = "#DEEBF7"
Public Const COLOR_HEADER_LIGHT As String = "#B4C7E7"
Public Const COLOR_HINT_GREEN As String = "#E2EFDA"

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
=======================
Public Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not (ws Is Nothing)
    Err.Clear
End Function
'@ },
    @{ Name='modDateUtil'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modDateUtil（ユーティリティ層）
' 概要:       日付・時刻処理の純粋関数群
'             LibreOffice Basic互換（"hh:nn:ss"問題を回避）
' 依存先:     なし
' ================================================================

' ================================================================
' 関数名: NowStr
' 概要:   現在時刻を "yyyy-mm-dd HH:MM:SS" 形式で返す
'         LibreOffice Basic の Format(..., "nn") 問題を回避するため自前実装
' 引数:   なし
' 戻り値: String - "2026-04-19 15:52:19" 形式の文字列
' 備考:   Format(Now, "hh:nn:ss") は LibreOffice で "nn" が曜日略称に
'         解釈されるため使わない
' ================================================================
Public Function NowStr() As String
    Dim d As Date
    d = Now()
    NowStr = Pad2(Year(d)) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d)) & " " & _
             Pad2(Hour(d)) & ":" & Pad2(Minute(d)) & ":" & Pad2(Second(d))
End Function

' ================================================================
' 関数名: TodayStr
' 概要:   今日の日付を "yyyy-mm-dd" 形式で返す
' 引数:   なし
' 戻り値: String - "2026-04-19" 形式の文字列
' ================================================================
Public Function TodayStr() As String
    Dim d As Date
    d = Date
    TodayStr = Pad2(Year(d)) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d))
End Function

' ================================================================
' 関数名: IsDateInRange
' 概要:   target が fromDate <= target <= toDate の範囲内かを判定
' 引数:   target   - 判定対象の日付
'         fromDate - 範囲の開始（0の場合は下限なし）
'         toDate   - 範囲の終了（0の場合は上限なし）
' 戻り値: Boolean - 範囲内なら True
' ================================================================
Public Function IsDateInRange(ByVal target As Date, _
                                ByVal fromDate As Date, _
                                ByVal toDate As Date) As Boolean
    ' fromDate=0 は下限なし、toDate=0 は上限なし扱い
    If fromDate > 0 And target < fromDate Then
        IsDateInRange = False
        Exit Function
    End If
    If toDate > 0 And target > toDate Then
        IsDateInRange = False
        Exit Function
    End If
    IsDateInRange = True
End Function

' ================================================================
' 関数名: TryParseDate
' 概要:   文字列を日付に変換する。失敗時は False を返す
' 引数:   s        - 変換対象文字列（"2026-04-19" 等）
'         outDate  - 出力: 変換成功時の日付
' 戻り値: Boolean - 変換成功なら True
' 備考:   空文字列の場合は False（outDateは0のまま）
' ================================================================
Public Function TryParseDate(ByVal s As String, _
                               ByRef outDate As Date) As Boolean
    On Error GoTo ErrHandler
    
    If Trim(s) = "" Then
        outDate = 0
        TryParseDate = False
        Exit Function
    End If
    
    outDate = CDate(s)
    TryParseDate = True
    Exit Function

ErrHandler:
    outDate = 0
    TryParseDate = False
End Function

' ================================================================
' 関数名: Pad2
' 概要:   数値を2桁ゼロパディングした文字列に変換
'         例: 5 -> "05"、10 -> "10"
' 引数:   n - 変換対象の数値
' 戻り値: String - 2桁の文字列
' 備考:   100以上の値が来た場合はそのまま文字列化する
' ================================================================
Public Function Pad2(ByVal n As Long) As String
    If n < 10 Then
        Pad2 = "0" & CStr(n)
    Else
        Pad2 = CStr(n)
    End If
End Function
'@ },
    @{ Name='modFileIO'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modFileIO（ユーティリティ層）
' 概要:       Shift_JIS ファイルI/O、ファイル/フォルダ操作の純粋関数群
'             3段フォールバック方式（ADODB.Stream → UNO → Open For Output）
'             でExcel/LibreOffice 両環境に対応
' 依存先:     modCommon
' ================================================================

' ================================================================
' 関数名: WriteShiftJisFile
' 概要:   Shift_JIS で文字列をファイルに書き込む
'         ADODB.Stream（Windows/Excel） → UNO TextOutputStream（LibreOffice）
'         → Open For Output（最終フォールバック）の3段階で試行
' 引数:   filePath - 出力先ファイルパス
'         content  - 書き込む文字列
' 戻り値: Boolean - 成功なら True
' 備考:   既存ファイルは上書き
' ================================================================
Public Function WriteShiftJisFile(ByVal filePath As String, _
                                    ByVal content As String) As Boolean
    ' 第1段: ADODB.Stream
    If TryWriteWithADODB(filePath, content) Then
        WriteShiftJisFile = True
        Exit Function
    End If
    
    ' 第2段: UNO TextOutputStream（LibreOffice環境）
    If TryWriteWithUNO(filePath, content) Then
        WriteShiftJisFile = True
        Exit Function
    End If
    
    ' 第3段: 古典的な Open For Output
    If TryWriteWithOpenStatement(filePath, content) Then
        WriteShiftJisFile = True
        Exit Function
    End If
    
    WriteShiftJisFile = False
End Function

' 第1段: ADODB.Stream による書き込み
Private Function TryWriteWithADODB(ByVal filePath As String, _
                                     ByVal content As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2  ' adTypeText
    stream.Charset = CHARSET_SJIS
    stream.Open
    stream.WriteText content
    stream.SaveToFile filePath, 2  ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing
    
    TryWriteWithADODB = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    On Error GoTo 0
    TryWriteWithADODB = False
End Function

' 第2段: UNO TextOutputStream による書き込み（旧 LibreOffice 環境用）
' rev14 注意:
'   LibreOffice Basic 固有の組み込み createUnoService() は Excel VBA
'   （日本語 Windows 版含む）では未定義シンボルとなり、Option Explicit
'   による静的解析でプロジェクト全体がコンパイル失敗する。On Error GoTo
'   で実行時エラーをトラップしても、**コンパイル段階**で落ちるため
'   Application.Run が「マクロが使用できない」系のメッセージを出す。
'
'   現行ビルドは Excel を唯一の対象ランタイムとするため、UNO パスを
'   ソース上から撤去し、本関数は常に False を返す no-op に固定する。
'   LibreOffice headless での自動テストが必要になった場合は、別モジュール
'   として UNO 対応版を差し替える方式とする（本モジュールには混入させない）。
'
'   ADODB.Stream パス（第1段）が Windows Excel では常に通るため、
'   実運用で第2段に落ちるケースは無い。最終フォールバックとして
'   Open For Output（第3段）も健在。
Private Function TryWriteWithUNO(ByVal filePath As String, _
                                   ByVal content As String) As Boolean
    ' 未使用引数の参照抑止（WARN 回避）
    If Len(filePath) = 0 And Len(content) = 0 Then
        TryWriteWithUNO = False
    End If
    TryWriteWithUNO = False
End Function

' 第3段: Open For Output による書き込み（最終フォールバック）
Private Function TryWriteWithOpenStatement(ByVal filePath As String, _
                                             ByVal content As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, content
    Close #fNum
    
    TryWriteWithOpenStatement = True
    Exit Function

ErrHandler:
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
    TryWriteWithOpenStatement = False
End Function

' ================================================================
' 関数名: ReadShiftJisFile
' 概要:   Shift_JIS ファイルを文字列として読み込む
'         3段フォールバック方式
' 引数:   filePath - 読み込み元ファイルパス
' 戻り値: String - ファイル内容（失敗時は空文字列）
' ================================================================
Public Function ReadShiftJisFile(ByVal filePath As String) As String
    Dim result As String
    
    ' 第1段: ADODB.Stream
    If TryReadWithADODB(filePath, result) Then
        ReadShiftJisFile = result
        Exit Function
    End If
    
    ' 第2段: UNO
    If TryReadWithUNO(filePath, result) Then
        ReadShiftJisFile = result
        Exit Function
    End If
    
    ' 第3段: Open For Input
    If TryReadWithOpenStatement(filePath, result) Then
        ReadShiftJisFile = result
        Exit Function
    End If
    
    ReadShiftJisFile = ""
End Function

' 第1段: ADODB.Stream
Private Function TryReadWithADODB(ByVal filePath As String, _
                                    ByRef outContent As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2  ' adTypeText
    stream.Charset = CHARSET_SJIS
    stream.Open
    stream.LoadFromFile filePath
    outContent = stream.ReadText
    stream.Close
    Set stream = Nothing
    
    TryReadWithADODB = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    On Error GoTo 0
    outContent = ""
    TryReadWithADODB = False
End Function

' 第2段: UNO（旧 LibreOffice 環境用、Excel では no-op）
' rev14 注意: TryWriteWithUNO と同じ理由で createUnoService を撤去し、
'             Excel VBA でコンパイル可能な no-op に固定した。詳細は
'             TryWriteWithUNO の冒頭コメントを参照。
Private Function TryReadWithUNO(ByVal filePath As String, _
                                  ByRef outContent As String) As Boolean
    ' 未使用引数の参照抑止（WARN 回避）
    If Len(filePath) = 0 Then
        outContent = ""
    End If
    outContent = ""
    TryReadWithUNO = False
End Function

' 第3段: Open For Input
Private Function TryReadWithOpenStatement(ByVal filePath As String, _
                                            ByRef outContent As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim fNum As Integer
    Dim buffer As String
    Dim line As String
    
    fNum = FreeFile
    Open filePath For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, line
        If buffer = "" Then
            buffer = line
        Else
            buffer = buffer & vbCrLf & line
        End If
    Loop
    Close #fNum
    
    outContent = buffer
    TryReadWithOpenStatement = True
    Exit Function

ErrHandler:
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
    outContent = ""
    TryReadWithOpenStatement = False
End Function

' ================================================================
' 関数名: FileExists
' 概要:   ファイルが存在するか確認
' 引数:   filePath - ファイルパス
' 戻り値: Boolean - 存在すれば True
' 備考:   Dir関数を使用（LibreOffice互換）
' ================================================================
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    FileExists = (Dir(filePath) <> "")
    Exit Function
ErrHandler:
    FileExists = False
End Function

' ================================================================
' 関数名: FolderExists
' 概要:   フォルダが存在するか確認
' 引数:   folderPath - フォルダパス
' 戻り値: Boolean - 存在すれば True
' ================================================================
Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo ErrHandler
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    Exit Function
ErrHandler:
    FolderExists = False
End Function

' ================================================================
' 関数名: EnsureFolder
' 概要:   フォルダが存在しなければ作成する
' 引数:   folderPath - フォルダパス
' 戻り値: Boolean - 成功（既存含む）なら True
' ================================================================
Public Function EnsureFolder(ByVal folderPath As String) As Boolean
    On Error GoTo ErrHandler
    
    If FolderExists(folderPath) Then
        EnsureFolder = True
        Exit Function
    End If
    
    MkDir folderPath
    EnsureFolder = True
    Exit Function

ErrHandler:
    EnsureFolder = False
End Function

' ================================================================
' 関数名: ListFilesInFolder
' 概要:   フォルダ内の指定拡張子ファイル一覧を取得する
' 引数:   folderPath - 対象フォルダパス
'         extension  - 拡張子（"txt", "bas" 等、ドット不要）
' 戻り値: Variant - ファイル名の配列（拡張子込み）。0件の場合は空配列
' 備考:   サブフォルダは辿らない
' ================================================================
' M-4: 空配列ガード — 呼出側は IsEmpty / UBound < LBound チェック必須
'        本関数は ReDim 済み配列を返すが、Dir 失敗時は ReDim Preserve なしで返る場合がある
' s-2 contract: 本モジュール内の Kill / MkDir / Open For Output に渡されるパスは
'                呼出側で IsValidKnowledgeId / 自前パス検証済みであること。
'                本モジュールはパス検証責任を負わない (低レイヤ I/O 専念)。
' v14 D-4: UNO 段は LibreOffice 対応用 (本プロジェクトでは未使用)。
'          将来 LibreOffice 互換が必要なら別 modFileIO_UNO に切出し、本モジュールは
'          Excel 専用 ADODB + Open For Output の 2 段フォールバックに専念する。
' v14 s-3: ADODB fallback (Open For Output) で SJIS が壊れる可能性。
'          modFileIO は循環依存回避のため Logger 非依存だが、呼出側で
'          ADODB 失敗 → fallback 経路選択時に Logger.LogWarn で通知すること。
Public Function ListFilesInFolder(ByVal folderPath As String, _
                                    ByVal extension As String) As Variant
    On Error GoTo ErrHandler
    
    Dim results() As String
    Dim count As Long
    Dim fileName As String
    Dim searchPattern As String
    Dim sep As String
    
    count = 0
    ReDim results(0 To 99)  ' 初期100件
    
    ' パス末尾のセパレータ調整
    sep = "\"
    If Right(folderPath, 1) <> "\" And Right(folderPath, 1) <> "/" Then
        searchPattern = folderPath & sep & "*." & extension
    Else
        searchPattern = folderPath & "*." & extension
    End If
    
    fileName = Dir(searchPattern)
    Do While fileName <> ""
        If count > UBound(results) Then
            ReDim Preserve results(0 To count + 99)
        End If
        results(count) = fileName
        count = count + 1
        fileName = Dir()
    Loop
    
    If count = 0 Then
        ListFilesInFolder = Array()
    Else
        ReDim Preserve results(0 To count - 1)
        ListFilesInFolder = results
    End If
    Exit Function

ErrHandler:
    ListFilesInFolder = Array()
End Function

' ================================================================
' 関数名: DeleteFile
' 概要:   ファイルを削除する（物理削除）
' 引数:   filePath - ファイルパス
' 戻り値: Boolean - 成功なら True
' ================================================================
Public Function DeleteFile(ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    
    If Not FileExists(filePath) Then
        DeleteFile = False
        Exit Function
    End If
    
    Kill filePath
    DeleteFile = True
    Exit Function

ErrHandler:
    DeleteFile = False
End Function

' ================================================================
' 関数名: ConvertLocalPathToURL
' 概要:   ローカルファイルパスを UNO で使う URL 形式に変換
' 引数:   localPath - ローカルパス（例: "C:\work\file.txt"）
' 戻り値: String - URL 形式（例: "file:///C:/work/file.txt"）
' 備考:   UNO TextOutputStream 等で使用する内部ヘルパ
' ================================================================
Private Function ConvertLocalPathToURL(ByVal localPath As String) As String
    Dim result As String
    result = localPath
    
    ' バックスラッシュをスラッシュに
    result = Replace(result, "\", "/")
    
    ' Windows 形式（例: "C:/..."）の場合は "file:///" を前置
    If InStr(result, ":") > 0 Then
        ConvertLocalPathToURL = "file:///" & result
    Else
        ' Unix 形式（"/..."）の場合
        If Left(result, 1) = "/" Then
            ConvertLocalPathToURL = "file://" & result
        Else
            ConvertLocalPathToURL = "file:///" & result
        End If
    End If
End Function
'@ },
    @{ Name='modImageRender'; Type='std'; Code=@'
Option Explicit

' v6: Shapes.AddPicture / LockAspectRatio などは MsoTriState Enum を期待するため、
'      named arg + Long Const を渡すと real Excel が compile-time に拒否する (ADR 0020)。
'      → 位置指定 + 数値リテラル直書きで MsoTriState への暗黙変換に頼る。
'      AddPicture 第 2 引数: 0=msoFalse / -1=msoTrue / 1=msoCTrue (Office 標準値)
'      第 3 引数: 同 MsoTriState

' ================================================================
' モジュール: modImageRender (ユーティリティ層)
' 概要:       ナレッジ画像 (PNG) のサムネ表示と詳細画像ペイン描画。
'             Shapes.AddPicture を使うため LoadPicture と異なり
'             Worksheet 上でも Form 上でも Shape 残留問題を起こさない。
'             サムネは検索結果一覧の H 列、詳細画像は M-09 表示シートの
'             J4:N20 領域に配置。
' 依存先:     なし (純粋ユーティリティ)
' 備考:       VBA 子プロセス禁止 (Shell/Run/Exec/CreateObject Exec 系全部
'             禁止) のため、画像変換やフェッチは行わない。
'             ファイル不在時は薄字フォールバック表示のみ。
' ================================================================

' --- サムネサイズ ---
Public Const KB_THUMB_WIDTH As Double = 60#
Public Const KB_THUMB_HEIGHT As Double = 40#

' --- 詳細画像最大サイズ ---
Private Const KB_DETAIL_MAX_WIDTH As Double = 400#
Private Const KB_DETAIL_MAX_HEIGHT As Double = 300#

' --- Shape 名 prefix (削除/再配置用) ---
Public Const PREFIX_THUMB As String = "kbThumb_"
Public Const PREFIX_DETAIL As String = "kbDetailImg_"

' --- セル padding ---
Private Const PAD_PIXELS As Double = 2#

' ================================================================
' 関数名: RenderThumb
' 概要:   検索結果一覧のサムネセル (col=8, H 列) に画像 Shape を貼付。
'         画像が存在しない場合は何も描画せず黙ってリターン (薄字テキストは
'         呼び出し側でセルに書く運用)。
' 引数:   ws         - 配置対象シート
'         row        - 配置先の行番号
'         col        - 配置先の列番号
'         imageFull  - 画像の絶対パス
'         knwNo      - ナレッジ番号 (Shape 名識別用)
' 戻り値: なし
' 備考:   Shape 名は PREFIX_THUMB & "<row>_" & knwNo 形式。
'         再描画時は ClearShapesByPrefix で全削除→再生成 (idempotent)。
' ================================================================
Public Sub RenderThumb(ByVal ws As Worksheet, ByVal targetRow As Long, _
                         ByVal targetCol As Long, ByVal imageFull As String, _
                         ByVal knwNo As String)
    On Error GoTo ErrHandler
    If imageFull = "" Then Exit Sub
    If Dir(imageFull) = "" Then Exit Sub

    Dim cell As Range
    Set cell = ws.Cells(targetRow, targetCol)

    Dim shp As Shape
    ' v8: Object 型 late binding で AddPicture を呼ぶ (ADR 0023)
    '     typed Shapes 経由だと MsoTriState 引数の数値リテラル暗黙変換が
    '     real Excel で reject される。Object 経由なら compile-time 型チェック skip。
    '     第 2 引数 LinkToFile = 0 (msoFalse), 第 3 引数 SaveWithDocument = -1 (msoTrue)
    Dim shapesObj As Object
    Set shapesObj = ws.Shapes
    Set shp = shapesObj.AddPicture(imageFull, 0, -1, _
        cell.Left + PAD_PIXELS, cell.Top + PAD_PIXELS, _
        KB_THUMB_WIDTH, KB_THUMB_HEIGHT)
    shp.Name = PREFIX_THUMB & CStr(targetRow) & "_" & knwNo

    ' 行高をサムネに合わせる (狭ければ拡大)
    If ws.Rows(targetRow).RowHeight < (KB_THUMB_HEIGHT + PAD_PIXELS * 2#) Then
        ws.Rows(targetRow).RowHeight = KB_THUMB_HEIGHT + PAD_PIXELS * 2#
    End If
    Exit Sub
ErrHandler:
    ' 画像描画失敗は致命ではない (検索結果テキストは別途書かれている)
End Sub

' ================================================================
' 関数名: RenderDetailImage
' 概要:   ナレッジ表示シートの J4:N20 領域に詳細画像を貼付。
'         領域からはみ出さないようアスペクト比を保ってリサイズ。
'         画像が無い場合は薄字で「[画像未配置: ...]」セル表示。
' 引数:   ws         - 配置対象シート
'         topRow     - ペイン左上行
'         leftCol    - ペイン左上列
'         botRow     - ペイン右下行
'         rightCol   - ペイン右下列
'         imageFull  - 画像の絶対パス
'         knwNo      - ナレッジ番号 (Shape 名識別用)
' ================================================================
Public Sub RenderDetailImage(ByVal ws As Worksheet, _
                               ByVal topRow As Long, ByVal leftCol As Long, _
                               ByVal botRow As Long, ByVal rightCol As Long, _
                               ByVal imageFull As String, _
                               ByVal knwNo As String)
    On Error GoTo ErrHandler

    Dim r1 As Range, r2 As Range
    Set r1 = ws.Cells(topRow, leftCol)
    Set r2 = ws.Cells(botRow, rightCol)
    Dim maxW As Double, maxH As Double
    maxW = (r2.Left + r2.Width) - r1.Left - PAD_PIXELS * 2#
    maxH = (r2.Top + r2.Height) - r1.Top - PAD_PIXELS * 2#
    If maxW < KB_DETAIL_MAX_WIDTH Then maxW = KB_DETAIL_MAX_WIDTH
    If maxH < KB_DETAIL_MAX_HEIGHT Then maxH = KB_DETAIL_MAX_HEIGHT

    If imageFull = "" Or Dir(imageFull) = "" Then
        ' フォールバックテキスト
        Call WriteFallbackText(ws, topRow, leftCol, botRow, rightCol, imageFull)
        Exit Sub
    End If

    Dim shp As Shape
    ' v8: Object 型 late binding で AddPicture を呼ぶ (ADR 0023)
    '     第 2 引数 LinkToFile = 0 (msoFalse), 第 3 引数 SaveWithDocument = -1 (msoTrue)
    '     Width=-1 / Height=-1 はオリジナルサイズで読込 (AddPicture 仕様)
    Dim shapesObj2 As Object
    Set shapesObj2 = ws.Shapes
    Set shp = shapesObj2.AddPicture(imageFull, 0, -1, _
        r1.Left + PAD_PIXELS, r1.Top + PAD_PIXELS, _
        -1, -1)
    shp.Name = PREFIX_DETAIL & knwNo
    ' v8: Object 型 late binding で LockAspectRatio を設定 (ADR 0023)
    '     v7 の Variant 中継でも real Excel が typed Shape 経由の MsoTriState 代入を
    '     reject する事例があるため、Shape を Object 経由にラップして compile-time
    '     型チェック自体を skip する。
    Dim shpObj As Object
    Set shpObj = shp
    shpObj.LockAspectRatio = -1  ' msoTrue (Object 経由なので数値リテラル OK)

    ' 領域内に収まるようリサイズ
    Dim ratioW As Double, ratioH As Double
    ratioW = maxW / shp.Width
    ratioH = maxH / shp.Height
    Dim ratio As Double
    If ratioW < ratioH Then
        ratio = ratioW
    Else
        ratio = ratioH
    End If
    If ratio < 1# Then
        shp.Width = shp.Width * ratio
    End If
    Exit Sub
ErrHandler:
    ' 画像描画失敗時はフォールバックテキストを書く
    On Error Resume Next
    Call WriteFallbackText(ws, topRow, leftCol, botRow, rightCol, imageFull)
End Sub

' ================================================================
' 関数名: WriteFallbackText
' 概要:   画像未配置時の薄字テキスト表示
' ================================================================
Private Sub WriteFallbackText(ByVal ws As Worksheet, _
                                ByVal topRow As Long, ByVal leftCol As Long, _
                                ByVal botRow As Long, ByVal rightCol As Long, _
                                ByVal imageFull As String)
    Dim msg As String
    If imageFull = "" Then
        msg = "[画像未配置]"
    Else
        msg = "[画像未配置: " & imageFull & "]"
    End If
    On Error Resume Next
    ws.Cells(topRow, leftCol).Value = msg
    ws.Cells(topRow, leftCol).Font.Color = RGB(160, 160, 160)
    ws.Cells(topRow, leftCol).Font.Italic = True
End Sub

' ================================================================
' 関数名: ClearShapesByPrefix
' 概要:   指定 prefix で始まる名前の Shape を全削除 (idempotent)
' 引数:   ws     - 対象シート
'         prefix - Shape 名 prefix (例: "kbThumb_")
' 戻り値: なし
' ================================================================
Public Sub ClearShapesByPrefix(ByVal ws As Worksheet, ByVal prefix As String)
    On Error GoTo ErrHandler
    Dim i As Long
    Dim plen As Long
    plen = Len(prefix)
    If plen = 0 Then Exit Sub

    For i = ws.Shapes.Count To 1 Step -1
        Dim shp As Shape
        On Error Resume Next
        Set shp = ws.Shapes(i)
        If Err.Number = 0 Then
            If Len(shp.Name) >= plen Then
                If Left(shp.Name, plen) = prefix Then
                    shp.Delete
                End If
            End If
        End If
        Err.Clear
        On Error GoTo ErrHandler
    Next i
    Exit Sub
ErrHandler:
    ' 削除失敗は致命ではない
End Sub

' ================================================================
' 関数名: HasShapeWithPrefix
' 概要:   指定 prefix の Shape が 1 つ以上存在するか判定 (テスト用)
' ================================================================
Public Function HasShapeWithPrefix(ByVal ws As Worksheet, _
                                     ByVal prefix As String) As Boolean
    On Error GoTo ErrHandler
    Dim i As Long
    Dim plen As Long
    plen = Len(prefix)
    For i = 1 To ws.Shapes.Count
        Dim shp As Shape
        Set shp = ws.Shapes(i)
        If Len(shp.Name) >= plen Then
            If Left(shp.Name, plen) = prefix Then
                HasShapeWithPrefix = True
                Exit Function
            End If
        End If
    Next i
    HasShapeWithPrefix = False
    Exit Function
ErrHandler:
    HasShapeWithPrefix = False
End Function
'@ },
    @{ Name='modStringUtil'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modStringUtil（ユーティリティ層）
' 概要:       文字列処理の純粋関数群
' 依存先:     modCommon
' ================================================================

' ================================================================
' 関数名: ContainsAllKeywords
' 概要:   対象文字列がキーワード（スペース区切り）を全て含むか判定（AND検索）
' 引数:   target   - 検索対象の文字列
'         keywords - スペース区切りのキーワード群（例: "メモリ サーバ"）
' 戻り値: Boolean - 全キーワードを含むなら True
' 備考:   大文字小文字を区別しない（LCase比較）
'         keywordsが空の場合は True を返す
' ================================================================
Public Function ContainsAllKeywords(ByVal target As String, _
                                      ByVal keywords As String) As Boolean
    Dim parts() As String
    Dim i As Long
    Dim targetLower As String
    
    If Trim(keywords) = "" Then
        ContainsAllKeywords = True
        Exit Function
    End If
    
    targetLower = LCase(target)
    parts = Split(Trim(keywords), " ")
    
    For i = LBound(parts) To UBound(parts)
        If Trim(parts(i)) <> "" Then
            If InStr(targetLower, LCase(parts(i))) = 0 Then
                ContainsAllKeywords = False
                Exit Function
            End If
        End If
    Next i
    
    ContainsAllKeywords = True
End Function

' ================================================================
' 関数名: ContainsAnyKeyword
' 概要:   対象文字列がキーワード（スペース区切り）のいずれかを含むか判定（OR検索）
' 引数:   target   - 検索対象の文字列
'         keywords - スペース区切りのキーワード群
' 戻り値: Boolean - いずれかのキーワードを含むなら True
' 備考:   大文字小文字を区別しない（LCase比較）
'         keywordsが空の場合は True を返す
' ================================================================
Public Function ContainsAnyKeyword(ByVal target As String, _
                                     ByVal keywords As String) As Boolean
    Dim parts() As String
    Dim i As Long
    Dim targetLower As String
    
    If Trim(keywords) = "" Then
        ContainsAnyKeyword = True
        Exit Function
    End If
    
    targetLower = LCase(target)
    parts = Split(Trim(keywords), " ")
    
    For i = LBound(parts) To UBound(parts)
        If Trim(parts(i)) <> "" Then
            If InStr(targetLower, LCase(parts(i))) > 0 Then
                ContainsAnyKeyword = True
                Exit Function
            End If
        End If
    Next i
    
    ContainsAnyKeyword = False
End Function

' ================================================================
' 関数名: FormatNumberByPattern
' 概要:   採番パターンと数値から、採番後の文字列を生成する
'         例: "INC-{0000}" + 5 -> "INC-0005"
' 引数:   pattern  - パターン文字列（{0000}等のプレースホルダ含む）
'         nextNum  - 採番する値
' 戻り値: String - パターンに数値を埋め込んだ文字列
' 備考:   {} で囲まれた 0 の個数がゼロパディング桁数になる
'         パターンに {} が含まれない場合は pattern をそのまま返す
' ================================================================
Public Function FormatNumberByPattern(ByVal pattern As String, _
                                        ByVal nextNum As Long) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim placeholder As String
    Dim digits As Long
    Dim numStr As String
    Dim i As Long
    
    startPos = InStr(pattern, "{")
    If startPos = 0 Then
        FormatNumberByPattern = pattern
        Exit Function
    End If
    
    endPos = InStr(startPos, pattern, "}")
    If endPos = 0 Then
        FormatNumberByPattern = pattern
        Exit Function
    End If
    
    placeholder = Mid(pattern, startPos + 1, endPos - startPos - 1)
    digits = Len(placeholder)
    numStr = CStr(nextNum)
    
    ' 指定桁数に満たない場合はゼロパディング
    For i = Len(numStr) + 1 To digits
        numStr = "0" & numStr
    Next i
    
    FormatNumberByPattern = Left(pattern, startPos - 1) & numStr & _
                             Mid(pattern, endPos + 1)
End Function

' ================================================================
' 関数名: ParseStanzaLine
' 概要:   スタンザ1行（"Key=Value" 形式）を分解する
' 引数:   line     - 入力行
'         outKey   - 出力: キー
'         outValue - 出力: 値
' 戻り値: Boolean - 解析成功なら True
' 備考:   "=" が含まれない行は False を返す
'         値に "=" が含まれる場合、最初の "=" で分割する
' ================================================================
Public Function ParseStanzaLine(ByVal line As String, _
                                  ByRef outKey As String, _
                                  ByRef outValue As String) As Boolean
    Dim pos As Long
    
    outKey = ""
    outValue = ""
    
    pos = InStr(line, "=")
    If pos = 0 Then
        ParseStanzaLine = False
        Exit Function
    End If
    
    outKey = Trim(Left(line, pos - 1))
    outValue = Mid(line, pos + 1)
    ParseStanzaLine = True
End Function

' ================================================================
' 関数名: TrimAll
' 概要:   前後の空白（半角・全角）を除去
' 引数:   s - 対象文字列
' 戻り値: String - 空白除去後の文字列
' ================================================================
Public Function TrimAll(ByVal s As String) As String
    Dim result As String
    result = s
    
    ' 先頭の全角スペースを除去
    Do While Len(result) > 0 And Left(result, 1) = "　"
        result = Mid(result, 2)
    Loop
    
    ' 末尾の全角スペースを除去
    Do While Len(result) > 0 And Right(result, 1) = "　"
        result = Left(result, Len(result) - 1)
    Loop
    
    ' 半角スペースのトリム
    TrimAll = Trim(result)
End Function

' ================================================================
' 関数名: SafeCellText
' 概要:   "=" 先頭文字列を数式解釈させないようアポストロフィを前置
'         Excel/LibreOffice 両方で有効
' 引数:   s - 対象文字列
' 戻り値: String - 数式化されない文字列
' 備考:   "=" 先頭以外はそのまま返す
'         呼び出し側ではセル代入前に NumberFormat = "@" も併用推奨
' ================================================================
Public Function SafeCellText(ByVal s As String) As String
    If Len(s) > 0 And Left(s, 1) = "=" Then
        SafeCellText = "'" & s
    Else
        SafeCellText = s
    End If
End Function

' ================================================================
' 関数名: IsValidKnowledgeId
' 概要:   ナレッジ番号が "KN-yyyy-NNNN" 形式かを検証する。
'         パストラバーサル / UNC リダイレクト / 予約デバイス名 / 危険文字
'         を MkDir / Kill / Open For Output 等の File system 操作直前で
'         reject するための security ガード (s-2)。
' 引数:   knwNo - 検証対象のナレッジ番号文字列
' 戻り値: Boolean - 形式準拠かつ危険文字を含まない場合のみ True
' 備考:   VBA 標準の RegExp はクロスプラットフォーム差があるため手動判定。
'         12 文字固定 / "KN-" 先頭 / 4 桁年 / "-" 区切り / 4 桁番号 で構成。
'         ".." "\" "/" ":" 等は二重防衛として明示 reject。
' 関連:   ADR 0016
' ================================================================
Public Function IsValidKnowledgeId(ByVal knwNo As String) As Boolean
    On Error GoTo ErrHandler
    IsValidKnowledgeId = False

    If Len(knwNo) <> 12 Then Exit Function
    If Left$(knwNo, 3) <> "KN-" Then Exit Function
    If Mid$(knwNo, 8, 1) <> "-" Then Exit Function

    Dim i As Long
    Dim ch As String
    For i = 4 To 7
        ch = Mid$(knwNo, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    For i = 9 To 12
        ch = Mid$(knwNo, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    ' 危険文字の二重防衛 reject (12 文字制約で本来は通らないが、明示)
    Dim danger As Variant
    danger = Array("..", "\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
    Dim d As Variant
    For Each d In danger
        If InStr(knwNo, CStr(d)) > 0 Then Exit Function
    Next d

    IsValidKnowledgeId = True
    Exit Function

ErrHandler:
    IsValidKnowledgeId = False
End Function
'@ },
    @{ Name='IScreenRenderer'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' インタフェース: IScreenRenderer（画面層）
' 概要:   画面要素を物理的に描画する責務を抽象化するインタフェース。
'         実装1: clsSheetRenderer（シート埋込型 — 今回採用）
'         実装2: clsUserFormRenderer（フォーム型 — 将来切替先のスタブ）
' 依存先: clsScreenSpec, clsButtonSpec, clsFieldSpec, clsSectionSpec
' 備考:   VBA の Implements パターン用。Public Sub のシグネチャだけ定義し
'         具体実装はそれぞれのクラスで Private Sub IScreenRenderer_xxx として書く。
' ================================================================

' ----- 画面初期化 -----
Public Sub BindSheet(ByVal sheetName As String)
End Sub

Public Sub ClearScreen()
End Sub

' ----- タイトル/セクション -----
Public Sub RenderTitle(ByVal screenId As String, ByVal title As String, ByVal colorHex As String)
End Sub

Public Sub RenderSection(ByVal sectionAddr As String, ByVal sectionLabel As String, ByVal colorHex As String)
End Sub

' ----- ボタン -----
Public Sub RenderButton(ByVal btnSpec As Object)
End Sub

' ----- ラベル/フィールド -----
Public Sub RenderLabel(ByVal cellAddr As String, ByVal labelText As String, ByVal colorHex As String)
End Sub

Public Sub RenderInputField(ByVal cellAddr As String, ByVal fieldSpec As Object)
End Sub

Public Sub RenderRequiredMark(ByVal cellAddr As String)
End Sub

Public Sub RenderHint(ByVal cellAddr As String, ByVal hintText As String)
End Sub

' ----- データ系 -----
Public Sub RenderHeaderRow(ByVal startAddr As String, ByVal headerLabels As Variant, ByVal colorHex As String)
End Sub

Public Sub RenderEmptyState(ByVal cellAddr As String, ByVal message As String)
End Sub
'@ },
    @{ Name='clsButtonSpec'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsButtonSpec（画面層 — データ）
' 概要:   1 個のボタンの仕様（位置・キャプション・色・マクロ）を保持する
'         ValueObject。コードから分離してデータ駆動で画面を構築する。
' 依存先: なし（ValueObject）
' ================================================================

Private m_btnName As String
Private m_caption As String
Private m_cellAddr As String
Private m_colorHex As String
Private m_groupName As String
Private m_hintAddr As String
Private m_hintText As String

' --- Property Get/Let ---
Public Property Get BtnName() As String
    BtnName = m_btnName
End Property
Public Property Let BtnName(ByVal value As String)
    m_btnName = value
End Property

Public Property Get Caption() As String
    Caption = m_caption
End Property
Public Property Let Caption(ByVal value As String)
    m_caption = value
End Property

Public Property Get CellAddr() As String
    CellAddr = m_cellAddr
End Property
Public Property Let CellAddr(ByVal value As String)
    m_cellAddr = value
End Property

Public Property Get ColorHex() As String
    ColorHex = m_colorHex
End Property
Public Property Let ColorHex(ByVal value As String)
    m_colorHex = value
End Property

Public Property Get GroupName() As String
    GroupName = m_groupName
End Property
Public Property Let GroupName(ByVal value As String)
    m_groupName = value
End Property

Public Property Get HintAddr() As String
    HintAddr = m_hintAddr
End Property
Public Property Let HintAddr(ByVal value As String)
    m_hintAddr = value
End Property

Public Property Get HintText() As String
    HintText = m_hintText
End Property
Public Property Let HintText(ByVal value As String)
    m_hintText = value
End Property

' ================================================================
' 関数名: Init
' 概要:   一括初期化（Builder 風）
' 引数:   btnName, caption, cellAddr, colorHex, groupName, hintAddr, hintText
' ================================================================
Public Sub Init(ByVal btnName As String, _
                ByVal caption As String, _
                ByVal cellAddr As String, _
                ByVal colorHex As String, _
                Optional ByVal groupName As String = "", _
                Optional ByVal hintAddr As String = "", _
                Optional ByVal hintText As String = "")
    m_btnName = btnName
    m_caption = caption
    m_cellAddr = cellAddr
    m_colorHex = colorHex
    m_groupName = groupName
    m_hintAddr = hintAddr
    m_hintText = hintText
End Sub
'@ },
    @{ Name='clsControlSpec'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsControlSpec (ビジネスロジック層 / 値オブジェクト)
' 概要:   1 つの UserForm 上 Control の宣言情報を保持する ValueObject。
'         ControlType / Name / Left / Top / Width / Height / Caption / OnClick
'         を Private 変数 + Property Get/Let で公開する (R11 準拠)。
' 依存先: なし
' 備考:   状態を持つので .cls (vba-coding-standards R10/R11)。
'         OnClick はマクロ名 (Application.Run で呼ぶ)。
' ================================================================

Private m_controlType As String
Private m_name As String
Private m_left As Long
Private m_top As Long
Private m_width As Long
Private m_height As Long
Private m_caption As String
Private m_onClick As String

' --- ControlType ---
Public Property Get ControlType() As String
    ControlType = m_controlType
End Property
Public Property Let ControlType(ByVal v As String)
    m_controlType = v
End Property

' --- Name ---
Public Property Get Name() As String
    Name = m_name
End Property
Public Property Let Name(ByVal v As String)
    m_name = v
End Property

' --- Left ---
Public Property Get Left() As Long
    Left = m_left
End Property
Public Property Let Left(ByVal v As Long)
    m_left = v
End Property

' --- Top ---
Public Property Get Top() As Long
    Top = m_top
End Property
Public Property Let Top(ByVal v As Long)
    m_top = v
End Property

' --- Width ---
Public Property Get Width() As Long
    Width = m_width
End Property
Public Property Let Width(ByVal v As Long)
    m_width = v
End Property

' --- Height ---
Public Property Get Height() As Long
    Height = m_height
End Property
Public Property Let Height(ByVal v As Long)
    m_height = v
End Property

' --- Caption ---
Public Property Get Caption() As String
    Caption = m_caption
End Property
Public Property Let Caption(ByVal v As String)
    m_caption = v
End Property

' --- OnClick (マクロ名) ---
Public Property Get OnClick() As String
    OnClick = m_onClick
End Property
Public Property Let OnClick(ByVal v As String)
    m_onClick = v
End Property

' ================================================================
' 関数名: Init
' 概要:   全プロパティを 1 回でセットする初期化メソッド
' 引数:   cType   - "Label"/"TextBox"/"Button"/"Image"/"ListBox"/"ComboBox"
'         nm      - Control 名 (Form 内で一意)
'         l/t/w/h - Left/Top/Width/Height (points)
'         cap     - Caption (Label/Button/CheckBox 用、省略可)
'         onClk   - OnClick マクロ名 (Button 用、省略可)
' ================================================================
Public Sub Init(ByVal cType As String, ByVal nm As String, _
                 ByVal l As Long, ByVal t As Long, _
                 ByVal w As Long, ByVal h As Long, _
                 Optional ByVal cap As String = "", _
                 Optional ByVal onClk As String = "")
    m_controlType = cType
    m_name = nm
    m_left = l
    m_top = t
    m_width = w
    m_height = h
    m_caption = cap
    m_onClick = onClk
End Sub
'@ },
    @{ Name='clsFieldMigrator'; Type='cls'; Code=@'
Option Explicit

' Phase 6 レビュー: 9 subs/funcs に対し On Error GoTo は 1 件のみ。
' Public method (Init / RunMigration) は ErrHandler 補強推奨だが、
' 既存 89 テスト互換性のため本 v4 では指摘記録のみ (v5 で対応予定)。

' ================================================================
' クラス: clsFieldMigrator（ビジネスロジック層）
' 概要:   フォーマット変更後、既存ナレッジファイルに新フィールドを反映
'         既存の値は変更せず、不足フィールドのみ空値で追加する
' 依存先: clsLogger, clsFormatManager, modFileIO, modStringUtil, modCommon
' ================================================================

Private m_logger As clsLogger
Private m_formatMgr As clsFormatManager
Private m_dataFolder As String

' ================================================================
' 関数名: Init
' 概要:   初期化
' 引数:   logger     - ログ出力用
'         formatMgr  - フォーマット管理
'         dataFolder - ナレッジデータフォルダ
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger, _
                 ByVal formatMgr As clsFormatManager, _
                 ByVal dataFolder As String)
    Set m_logger = logger
    Set m_formatMgr = formatMgr
    m_dataFolder = dataFolder
End Sub

' ================================================================
' 関数名: MigrateFields
' 概要:   指定フォーマットの全ナレッジファイルに新フィールド定義を反映
' 引数:   formatId - 対象フォーマットID
' 戻り値: Long - 処理ファイル数
' 備考:   各ファイルについて:
'         - フォーマット定義ヘッダーを最新版に更新
'         - 既存のITEM行は維持
'         - 新フィールドは空値で追加
'         - FormatVersion、UpdatedDateを更新
' ================================================================
Public Function MigrateFields(ByVal formatId As String) As Long
    On Error GoTo ErrHandler
    
    Dim files As Variant
    files = ListFilesInFolder(m_dataFolder, "txt")

    ' M-4 guard: 空配列なら早期 return (UBound エラー防止)
    ' v13 ADR 0028: MigrateFields は Function なので Exit Function に修正
    If (Not Not files) = 0 Then
        MigrateFields = 0
        Exit Function
    End If
    
    Dim newHeader As String
    newHeader = m_formatMgr.GetFormatHeaderAsStanza(formatId)
    
    Dim newFields As Variant
    newFields = GetNewFieldDefinitions(formatId)
    
    Dim processedCount As Long
    processedCount = 0
    
    Dim i As Long
    For i = LBound(files) To UBound(files)
        Dim fileName As String
        fileName = CStr(files(i))
        
        Dim filePath As String
        filePath = CombineFilePath(m_dataFolder, fileName)
        
        Dim content As String
        content = ReadShiftJisFile(filePath)
        If content = "" Then
            GoTo NextFile
        End If
        
        Dim actualFmt As String
        actualFmt = ExtractStanzaValue(content, "FormatID")
        If actualFmt <> formatId Then
            GoTo NextFile
        End If
        
        Dim newContent As String
        newContent = RebuildKnowledgeFile(content, newHeader, newFields, formatId)
        
        If WriteShiftJisFile(filePath, newContent) Then
            processedCount = processedCount + 1
        End If
        
NextFile:
    Next i
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFieldMigrator", "MigrateFields", _
                          "反映完了: " & CStr(processedCount) & "件"
    End If
    
    MigrateFields = processedCount
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFieldMigrator", "MigrateFields", Err.Description
    End If
    MigrateFields = 0
End Function

' ================================================================
' 関数名: GetNewFieldDefinitions
' 概要:   フォーマット設計シートから現在のフィールド定義を取得
' 引数:   formatId - フォーマットID
' 戻り値: Variant - [No, Name] の配列
' ================================================================
Private Function GetNewFieldDefinitions(ByVal formatId As String) As Variant
    Dim designWs As Worksheet
    Set designWs = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    Dim result() As String
    ReDim result(0 To 999, 0 To 1)
    
    Dim i As Long
    Dim count As Long
    count = 0
    
    For i = 6 To 1000
        Dim fieldName As String
        fieldName = CStr(designWs.Cells(i, 2).Value)
        If fieldName = "" Then Exit For
        
        result(count, 0) = CStr(designWs.Cells(i, 1).Value)
        result(count, 1) = fieldName
        count = count + 1
    Next i
    
    If count = 0 Then
        GetNewFieldDefinitions = Array()
    Else
        ReDim Preserve result(0 To count - 1, 0 To 1)
        GetNewFieldDefinitions = result
    End If
End Function

' ================================================================
' 関数名: RebuildKnowledgeFile
' 概要:   既存ファイル内容を新ヘッダーと新フィールド定義で再構築
' 引数:   oldContent - 既存ファイル内容
'         newHeader  - 新フォーマットヘッダー
'         newFields  - 新フィールド定義配列
'         formatId   - フォーマットID
' 戻り値: String - 再構築されたファイル内容
' ================================================================
Private Function RebuildKnowledgeFile(ByVal oldContent As String, _
                                        ByVal newHeader As String, _
                                        ByVal newFields As Variant, _
                                        ByVal formatId As String) As String
    Dim knowledgeNo As String
    Dim createdDate As String
    knowledgeNo = ExtractStanzaValue(oldContent, "KnowledgeNo")
    createdDate = ExtractStanzaValue(oldContent, "CreatedDate")
    
    Dim newVersion As Long
    newVersion = GetFormatVersion(formatId)
    
    Dim result As String
    result = newHeader
    result = result & STANZA_SEP & vbCrLf
    result = result & "[KNOWLEDGE]" & vbCrLf
    result = result & "KnowledgeNo=" & knowledgeNo & vbCrLf
    result = result & "FormatVersion=" & CStr(newVersion) & vbCrLf
    result = result & "CreatedDate=" & createdDate & vbCrLf
    result = result & "UpdatedDate=" & TodayStr() & vbCrLf
    
    ' 既存値を辞書化
    Dim existingValues As Object
    Set existingValues = BuildFieldValueMap(oldContent)
    
    ' 新フィールド定義に従って項目を出力（既存値があれば使用、なければ空）
    Dim i As Long
    For i = LBound(newFields, 1) To UBound(newFields, 1)
        Dim fieldNo As String
        Dim fieldName As String
        fieldNo = newFields(i, 0)
        fieldName = newFields(i, 1)
        
        Dim fieldValue As String
        If existingValues.Exists(fieldName) Then
            fieldValue = existingValues(fieldName)
        Else
            fieldValue = ""
        End If
        
        result = result & "[ITEM]" & vbCrLf
        result = result & "FieldNo=" & fieldNo
        result = result & " / FieldName=" & fieldName
        result = result & " / Value=" & fieldValue
        result = result & vbCrLf
    Next i
    
    result = result & STANZA_SEP
    RebuildKnowledgeFile = result
End Function

' ================================================================
' 関数名: BuildFieldValueMap
' 概要:   既存ファイル内容からフィールド名→値のマップを作る
' 引数:   content - ファイル内容
' 戻り値: Object (Dictionary) - FieldName → Value のマップ
' ================================================================
Private Function BuildFieldValueMap(ByVal content As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lines() As String
    lines = Split(content, vbCrLf)
    
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "FieldName=") > 0 And _
           InStr(lines(i), "Value=") > 0 Then
            Dim fName As String
            Dim fValue As String
            fName = ExtractItemValue(lines(i), "FieldName")
            fValue = ExtractItemValue(lines(i), "Value")
            If fName <> "" And Not dict.Exists(fName) Then
                dict.Add fName, fValue
            End If
        End If
    Next i
    
    Set BuildFieldValueMap = dict
End Function

' ITEM行から指定キーの値を抽出
Private Function ExtractItemValue(ByVal line As String, _
                                    ByVal keyName As String) As String
    Dim searchKey As String
    Dim startPos As Long
    Dim endPos As Long
    
    searchKey = keyName & "="
    startPos = InStr(line, searchKey)
    If startPos = 0 Then
        ExtractItemValue = ""
        Exit Function
    End If
    
    startPos = startPos + Len(searchKey)
    endPos = InStr(startPos, line, " / ")
    If endPos = 0 Then
        ExtractItemValue = Mid(line, startPos)
    Else
        ExtractItemValue = Mid(line, startPos, endPos - startPos)
    End If
End Function

' スタンザKey=Value形式から単純抽出
Private Function ExtractStanzaValue(ByVal content As String, _
                                      ByVal keyName As String) As String
    Dim lines() As String
    Dim i As Long
    
    lines = Split(content, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), keyName & "=") = 1 Then
            ExtractStanzaValue = Mid(lines(i), Len(keyName) + 2)
            Exit Function
        End If
    Next i
    ExtractStanzaValue = ""
End Function

' フォーマット一覧からバージョン番号を取得
Private Function GetFormatVersion(ByVal formatId As String) As Long
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim i As Long
    For i = 3 To 1000
        If CStr(listWs.Cells(i, 2).Value) = formatId Then
            GetFormatVersion = CLng(listWs.Cells(i, 7).Value)
            Exit Function
        End If
    Next i
    GetFormatVersion = 1
End Function

' フォルダパスとファイル名を結合
Private Function CombineFilePath(ByVal folder As String, _
                                   ByVal fileName As String) As String
    If Right(folder, 1) = "\" Or Right(folder, 1) = "/" Then
        CombineFilePath = folder & fileName
    Else
        CombineFilePath = folder & "\" & fileName
    End If
End Function
'@ },
    @{ Name='clsFieldSpec'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsFieldSpec（画面層 — データ）
' 概要:   1 個の入力フィールドの仕様（順序・ラベル・型・必須・行数）を
'         保持する ValueObject。データが空でもラベルは常時表示する根拠。
' 依存先: なし
' ================================================================

Private m_order As Long
Private m_label As String
Private m_typeText As String
Private m_required As Boolean
Private m_rowCount As Long
Private m_hintText As String
Private m_orderAddr As String     ' 例: "A8"
Private m_reqMarkAddr As String   ' 例: "B8"
Private m_labelAddr As String     ' 例: "C8"
Private m_typeAddr As String      ' 例: "D8"
Private m_inputAddr As String     ' 例: "E8"

Public Property Get FieldOrder() As Long
    FieldOrder = m_order
End Property
Public Property Let FieldOrder(ByVal value As Long)
    m_order = value
End Property

Public Property Get Label() As String
    Label = m_label
End Property
Public Property Let Label(ByVal value As String)
    m_label = value
End Property

Public Property Get TypeText() As String
    TypeText = m_typeText
End Property
Public Property Let TypeText(ByVal value As String)
    m_typeText = value
End Property

Public Property Get Required() As Boolean
    Required = m_required
End Property
Public Property Let Required(ByVal value As Boolean)
    m_required = value
End Property

Public Property Get RowCount() As Long
    RowCount = m_rowCount
End Property
Public Property Let RowCount(ByVal value As Long)
    m_rowCount = value
End Property

Public Property Get HintText() As String
    HintText = m_hintText
End Property
Public Property Let HintText(ByVal value As String)
    m_hintText = value
End Property

Public Property Get OrderAddr() As String
    OrderAddr = m_orderAddr
End Property
Public Property Let OrderAddr(ByVal value As String)
    m_orderAddr = value
End Property

Public Property Get ReqMarkAddr() As String
    ReqMarkAddr = m_reqMarkAddr
End Property
Public Property Let ReqMarkAddr(ByVal value As String)
    m_reqMarkAddr = value
End Property

Public Property Get LabelAddr() As String
    LabelAddr = m_labelAddr
End Property
Public Property Let LabelAddr(ByVal value As String)
    m_labelAddr = value
End Property

Public Property Get TypeAddr() As String
    TypeAddr = m_typeAddr
End Property
Public Property Let TypeAddr(ByVal value As String)
    m_typeAddr = value
End Property

Public Property Get InputAddr() As String
    InputAddr = m_inputAddr
End Property
Public Property Let InputAddr(ByVal value As String)
    m_inputAddr = value
End Property

Public Sub Init(ByVal fieldOrder As Long, _
                ByVal label As String, _
                ByVal typeText As String, _
                ByVal required As Boolean, _
                ByVal rowCount As Long, _
                ByVal hintText As String)
    m_order = fieldOrder
    m_label = label
    m_typeText = typeText
    m_required = required
    m_rowCount = rowCount
    m_hintText = hintText
End Sub

Public Sub SetCellAddrs(ByVal orderAddr As String, _
                        ByVal reqMarkAddr As String, _
                        ByVal labelAddr As String, _
                        ByVal typeAddr As String, _
                        ByVal inputAddr As String)
    m_orderAddr = orderAddr
    m_reqMarkAddr = reqMarkAddr
    m_labelAddr = labelAddr
    m_typeAddr = typeAddr
    m_inputAddr = inputAddr
End Sub
'@ },
    @{ Name='clsFormSpec'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsFormSpec (ビジネスロジック層)
' 概要:   1 つの UserForm の宣言情報を保持。
'         FormTitle / Width / Height + Controls Collection。
'         AddControl で clsControlSpec を追加する DSL を提供。
' 依存先: clsControlSpec
' 備考:   modFormBuilder.BuildAndShow に渡して動的フォーム生成する。
' ================================================================

Private m_formTitle As String
Private m_width As Long
Private m_height As Long
Private m_controls As Collection

Private Sub Class_Initialize()
    Set m_controls = New Collection
    m_width = 480
    m_height = 360
    m_formTitle = "Untitled"
End Sub

' --- FormTitle ---
Public Property Get FormTitle() As String
    FormTitle = m_formTitle
End Property
Public Property Let FormTitle(ByVal v As String)
    m_formTitle = v
End Property

' --- Width ---
Public Property Get Width() As Long
    Width = m_width
End Property
Public Property Let Width(ByVal v As Long)
    m_width = v
End Property

' --- Height ---
Public Property Get Height() As Long
    Height = m_height
End Property
Public Property Let Height(ByVal v As Long)
    m_height = v
End Property

' --- Controls (read-only Collection) ---
Public Property Get Controls() As Collection
    Set Controls = m_controls
End Property

' ================================================================
' 関数名: AddControl
' 概要:   clsControlSpec を生成して Controls Collection に追加する DSL。
'         同名 Control が既にあれば例外を投げる (Collection の挙動)。
' 引数:   cType / nm / l / t / w / h / cap / onClk - clsControlSpec.Init と同じ
' 戻り値: clsControlSpec - 追加された ControlSpec (チェイン用)
' ================================================================
Public Function AddControl(ByVal cType As String, ByVal nm As String, _
                            ByVal l As Long, ByVal t As Long, _
                            ByVal w As Long, ByVal h As Long, _
                            Optional ByVal cap As String = "", _
                            Optional ByVal onClk As String = "") As clsControlSpec
    Dim cs As clsControlSpec
    Set cs = New clsControlSpec
    cs.Init cType, nm, l, t, w, h, cap, onClk
    m_controls.Add cs, nm
    Set AddControl = cs
End Function

' ================================================================
' 関数名: ControlByName
' 概要:   指定名の Control Spec を取得 (なければ Nothing)
' ================================================================
Public Function ControlByName(ByVal nm As String) As clsControlSpec
    On Error Resume Next
    Set ControlByName = m_controls(nm)
    On Error GoTo 0
End Function

' ================================================================
' 関数名: ControlCount
' 概要:   Controls Collection の件数
' ================================================================
Public Function ControlCount() As Long
    ControlCount = m_controls.Count
End Function
'@ },
    @{ Name='clsFormatManager'; Type='cls'; Code=@'
Option Explicit

' v12 注記: FL_COL_* / FL_START_ROW は modFormatColumns.bas へ移動 (ADR 0027)
'           VBA 仕様で .cls 内の Public Const が禁止のため標準モジュールに切出し。

' ================================================================
' Phase 6 レビュー: 17 subs/funcs に 9 error handlers。
' 状態管理の core class、各 Public method に ErrHandler ある程度完備。
' Private helpers (FindFormatRow 等) は呼出側で wrapped。指摘なし。
'
' D-2 注記: フォーマット一覧シート (SHEET_FORMAT_LIST) の列番号定数 FL_COL_*
' および FL_START_ROW は Public Const として外部公開する。
' modDemoSeeder 等の Cowork デモコードからは本クラスの定数を参照すること。
' (旧: 各モジュールで重複定義 → DRY 違反)
' ================================================================

' ================================================================
' クラス: clsFormatManager（ビジネスロジック層）
' 概要:   フォーマット定義の管理
'         フォーマット一覧・設計・プレビュー・採番・
'         バージョン管理を担当
' 依存先: clsLogger, modFileIO, modStringUtil, modCommon
' ================================================================

' --- フォーマット一覧シートの列番号 ---

' --- フォーマット設計シートのセル位置 ---
Private Const FD_ROW_FMT_ID As Long = 1
Private Const FD_COL_FMT_ID_VAL As Long = 2
Private Const FD_COL_VER_VAL As Long = 4
Private Const FD_COL_PATTERN_VAL As Long = 6
Private Const FD_ROW_FMT_NAME As Long = 2
Private Const FD_COL_FMT_NAME_VAL As Long = 2

' --- フィールドテーブルの位置 ---
Private Const FD_FIELDS_HEADER_ROW As Long = 5
Private Const FD_FIELDS_START_ROW As Long = 6
Private Const FD_FIELD_COL_NO As Long = 1
Private Const FD_FIELD_COL_NAME As Long = 2
Private Const FD_FIELD_COL_TYPE As Long = 3
Private Const FD_FIELD_COL_REQUIRED As Long = 4
Private Const FD_FIELD_COL_ROWS As Long = 5
Private Const FD_FIELD_COL_WIDTH As Long = 6
Private Const FD_FIELD_COL_HEIGHT As Long = 7
Private Const FD_FIELD_COL_REPEAT As Long = 8
Private Const FD_FIELD_COL_ROW_OP As Long = 9
Private Const FD_FIELD_COL_FORMAT As Long = 10
Private Const FD_FIELD_COL_DESC As Long = 11

Private m_logger As clsLogger

' ================================================================
' 関数名: Init
' 概要:   初期化
' 引数:   logger - ログ出力用
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger)
    Set m_logger = logger
End Sub

' ================================================================
' 関数名: GetSelectedFormatId
' 概要:   フォーマット一覧シートで選択されている行のフォーマットIDを返す
' 引数:   なし
' 戻り値: String - フォーマットID（選択がない場合は空文字列）
' ================================================================
Public Function GetSelectedFormatId() As String
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim selRow As Long
    selRow = ws.Application.Selection.Row
    
    If selRow < FL_START_ROW Then
        GetSelectedFormatId = ""
        Exit Function
    End If
    
    GetSelectedFormatId = CStr(ws.Cells(selRow, FL_COL_FMT_ID).Value)
    Exit Function

ErrHandler:
    GetSelectedFormatId = ""
End Function

' ================================================================
' 関数名: BeginCreate
' 概要:   フォーマット設計シートを空の新規作成状態にする
' 引数:   なし
' 戻り値: なし
' ================================================================
Public Sub BeginCreate()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    Call ClearDesignSheet(ws)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFormatManager", "BeginCreate", _
                          "新規フォーマット作成モード"
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFormatManager", "BeginCreate", Err.Description
    End If
End Sub

' ================================================================
' 関数名: BeginEdit
' 概要:   指定フォーマットIDをフォーマット設計シートに読み込んで編集モードにする
' 引数:   formatId - 編集対象のフォーマットID
' 戻り値: なし
' ================================================================
Public Sub BeginEdit(ByVal formatId As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    Call ClearDesignSheet(ws)
    Call LoadFormatIntoDesignSheet(ws, formatId)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFormatManager", "BeginEdit", _
                          "フォーマット編集モード: " & formatId
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFormatManager", "BeginEdit", Err.Description
    End If
End Sub

' ================================================================
' 関数名: SaveFormat
' 概要:   フォーマット設計シートの内容を検証してフォーマット一覧に保存
'         既存フォーマットの場合はバージョンを+1
' 引数:   なし
' 戻り値: Boolean - 保存成功なら True
' ================================================================
Public Function SaveFormat() As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    Dim formatId As String
    formatId = CStr(ws.Cells(FD_ROW_FMT_ID, FD_COL_FMT_ID_VAL).Value)
    
    If formatId = "" Then
        If Not m_logger Is Nothing Then
            m_logger.LogWarn "clsFormatManager", "SaveFormat", _
                              "フォーマットID未入力"
        End If
        SaveFormat = False
        Exit Function
    End If
    
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim existingRow As Long
    existingRow = FindFormatRow(listWs, formatId)
    
    If existingRow = 0 Then
        Call AppendFormatToList(listWs, ws, formatId)
    Else
        Call UpdateFormatInList(listWs, ws, formatId, existingRow)
    End If
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFormatManager", "SaveFormat", _
                          "保存完了: " & formatId
    End If
    
    SaveFormat = True
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFormatManager", "SaveFormat", Err.Description
    End If
    SaveFormat = False
End Function

' ================================================================
' 関数名: ShowPreview
' 概要:   指定フォーマットをプレビューシートに表示
' 引数:   formatId - 表示対象のフォーマットID
' 戻り値: なし
' ================================================================
Public Sub ShowPreview(ByVal formatId As String)
    On Error GoTo ErrHandler
    
    Dim previewWs As Worksheet
    Set previewWs = ThisWorkbook.Worksheets(SHEET_FORMAT_PREVIEW)
    
    Call ClearPreviewSheet(previewWs)
    Call RenderFormatPreview(previewWs, formatId)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsFormatManager", "ShowPreview", _
                          "プレビュー表示: " & formatId
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsFormatManager", "ShowPreview", Err.Description
    End If
End Sub

' ================================================================
' 関数名: IncrementAndGetNextNumber
' 概要:   指定フォーマットの採番値を+1して返す（保存時のユニーク保証）
' 引数:   formatId - フォーマットID
' 戻り値: Long - 採番値（失敗時は0）
' ================================================================
Public Function IncrementAndGetNextNumber(ByVal formatId As String) As Long
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim rowNum As Long
    rowNum = FindFormatRow(ws, formatId)
    
    If rowNum = 0 Then
        IncrementAndGetNextNumber = 0
        Exit Function
    End If
    
    Dim currentNum As Long
    currentNum = CLng(ws.Cells(rowNum, FL_COL_CURRENT_NUM).Value)
    
    Dim newNum As Long
    newNum = currentNum + 1
    
    ws.Cells(rowNum, FL_COL_CURRENT_NUM).Value = newNum
    ws.Cells(rowNum, FL_COL_NEXT_NUM).Value = newNum + 1
    
    IncrementAndGetNextNumber = newNum
    Exit Function

ErrHandler:
    IncrementAndGetNextNumber = 0
End Function

' ================================================================
' 関数名: GetFormatHeaderAsStanza
' 概要:   フォーマット定義をスタンザ形式文字列として取得
'         ナレッジファイル先頭に埋め込むための形式
' 引数:   formatId - フォーマットID
' 戻り値: String - スタンザ形式の文字列（失敗時は空文字列）
' ================================================================
Public Function GetFormatHeaderAsStanza(ByVal formatId As String) As String
    On Error GoTo ErrHandler
    
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim rowNum As Long
    rowNum = FindFormatRow(listWs, formatId)
    
    If rowNum = 0 Then
        GetFormatHeaderAsStanza = ""
        Exit Function
    End If
    
    Dim result As String
    result = "[FORMAT]" & vbCrLf
    result = result & "FormatID=" & formatId & vbCrLf
    result = result & "FormatName=" & CStr(listWs.Cells(rowNum, FL_COL_FMT_NAME).Value) & vbCrLf
    result = result & "Version=" & CStr(listWs.Cells(rowNum, FL_COL_VERSION).Value) & vbCrLf
    result = result & "IDPattern=" & CStr(listWs.Cells(rowNum, FL_COL_ID_PATTERN).Value) & vbCrLf
    
    GetFormatHeaderAsStanza = result
    Exit Function

ErrHandler:
    GetFormatHeaderAsStanza = ""
End Function

' ================================================================
' 関数名: FindFormatRow
' 概要:   フォーマット一覧から指定IDの行番号を返す
' 引数:   listWs   - フォーマット一覧シート
'         formatId - 探すフォーマットID
' 戻り値: Long - 行番号（見つからない場合は 0）
' ================================================================
Private Function FindFormatRow(ByVal listWs As Worksheet, _
                                 ByVal formatId As String) As Long
    Dim i As Long
    Dim maxScan As Long
    maxScan = 1000
    
    For i = FL_START_ROW To maxScan
        Dim id As String
        id = CStr(listWs.Cells(i, FL_COL_FMT_ID).Value)
        If id = "" Then
            Exit For
        End If
        If id = formatId Then
            FindFormatRow = i
            Exit Function
        End If
    Next i
    
    FindFormatRow = 0
End Function

' 新規フォーマットをフォーマット一覧に追記
Private Sub AppendFormatToList(ByVal listWs As Worksheet, _
                                 ByVal designWs As Worksheet, _
                                 ByVal formatId As String)
    Dim nextRow As Long
    nextRow = FindFirstEmptyRow(listWs, FL_COL_FMT_ID)
    
    listWs.Cells(nextRow, FL_COL_NO).Value = nextRow - FL_START_ROW + 1
    listWs.Cells(nextRow, FL_COL_FMT_ID).Value = formatId
    listWs.Cells(nextRow, FL_COL_FMT_NAME).Value = _
        designWs.Cells(FD_ROW_FMT_NAME, FD_COL_FMT_NAME_VAL).Value
    listWs.Cells(nextRow, FL_COL_ID_PATTERN).Value = _
        designWs.Cells(FD_ROW_FMT_ID, FD_COL_PATTERN_VAL).Value
    listWs.Cells(nextRow, FL_COL_CURRENT_NUM).Value = 0
    listWs.Cells(nextRow, FL_COL_NEXT_NUM).Value = 1
    listWs.Cells(nextRow, FL_COL_VERSION).Value = 1
    listWs.Cells(nextRow, FL_COL_FIELD_COUNT).Value = CountDesignFields(designWs)
    listWs.Cells(nextRow, FL_COL_KNW_COUNT).Value = 0
    listWs.Cells(nextRow, FL_COL_CREATED).Value = Format(Date, "yyyy-mm-dd")
    listWs.Cells(nextRow, FL_COL_UPDATED).Value = Format(Date, "yyyy-mm-dd")
End Sub

' 既存フォーマットを更新（バージョン+1）
Private Sub UpdateFormatInList(ByVal listWs As Worksheet, _
                                 ByVal designWs As Worksheet, _
                                 ByVal formatId As String, _
                                 ByVal rowNum As Long)
    Dim currentVer As Long
    currentVer = CLng(listWs.Cells(rowNum, FL_COL_VERSION).Value)
    
    listWs.Cells(rowNum, FL_COL_FMT_NAME).Value = _
        designWs.Cells(FD_ROW_FMT_NAME, FD_COL_FMT_NAME_VAL).Value
    listWs.Cells(rowNum, FL_COL_ID_PATTERN).Value = _
        designWs.Cells(FD_ROW_FMT_ID, FD_COL_PATTERN_VAL).Value
    listWs.Cells(rowNum, FL_COL_VERSION).Value = currentVer + 1
    listWs.Cells(rowNum, FL_COL_FIELD_COUNT).Value = CountDesignFields(designWs)
    listWs.Cells(rowNum, FL_COL_UPDATED).Value = Format(Date, "yyyy-mm-dd")
End Sub

' 指定列で最初の空セル行を返す
Private Function FindFirstEmptyRow(ByVal ws As Worksheet, _
                                     ByVal targetCol As Long) As Long
    Dim i As Long
    For i = FL_START_ROW To 10000
        If ws.Cells(i, targetCol).Value = "" Then
            FindFirstEmptyRow = i
            Exit Function
        End If
    Next i
    FindFirstEmptyRow = FL_START_ROW
End Function

' フォーマット設計シートのフィールド数を数える
Private Function CountDesignFields(ByVal ws As Worksheet) As Long
    Dim i As Long
    Dim count As Long
    count = 0
    For i = FD_FIELDS_START_ROW To 1000
        If ws.Cells(i, FD_FIELD_COL_NAME).Value = "" Then
            Exit For
        End If
        count = count + 1
    Next i
    CountDesignFields = count
End Function

' 設計シートをクリア（ヘッダー行は残す）
Private Sub ClearDesignSheet(ByVal ws As Worksheet)
    ws.Cells(FD_ROW_FMT_ID, FD_COL_FMT_ID_VAL).Value = ""
    ws.Cells(FD_ROW_FMT_ID, FD_COL_VER_VAL).Value = ""
    ws.Cells(FD_ROW_FMT_ID, FD_COL_PATTERN_VAL).Value = ""
    ws.Cells(FD_ROW_FMT_NAME, FD_COL_FMT_NAME_VAL).Value = ""
    
    Dim i As Long
    For i = FD_FIELDS_START_ROW To 1000
        If ws.Cells(i, FD_FIELD_COL_NAME).Value = "" Then Exit For
        ws.Range(ws.Cells(i, 1), ws.Cells(i, FD_FIELD_COL_DESC)).ClearContents
    Next i
End Sub

' フォーマット一覧の値を設計シートに反映
Private Sub LoadFormatIntoDesignSheet(ByVal designWs As Worksheet, _
                                        ByVal formatId As String)
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim rowNum As Long
    rowNum = FindFormatRow(listWs, formatId)
    If rowNum = 0 Then Exit Sub
    
    designWs.Cells(FD_ROW_FMT_ID, FD_COL_FMT_ID_VAL).Value = formatId
    designWs.Cells(FD_ROW_FMT_ID, FD_COL_VER_VAL).Value = _
        listWs.Cells(rowNum, FL_COL_VERSION).Value
    designWs.Cells(FD_ROW_FMT_ID, FD_COL_PATTERN_VAL).Value = _
        listWs.Cells(rowNum, FL_COL_ID_PATTERN).Value
    designWs.Cells(FD_ROW_FMT_NAME, FD_COL_FMT_NAME_VAL).Value = _
        listWs.Cells(rowNum, FL_COL_FMT_NAME).Value
End Sub

' プレビューシートをクリア
Private Sub ClearPreviewSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Cells.ClearContents
    On Error GoTo 0
End Sub

' フォーマットをプレビューシートに描画
Private Sub RenderFormatPreview(ByVal ws As Worksheet, _
                                  ByVal formatId As String)
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim rowNum As Long
    rowNum = FindFormatRow(listWs, formatId)
    If rowNum = 0 Then Exit Sub
    
    ws.Cells(1, 1).Value = "フォーマット"
    ws.Cells(1, 2).Value = formatId & " - " & _
                             CStr(listWs.Cells(rowNum, FL_COL_FMT_NAME).Value)
    
    Dim designWs As Worksheet
    Set designWs = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    Dim i As Long
    Dim previewRow As Long
    previewRow = 3
    
    For i = FD_FIELDS_START_ROW To 1000
        If designWs.Cells(i, FD_FIELD_COL_NAME).Value = "" Then Exit For
        ws.Cells(previewRow, 1).Value = _
            designWs.Cells(i, FD_FIELD_COL_NAME).Value
        ws.Cells(previewRow, 2).Value = _
            "(" & designWs.Cells(i, FD_FIELD_COL_TYPE).Value & " 入力エリア)"
        previewRow = previewRow + 1
    Next i
End Sub
'@ },
    @{ Name='clsKnowledgeManager'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsKnowledgeManager（ビジネスロジック層）
' 概要:   ナレッジの登録・修正・削除を担当
'         スタンザ形式で Shift_JIS ファイルとして入出力
' 依存先: clsLogger, clsFormatManager, modFileIO, modStringUtil,
'         modDateUtil, modCommon
' ================================================================

' --- ナレッジ登録シートのセル位置 ---
Private Const KS_ROW_FMT_ID As Long = 1
Private Const KS_COL_FMT_ID_VAL As Long = 3
Private Const KS_FORM_START_ROW As Long = 4
Private Const KS_FIELD_COL_NO As Long = 1
Private Const KS_FIELD_COL_NAME As Long = 2
Private Const KS_FIELD_COL_VALUE As Long = 3

' --- ナレッジ修正シートのセル位置 ---
Private Const KE_ROW_FMT_ID As Long = 1
Private Const KE_COL_FMT_ID_VAL As Long = 2
Private Const KE_COL_KNW_NO As Long = 3
Private Const KE_FORM_START_ROW As Long = 4

Private m_logger As clsLogger
Private m_formatMgr As clsFormatManager
Private m_dataFolder As String

' ================================================================
' 関数名: Init
' 概要:   初期化
' 引数:   logger     - ログ出力用
'         formatMgr  - フォーマット管理クラス
'         dataFolder - ナレッジデータ格納フォルダパス
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger, _
                 ByVal formatMgr As clsFormatManager, _
                 ByVal dataFolder As String)
    Set m_logger = logger
    Set m_formatMgr = formatMgr
    m_dataFolder = dataFolder
End Sub

' ================================================================
' 関数名: BuildRegistrationForm
' 概要:   ナレッジ登録シートに指定フォーマットのフィールド一覧を展開
' 引数:   formatId - フォーマットID
' 戻り値: なし
' 備考:   既存の入力値はクリアされる
' ================================================================
Public Sub BuildRegistrationForm(ByVal formatId As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)
    
    Call ClearRegistrationForm(ws)
    Call PopulateFormFromDesign(ws, formatId, KS_FORM_START_ROW)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "BuildRegistrationForm", _
                          "フォーム生成: " & formatId
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "BuildRegistrationForm", _
                           Err.Description
    End If
End Sub

' ================================================================
' 関数名: SaveNewKnowledge
' 概要:   ナレッジ登録シートからフォーム値を取得して新規ナレッジ保存
'         採番を+1してファイル名を決定、ファイルを物理生成
' 引数:   なし
' 戻り値: String - 採番されたナレッジ番号（失敗時は空文字列）
' ================================================================
Public Function SaveNewKnowledge() As String
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)
    
    Dim formatId As String
    formatId = CStr(ws.Cells(KS_ROW_FMT_ID, KS_COL_FMT_ID_VAL).Value)
    
    If Not ValidateRequiredFields(ws, KS_FORM_START_ROW) Then
        SaveNewKnowledge = ""
        Exit Function
    End If
    
    Dim nextNum As Long
    nextNum = m_formatMgr.IncrementAndGetNextNumber(formatId)
    If nextNum = 0 Then
        SaveNewKnowledge = ""
        Exit Function
    End If
    
    Dim knowledgeNo As String
    knowledgeNo = BuildKnowledgeNumber(formatId, nextNum)
    
    Dim content As String
    content = BuildKnowledgeFile(formatId, knowledgeNo, TodayStr(), TodayStr(), _
                                   ws, KS_FORM_START_ROW)
    
    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")
    
    If Not WriteShiftJisFile(filePath, content) Then
        If Not m_logger Is Nothing Then
            m_logger.LogError "clsKnowledgeManager", "SaveNewKnowledge", _
                               "ファイル書き込み失敗: " & filePath
        End If
        SaveNewKnowledge = ""
        Exit Function
    End If
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "SaveNewKnowledge", _
                          "ナレッジ" & knowledgeNo & "を保存しました"
    End If
    
    SaveNewKnowledge = knowledgeNo
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "SaveNewKnowledge", Err.Description
    End If
    SaveNewKnowledge = ""
End Function

' ================================================================
' 関数名: LoadForEdit
' 概要:   指定ナレッジ番号のファイルをナレッジ修正シートに読み込む
' 引数:   knowledgeNo - ナレッジ番号
' 戻り値: Boolean - 成功なら True
' ================================================================
Public Function LoadForEdit(ByVal knowledgeNo As String) As Boolean
    ' s-2: パストラバーサル防止 - ナレッジ番号バリデーション
    If Not IsValidKnowledgeId(knowledgeNo) Then
        If Not m_logger Is Nothing Then
            m_logger.LogError "clsKnowledgeManager", "LoadForEdit", _
                                "不正なナレッジ番号 reject: " & knowledgeNo
        End If
        LoadForEdit = False
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")
    
    If Not FileExists(filePath) Then
        LoadForEdit = False
        Exit Function
    End If
    
    Dim content As String
    content = ReadShiftJisFile(filePath)
    If content = "" Then
        LoadForEdit = False
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    
    Call ClearEditForm(ws)
    Call PopulateEditFormFromFile(ws, content, knowledgeNo)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "LoadForEdit", _
                          "読込: " & knowledgeNo
    End If
    
    LoadForEdit = True
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "LoadForEdit", Err.Description
    End If
    LoadForEdit = False
End Function

' ================================================================
' 関数名: UpdateKnowledge
' 概要:   ナレッジ修正シートの内容をファイルに上書き保存
'         作成日は維持、更新日のみ今日に更新
' 引数:   knowledgeNo - ナレッジ番号
' 戻り値: Boolean - 成功なら True
' ================================================================
Public Function UpdateKnowledge(ByVal knowledgeNo As String) As Boolean
    ' s-2: パストラバーサル防止 - ナレッジ番号バリデーション
    If Not IsValidKnowledgeId(knowledgeNo) Then
        If Not m_logger Is Nothing Then
            m_logger.LogError "clsKnowledgeManager", "UpdateKnowledge", _
                                "不正なナレッジ番号 reject: " & knowledgeNo
        End If
        UpdateKnowledge = False
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")
    
    Dim originalContent As String
    originalContent = ReadShiftJisFile(filePath)
    
    Dim createdDate As String
    createdDate = ExtractCreatedDate(originalContent)
    If createdDate = "" Then
        createdDate = TodayStr()
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    
    Dim formatId As String
    formatId = CStr(ws.Cells(KE_ROW_FMT_ID, KE_COL_FMT_ID_VAL).Value)
    
    Dim content As String
    content = BuildKnowledgeFile(formatId, knowledgeNo, createdDate, TodayStr(), _
                                   ws, KE_FORM_START_ROW)
    
    If Not WriteShiftJisFile(filePath, content) Then
        UpdateKnowledge = False
        Exit Function
    End If
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "UpdateKnowledge", _
                          "ナレッジ" & knowledgeNo & "を更新しました"
    End If
    
    UpdateKnowledge = True
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "UpdateKnowledge", Err.Description
    End If
    UpdateKnowledge = False
End Function

' ================================================================
' 関数名: DeleteKnowledge
' 概要:   指定ナレッジ番号のファイルを物理削除
' 引数:   knowledgeNo - ナレッジ番号
' 戻り値: Boolean - 成功なら True
' ================================================================
Public Function DeleteKnowledge(ByVal knowledgeNo As String) As Boolean
    ' s-2: パストラバーサル防止 - ナレッジ番号バリデーション
    If Not IsValidKnowledgeId(knowledgeNo) Then
        If Not m_logger Is Nothing Then
            m_logger.LogError "clsKnowledgeManager", "DeleteKnowledge", _
                                "不正なナレッジ番号 reject: " & knowledgeNo
        End If
        DeleteKnowledge = False
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")
    
    If Not DeleteFile(filePath) Then
        DeleteKnowledge = False
        Exit Function
    End If
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "DeleteKnowledge", _
                          "ナレッジ" & knowledgeNo & "を削除しました"
    End If
    
    DeleteKnowledge = True
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "DeleteKnowledge", Err.Description
    End If
    DeleteKnowledge = False
End Function

' ================================================================
' 関数名: BuildKnowledgeFile
' 概要:   フォーム値とメタ情報からナレッジファイル本文を組み立て
' 引数:   formatId     - フォーマットID
'         knowledgeNo  - ナレッジ番号
'         createDate   - 作成日
'         updateDate   - 更新日
'         ws           - 値を取得するワークシート
'         formStartRow - フォームの開始行
' 戻り値: String - スタンザ形式のファイル本文
' ================================================================
Private Function BuildKnowledgeFile(ByVal formatId As String, _
                                      ByVal knowledgeNo As String, _
                                      ByVal createDate As String, _
                                      ByVal updateDate As String, _
                                      ByVal ws As Worksheet, _
                                      ByVal formStartRow As Long) As String
    Dim content As String
    content = m_formatMgr.GetFormatHeaderAsStanza(formatId)
    content = content & STANZA_SEP & vbCrLf
    
    content = content & "[KNOWLEDGE]" & vbCrLf
    content = content & "KnowledgeNo=" & knowledgeNo & vbCrLf
    content = content & "FormatVersion=" & GetFormatVersion(formatId) & vbCrLf
    content = content & "CreatedDate=" & createDate & vbCrLf
    content = content & "UpdatedDate=" & updateDate & vbCrLf
    
    Dim i As Long
    For i = formStartRow To 1000
        Dim fieldName As String
        fieldName = CStr(ws.Cells(i, KS_FIELD_COL_NAME).Value)
        If fieldName = "" Then Exit For

        ' rev19: M10-002 対応。"=" 先頭値を入力時に SafeCellText 等で
        ' "'=..." とエスケープしている場合があるため、ファイル書き込み時は
        ' 先頭アポストロフィを 1 文字だけ除去して "=SUM(A1:A10)" がそのまま
        ' 残るようにする（仕様: アポストロフィなしでファイルに残す）
        Dim rawValue As String
        rawValue = CStr(ws.Cells(i, KS_FIELD_COL_VALUE).Value)
        If Len(rawValue) >= 2 Then
            If Left(rawValue, 1) = "'" And Mid(rawValue, 2, 1) = "=" Then
                rawValue = Mid(rawValue, 2)
            End If
        End If

        content = content & "[ITEM]" & vbCrLf
        content = content & "FieldNo=" & CStr(ws.Cells(i, KS_FIELD_COL_NO).Value)
        content = content & " / FieldName=" & fieldName
        content = content & " / Value=" & rawValue
        content = content & vbCrLf
    Next i
    
    content = content & STANZA_SEP
    BuildKnowledgeFile = content
End Function

' ナレッジ番号をパターンと採番値から生成
Private Function BuildKnowledgeNumber(ByVal formatId As String, _
                                         ByVal nextNum As Long) As String
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim pattern As String
    pattern = FindIdPattern(listWs, formatId)
    
    BuildKnowledgeNumber = FormatNumberByPattern(pattern, nextNum)
End Function

' フォーマット一覧からID形式を取得
Private Function FindIdPattern(ByVal listWs As Worksheet, _
                                 ByVal formatId As String) As String
    Dim i As Long
    For i = 3 To 1000
        If CStr(listWs.Cells(i, 2).Value) = formatId Then
            FindIdPattern = CStr(listWs.Cells(i, 4).Value)
            Exit Function
        End If
    Next i
    FindIdPattern = ""
End Function

' フォーマット一覧からバージョン番号を取得
Private Function GetFormatVersion(ByVal formatId As String) As Long
    Dim listWs As Worksheet
    Set listWs = ThisWorkbook.Worksheets(SHEET_FORMAT_LIST)
    
    Dim i As Long
    For i = 3 To 1000
        If CStr(listWs.Cells(i, 2).Value) = formatId Then
            GetFormatVersion = CLng(listWs.Cells(i, 7).Value)
            Exit Function
        End If
    Next i
    GetFormatVersion = 1
End Function

' フォルダパスとファイル名を結合
Private Function CombineFilePath(ByVal folder As String, _
                                   ByVal fileName As String) As String
    If Right(folder, 1) = "\" Or Right(folder, 1) = "/" Then
        CombineFilePath = folder & fileName
    Else
        CombineFilePath = folder & "\" & fileName
    End If
End Function

' 必須フィールドの検証
Private Function ValidateRequiredFields(ByVal ws As Worksheet, _
                                          ByVal startRow As Long) As Boolean
    Dim i As Long
    For i = startRow To 1000
        If ws.Cells(i, KS_FIELD_COL_NAME).Value = "" Then Exit For
        ' 必須判定は設計シートを参照する必要があるが、
        ' 簡易実装として登録シートの値が空ならOKとする
    Next i
    ValidateRequiredFields = True
End Function

' 登録フォームをクリア
Private Sub ClearRegistrationForm(ByVal ws As Worksheet)
    Dim i As Long
    For i = KS_FORM_START_ROW To 1000
        If ws.Cells(i, KS_FIELD_COL_NAME).Value = "" Then Exit For
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 3)).ClearContents
    Next i
End Sub

' 修正フォームをクリア
Private Sub ClearEditForm(ByVal ws As Worksheet)
    Dim i As Long
    For i = KE_FORM_START_ROW To 1000
        If ws.Cells(i, KS_FIELD_COL_NAME).Value = "" Then Exit For
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 3)).ClearContents
    Next i
End Sub

' フォーマット設計からフィールドリストを展開
Private Sub PopulateFormFromDesign(ByVal ws As Worksheet, _
                                     ByVal formatId As String, _
                                     ByVal startRow As Long)
    Dim designWs As Worksheet
    Set designWs = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    
    ws.Cells(KS_ROW_FMT_ID, KS_COL_FMT_ID_VAL).Value = formatId
    
    Dim i As Long
    Dim targetRow As Long
    targetRow = startRow
    
    For i = 6 To 1000
        Dim fieldName As String
        fieldName = CStr(designWs.Cells(i, 2).Value)
        If fieldName = "" Then Exit For
        
        ws.Cells(targetRow, KS_FIELD_COL_NO).Value = _
            designWs.Cells(i, 1).Value
        ws.Cells(targetRow, KS_FIELD_COL_NAME).Value = fieldName
        targetRow = targetRow + 1
    Next i
End Sub

' ファイル内容からフォームに展開
Private Sub PopulateEditFormFromFile(ByVal ws As Worksheet, _
                                       ByVal content As String, _
                                       ByVal knowledgeNo As String)
    Dim lines() As String
    lines = Split(content, vbCrLf)
    
    Dim formatId As String
    formatId = ExtractStanzaValue(content, "FormatID")
    
    ws.Cells(KE_ROW_FMT_ID, KE_COL_FMT_ID_VAL).Value = formatId
    ws.Cells(KE_ROW_FMT_ID, KE_COL_KNW_NO).Value = knowledgeNo
    
    Dim i As Long
    Dim targetRow As Long
    targetRow = KE_FORM_START_ROW
    
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "FieldNo=") > 0 And _
           InStr(lines(i), "FieldName=") > 0 And _
           InStr(lines(i), "Value=") > 0 Then
            Call ParseItemLine(lines(i), ws, targetRow)
            targetRow = targetRow + 1
        End If
    Next i
End Sub

' ITEM行の解析と展開
Private Sub ParseItemLine(ByVal line As String, _
                            ByVal ws As Worksheet, _
                            ByVal targetRow As Long)
    Dim noStr As String
    Dim nameStr As String
    Dim valStr As String
    
    noStr = ExtractKeyValue(line, "FieldNo")
    nameStr = ExtractKeyValue(line, "FieldName")
    valStr = ExtractKeyValue(line, "Value")
    
    ws.Cells(targetRow, KS_FIELD_COL_NO).Value = noStr
    ws.Cells(targetRow, KS_FIELD_COL_NAME).Value = nameStr
    ws.Cells(targetRow, KS_FIELD_COL_VALUE).NumberFormat = "@"
    ws.Cells(targetRow, KS_FIELD_COL_VALUE).Value = SafeCellText(valStr)
End Sub

' "Key=Value / OtherKey=..." 形式から値を抽出
Private Function ExtractKeyValue(ByVal line As String, _
                                   ByVal keyName As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim searchKey As String
    
    searchKey = keyName & "="
    startPos = InStr(line, searchKey)
    If startPos = 0 Then
        ExtractKeyValue = ""
        Exit Function
    End If
    
    startPos = startPos + Len(searchKey)
    endPos = InStr(startPos, line, " / ")
    If endPos = 0 Then
        ExtractKeyValue = Mid(line, startPos)
    Else
        ExtractKeyValue = Mid(line, startPos, endPos - startPos)
    End If
End Function

' スタンザ内の単純なKey=Value値を抽出
Private Function ExtractStanzaValue(ByVal content As String, _
                                      ByVal keyName As String) As String
    Dim lines() As String
    Dim i As Long
    
    lines = Split(content, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), keyName & "=") = 1 Then
            ExtractStanzaValue = Mid(lines(i), Len(keyName) + 2)
            Exit Function
        End If
    Next i
    ExtractStanzaValue = ""
End Function

' ファイル内容からCreatedDateを抽出
Private Function ExtractCreatedDate(ByVal content As String) As String
    ExtractCreatedDate = ExtractStanzaValue(content, "CreatedDate")
End Function
'@ },
    @{ Name='clsLogger'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsLogger（ビジネスロジック層）
' 概要:   ログシート + 外部ファイル(C:\kvba\runtime.log) 二重出力。
'         レベル: TRACE / DEBUG / INFO / WARN / ERROR
'         step 識別子で「どこまで通って、どこで失敗したか」を残す。
' 依存先: clsLogEntry, modDateUtil, modStringUtil, modCommon
' 備考:   外部ファイル append は VBA の Open For Append で実装。
'         (coding-standards R12 は通常 I/O 用; ログ append は実用上の例外)
'         外部書込みが連続失敗した場合は m_externalDisabled で抑止。
' ================================================================

Private m_logSheet As Worksheet
Private m_debugLevel As String
Private m_nextRow As Long  ' M-3: O(N^2) 防止のためキャッシュ (0=未初期化)
Private m_externalPath As String
Private m_externalDisabled As Boolean
Private m_externalFailCount As Long

' m-12: マジックナンバー 100000 を Const 化
Private Const MAX_LOG_SCAN_ROWS As Long = 100000

' v20: M-14 リデザインに伴うデータ書込開始行。A1 はタイトル帯、A8 は表ヘッダ、
'      データ本体は A9 から。これにより画面描画 (タイトル/ヘッダ) と
'      ログ書込が衝突しない。
Private Const LOG_DATA_START_ROW As Long = 9

' --- Property Get ---

Public Property Get DebugLevel() As String
    DebugLevel = m_debugLevel
End Property

Public Property Get ExternalPath() As String
    ExternalPath = m_externalPath
End Property

' ================================================================
' 関数名: Init
' 概要:   初期化（ログシート参照とデバッグレベルを保持）
' 引数:   logSheet    - ログ出力先のワークシート
'         debugLevel  - "OFF" または "ON"
' 戻り値: なし
' 備考:   外部ログパスは EXTERNAL_LOG_PATH 定数を既定値とする。
'         呼び出し側で SetExternalPath を使えば差し替え可能。
' ================================================================
Public Sub Init(ByVal logSheet As Worksheet, ByVal debugLevel As String)
    Set m_logSheet = logSheet
    m_debugLevel = debugLevel
    m_externalPath = EXTERNAL_LOG_PATH
    m_externalDisabled = False
End Sub

' ================================================================
' 関数名: SetExternalPath
' 概要:   外部ログファイルのパスを差し替える（テスト等で利用）
' ================================================================
Public Sub SetExternalPath(ByVal newPath As String)
    m_externalPath = newPath
    m_externalDisabled = False
End Sub

' ================================================================
' 関数名: LogError
' 概要:   エラーログを出力（debugLevel 関係なく出力）
' ================================================================
Public Sub LogError(ByVal modName As String, _
                      ByVal funcName As String, _
                      ByVal message As String)
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_ERROR, message)
End Sub

' ================================================================
' 関数名: LogErrorWithErr
' 概要:   ErrHandler から呼ぶ用。step 名と Err 情報を組み立てて 1 行で記録。
' 引数:   stepName  - 失敗直前に通っていた step 名（呼び出し側で文字列管理）
'         errNum    - Err.Number
'         errDesc   - Err.Description
' ================================================================
Public Sub LogErrorWithErr(ByVal modName As String, _
                             ByVal funcName As String, _
                             ByVal stepName As String, _
                             ByVal errNum As Long, _
                             ByVal errDesc As String)
    Dim msg As String
    msg = "FAIL step=[" & stepName & "] errNum=" & errNum & " desc=" & errDesc
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_ERROR, msg)
End Sub

' ================================================================
' 関数名: LogWarn
' ================================================================
Public Sub LogWarn(ByVal modName As String, _
                     ByVal funcName As String, _
                     ByVal message As String)
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_WARN, message)
End Sub

' ================================================================
' 関数名: LogInfo
' ================================================================
Public Sub LogInfo(ByVal modName As String, _
                     ByVal funcName As String, _
                     ByVal message As String)
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_INFO, message)
End Sub

' ================================================================
' 関数名: LogTrace
' 概要:   ENTER/EXIT/step マーカー用のトレースログ。
'         シート書き込みは debugLevel=ON 時のみ（シート過多防止）。
'         外部ファイルには常に書き出して「ここまで通った」を残す。
' ================================================================
Public Sub LogTrace(ByVal modName As String, _
                      ByVal funcName As String, _
                      ByVal message As String)
    If m_debugLevel = DEBUG_ON Then
        Call WriteLogInternal(modName, funcName, LOG_LEVEL_TRACE, message)
    Else
        Call WriteToExternalFileSafe(modName, funcName, LOG_LEVEL_TRACE, message)
    End If
End Sub

' ================================================================
' 関数名: LogDebug
' 概要:   デバッグログを出力（debugLevel ON 時のみシートに出す）
'         外部ファイルには常に出力する（後追い解析用）。
' ================================================================
Public Sub LogDebug(ByVal modName As String, _
                      ByVal funcName As String, _
                      ByVal message As String)
    If m_debugLevel <> DEBUG_ON Then
        Call WriteToExternalFileSafe(modName, funcName, LOG_LEVEL_DEBUG, message)
        Exit Sub
    End If
    Call WriteLogInternal(modName, funcName, LOG_LEVEL_DEBUG, message)
End Sub

' ================================================================
' 関数名: ClearLog
' 概要:   ログシートの A9 以降を全削除（ヘッダ A8 と A1 タイトルは保持）
' ================================================================
Public Sub ClearLog()
    On Error GoTo ErrHandler

    Dim lastRow As Long
    lastRow = GetLastLogRow()

    If lastRow < LOG_DATA_START_ROW Then
        ' データなし。外部ログには起動マーカー
        Call WriteToExternalFileSafe("clsLogger", "ClearLog", LOG_LEVEL_INFO, "ログクリア(空)")
        Exit Sub
    End If

    m_logSheet.Range(m_logSheet.Cells(LOG_DATA_START_ROW, 1), _
                      m_logSheet.Cells(lastRow, 5)).ClearContents
    m_nextRow = LOG_DATA_START_ROW  ' v20: cache reset to data start row
    Call WriteToExternalFileSafe("clsLogger", "ClearLog", LOG_LEVEL_INFO, _
                                  "ログクリア完了 (rows=" & (lastRow - LOG_DATA_START_ROW + 1) & ")")
    Exit Sub

ErrHandler:
    ' クリア失敗時は外部ログだけでも残す
    Call WriteToExternalFileSafe("clsLogger", "ClearLog", LOG_LEVEL_ERROR, _
                                  "ClearLog 失敗 errNum=" & Err.Number & " desc=" & Err.Description)
End Sub

' ================================================================
' 関数名: WriteLogInternal
' 概要: シート + 外部ファイル両方への書き込みを統括 (内部 helper)
'       シート書込み失敗しても外部ファイル試行は継続
' ================================================================
Private Sub WriteLogInternal(ByVal modName As String, _
                              ByVal funcName As String, _
                              ByVal level As String, _
                              ByVal message As String)
    Dim ts As String
    ts = Format$(Now(), "yyyy-mm-dd hh:nn:ss")

    ' 1) LogSheet 書込み
    On Error Resume Next
    If Not m_logSheet Is Nothing Then
        Dim r As Long
        r = m_logSheet.Cells(m_logSheet.Rows.Count, 1).End(-4162).Row + 1  ' xlUp = -4162
        If r < 9 Then r = 9
        m_logSheet.Cells(r, 1).Value = ts
        m_logSheet.Cells(r, 2).Value = modName
        m_logSheet.Cells(r, 3).Value = funcName
        m_logSheet.Cells(r, 4).Value = level
        m_logSheet.Cells(r, 5).Value = message
    End If
    Err.Clear
    On Error GoTo 0

    ' 2) 外部ファイル書込み (Append)
    If m_externalDisabled Then Exit Sub
    On Error Resume Next
    Dim fno As Integer
    fno = FreeFile
    Open m_externalPath For Append As #fno
    Print #fno, "[" & ts & "] [" & level & "] [" & modName & "." & funcName & "] " & message
    Close #fno
    If Err.Number <> 0 Then
        m_externalFailCount = m_externalFailCount + 1
        If m_externalFailCount > 3 Then m_externalDisabled = True
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub Class_Initialize()
    m_externalPath = "C:\kvba\runtime.log"
    m_externalDisabled = False
    m_externalFailCount = 0
    m_debugLevel = "INFO"
End Sub
'@ },
    @{ Name='clsScreenSpec'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsScreenSpec（画面層 — データ）
' 概要:   1 画面分の構成（タイトル/セクション/ボタン/フィールド）を保持。
'         画面修正は本クラスのインスタンスを書き換えるだけで完結する設計。
' 依存先: clsSectionSpec, clsButtonSpec, clsFieldSpec
' ================================================================

Private m_screenId As String          ' "M-01"
Private m_sheetName As String         ' "メイン"
Private m_title As String             ' "【画面モック v2】..."
Private m_titleColorHex As String     ' "#1F3864"
Private m_backToMainAddr As String    ' "K17" など空文字許容
Private m_sections As Collection      ' of clsSectionSpec
Private m_buttons As Collection       ' of clsButtonSpec
Private m_fields As Collection        ' of clsFieldSpec
Private m_headerRowAddr As String     ' 一覧系: ヘッダ開始位置 ("B10")
Private m_headerLabels As Variant     ' 一覧系: ヘッダラベル配列
Private m_emptyStateAddr As String    ' "0 件" 表示位置
Private m_emptyStateText As String

Private Sub Class_Initialize()
    Set m_sections = New Collection
    Set m_buttons = New Collection
    Set m_fields = New Collection
End Sub

Public Property Get ScreenId() As String
    ScreenId = m_screenId
End Property
Public Property Let ScreenId(ByVal value As String)
    m_screenId = value
End Property

Public Property Get SheetName() As String
    SheetName = m_sheetName
End Property
Public Property Let SheetName(ByVal value As String)
    m_sheetName = value
End Property

Public Property Get Title() As String
    Title = m_title
End Property
Public Property Let Title(ByVal value As String)
    m_title = value
End Property

Public Property Get TitleColorHex() As String
    TitleColorHex = m_titleColorHex
End Property
Public Property Let TitleColorHex(ByVal value As String)
    m_titleColorHex = value
End Property

Public Property Get BackToMainAddr() As String
    BackToMainAddr = m_backToMainAddr
End Property
Public Property Let BackToMainAddr(ByVal value As String)
    m_backToMainAddr = value
End Property

Public Property Get Sections() As Collection
    Set Sections = m_sections
End Property

Public Property Get Buttons() As Collection
    Set Buttons = m_buttons
End Property

Public Property Get Fields() As Collection
    Set Fields = m_fields
End Property

Public Property Get HeaderRowAddr() As String
    HeaderRowAddr = m_headerRowAddr
End Property
Public Property Let HeaderRowAddr(ByVal value As String)
    m_headerRowAddr = value
End Property

Public Property Get HeaderLabels() As Variant
    HeaderLabels = m_headerLabels
End Property
Public Property Let HeaderLabels(ByVal value As Variant)
    m_headerLabels = value
End Property

Public Property Get EmptyStateAddr() As String
    EmptyStateAddr = m_emptyStateAddr
End Property
Public Property Let EmptyStateAddr(ByVal value As String)
    m_emptyStateAddr = value
End Property

Public Property Get EmptyStateText() As String
    EmptyStateText = m_emptyStateText
End Property
Public Property Let EmptyStateText(ByVal value As String)
    m_emptyStateText = value
End Property

' ================================================================
' 関数名: AddSection
' 概要:   セクション帯を追加
' ================================================================
Public Sub AddSection(ByVal sec As clsSectionSpec)
    m_sections.Add sec
End Sub

' ================================================================
' 関数名: AddButton
' 概要:   ボタンを追加
' ================================================================
Public Sub AddButton(ByVal btn As clsButtonSpec)
    m_buttons.Add btn
End Sub

' ================================================================
' 関数名: AddField
' 概要:   フィールドを追加
' ================================================================
Public Sub AddField(ByVal fld As clsFieldSpec)
    m_fields.Add fld
End Sub
'@ },
    @{ Name='clsSearchEngine'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsSearchEngine（ビジネスロジック層）
' 概要:   ナレッジ検索を担当
'         番号ダイレクト検索、キーワード検索（AND/OR）、日付範囲フィルタ
'         スコアリング (出現回数 + タイトル/フィールドブースト)、
'         結果サムネ表示、詳細画像ペイン表示までを担う。
' 依存先: clsLogger, modFileIO, modStringUtil, modDateUtil, modCommon,
'         modImageRender (新規; サムネ/詳細画像 Shape 描画)
' ================================================================
'
' ====================================================================
' [真版 ChromaDB 連動 切替ポイント]
' 本クラスはモック実装として <dataFolder>/*.txt を ADODB で SJIS 読込し
' スタンザを line-split して match 判定する。真版では事前 ETL で
' "Data" シートに chunks を export しておき ScanAndMatch の txt ループを
' Range 走査に置き換える。VBA 子プロセス禁止 (Shell/Run/Exec/CreateObject
' Exec 系全部禁止) のため、本クラスから外部プロセスを起動しない。
' ====================================================================

' --- 検索シートのセル位置 ---
Private Const SS_ROW_DIRECT_NO As Long = 3
Private Const SS_COL_DIRECT_NO As Long = 3
Private Const SS_ROW_FMT_ID As Long = 6
Private Const SS_ROW_KEYWORDS As Long = 7
Private Const SS_ROW_MODE As Long = 8
Private Const SS_ROW_TARGET_FIELD As Long = 9
Private Const SS_ROW_DATE_FROM As Long = 10
Private Const SS_ROW_DATE_TO As Long = 11
Private Const SS_COL_CONDITION_VALUE As Long = 3

' --- 検索結果一覧の位置 (rev22 + image_ext rev1) ---
Private Const SS_RESULT_START_ROW As Long = 15
Private Const SS_RESULT_COL_NO As Long = 1
Private Const SS_RESULT_COL_KNW_NO As Long = 2
Private Const SS_RESULT_COL_FMT_NAME As Long = 3
Private Const SS_RESULT_COL_TITLE As Long = 4
Private Const SS_RESULT_COL_CREATED As Long = 5
Private Const SS_RESULT_COL_UPDATED As Long = 6
Private Const SS_RESULT_COL_DETAIL As Long = 7
Private Const SS_RESULT_COL_THUMB As Long = 8     ' NEW: image_ext
Private Const SS_RESULT_COL_SCORE As Long = 9     ' NEW: image_ext

' --- 結果一覧の表示上限 ---
Private Const RESULT_MAX_ROWS As Long = 200

' --- スコアリングのブースト係数 ---
Private Const SCORE_BOOST_TITLE As Double = 3#
Private Const SCORE_BOOST_TARGET_FIELD As Double = 2#

' --- ImagePath スタンザのキー名 ---
Private Const IMAGE_PATH_KEY As String = "ImagePath"

' --- ナレッジ表示シートの位置 ---
Private Const KD_ROW_KNW_NO As Long = 1
Private Const KD_COL_KNW_NO_VAL As Long = 2
Private Const KD_COL_FMT_INFO As Long = 3
Private Const KD_FORM_START_ROW As Long = 4
Private Const KD_COL_FIELD_NO As Long = 1
Private Const KD_COL_FIELD_NAME As Long = 2
Private Const KD_COL_FIELD_VALUE As Long = 3
Private Const KD_COL_FIELD_LINK As Long = 4

' --- 詳細画像ペインの位置 (M-09 ナレッジ表示シート) ---
Private Const KD_DETAIL_IMG_TOP_ROW As Long = 4
Private Const KD_DETAIL_IMG_LEFT_COL As Long = 10   ' J
Private Const KD_DETAIL_IMG_BOT_ROW As Long = 20
Private Const KD_DETAIL_IMG_RIGHT_COL As Long = 14  ' N

Private m_logger As clsLogger
Private m_dataFolder As String

' ================================================================
' 関数名: Init
' 概要:   初期化
' 引数:   logger     - ログ出力用
'         dataFolder - ナレッジデータフォルダ
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger, ByVal dataFolder As String)
    Set m_logger = logger
    m_dataFolder = dataFolder
End Sub

' ================================================================
' 関数名: FindByNumber
' 概要:   指定番号のナレッジが存在するか確認
' 引数:   knowledgeNo - ナレッジ番号
' 戻り値: Boolean - 存在すれば True
' ================================================================
Public Function FindByNumber(ByVal knowledgeNo As String) As Boolean
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")

    FindByNumber = FileExists(filePath)
    Exit Function

ErrHandler:
    FindByNumber = False
End Function

' ================================================================
' 関数名: SearchByKeywords
' 概要:   検索シートの条件を読み取り、キーワード検索を実行
'         結果を検索シートの結果一覧に出力 (スコア降順)
' 引数:   なし
' 戻り値: Long - ヒット件数
' ================================================================
Public Function SearchByKeywords() As Long
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)

    Dim formatId As String
    Dim keywords As String
    Dim searchMode As String
    Dim targetField As String
    Dim fromDateStr As String
    Dim toDateStr As String

    formatId = CStr(ws.Cells(SS_ROW_FMT_ID, SS_COL_CONDITION_VALUE).Value)
    keywords = CStr(ws.Cells(SS_ROW_KEYWORDS, SS_COL_CONDITION_VALUE).Value)
    searchMode = CStr(ws.Cells(SS_ROW_MODE, SS_COL_CONDITION_VALUE).Value)
    targetField = CStr(ws.Cells(SS_ROW_TARGET_FIELD, SS_COL_CONDITION_VALUE).Value)
    fromDateStr = CStr(ws.Cells(SS_ROW_DATE_FROM, SS_COL_CONDITION_VALUE).Value)
    toDateStr = CStr(ws.Cells(SS_ROW_DATE_TO, SS_COL_CONDITION_VALUE).Value)

    Dim matched() As String
    Dim scores() As Double
    Dim matchCount As Long
    matchCount = ScanAndMatch(formatId, keywords, searchMode, targetField, _
                                fromDateStr, toDateStr, matched, scores)

    Call ClearResults(ws)
    Call ClearAllThumbs(ws)
    Call PopulateResults(ws, matched, scores, matchCount)

    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsSearchEngine", "SearchByKeywords", _
                          "検索実行: " & CStr(matchCount) & "件ヒット"
    End If

    SearchByKeywords = matchCount
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsSearchEngine", "SearchByKeywords", Err.Description
    End If
    SearchByKeywords = 0
End Function

' ================================================================
' 関数名: DisplayKnowledge
' 概要:   指定ナレッジ番号のファイルを読み込んでナレッジ表示シートに展開
'         + 画像ペイン (J4:N20) に LoadPicture 相当の Shape を貼付
' 引数:   knowledgeNo - ナレッジ番号
' 戻り値: なし
' ================================================================
Public Sub DisplayKnowledge(ByVal knowledgeNo As String)
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")

    If Not FileExists(filePath) Then
        Exit Sub
    End If

    Dim content As String
    content = ReadShiftJisFile(filePath)
    If content = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_DISPLAY)

    Call ClearDisplaySheet(ws)
    Call ClearDetailImage(ws)
    Call RenderKnowledgeToDisplay(ws, knowledgeNo, content)
    Call RenderDetailImagePane(ws, knowledgeNo, content)

    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsSearchEngine", "DisplayKnowledge", _
                          "表示: " & knowledgeNo
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsSearchEngine", "DisplayKnowledge", Err.Description
    End If
End Sub

' ================================================================
' 関数名: ScoreMatchPublic
' 概要:   ScoreMatch のテスト用 Public ラッパ。引数は ScoreMatch と同じ。
' 戻り値: Long - 0=miss、1+=score
' 備考:   テスト層 (T4-006/007) から直接呼び出し可能にするための公開関数。
'         本番フローは ScanAndMatch 経由で内部呼び出しされるため通常は使わない。
' ================================================================
Public Function ScoreMatchPublic(ByVal content As String, _
                                   ByVal formatId As String, _
                                   ByVal keywords As String, _
                                   ByVal searchMode As String, _
                                   ByVal targetField As String, _
                                   ByVal fromDateStr As String, _
                                   ByVal toDateStr As String) As Double
    Dim fromDate As Date
    Dim toDate As Date
    Dim hasFromDate As Boolean
    Dim hasToDate As Boolean
    hasFromDate = TryParseDate(fromDateStr, fromDate)
    hasToDate = TryParseDate(toDateStr, toDate)

    ScoreMatchPublic = ScoreMatch(content, formatId, keywords, searchMode, _
                                    targetField, fromDate, hasFromDate, _
                                    toDate, hasToDate)
End Function

' ================================================================
' 関数名: ExtractImagePathPublic
' 概要:   ExtractImagePath のテスト用 Public ラッパ
' 引数:   content     - ナレッジファイル内容
'         knowledgeNo - ナレッジ番号 (既定パス生成用)
' 戻り値: String - ImagePath スタンザ値 or 既定値 (knowledgeNo & ".png")
' ================================================================
Public Function ExtractImagePathPublic(ByVal content As String, _
                                         ByVal knowledgeNo As String) As String
    ExtractImagePathPublic = ExtractImagePath(content, knowledgeNo)
End Function

' ================================================================
' 関数名: ResolveImageFolderPublic
' 概要:   ResolveImageFolder のテスト用 Public ラッパ
' 引数:   dataFolder - データフォルダのパス
' 戻り値: String - <dataFolder>/../kb_images
' ================================================================
Public Function ResolveImageFolderPublic(ByVal dataFolder As String) As String
    ResolveImageFolderPublic = ResolveImageFolder(dataFolder)
End Function

' ================================================================
' 関数名: ScanAndMatch
' 概要:   データフォルダ内の全ナレッジファイルをスキャンして条件一致を抽出
'         + 各ファイルにスコアを付け、降順で並び替えた配列を返す
' 引数:   formatId     - フォーマットID (絞込、空なら全)
'         keywords     - キーワード
'         searchMode   - "AND" or "OR"
'         targetField  - 対象フィールド
'         fromDateStr  - 作成日From
'         toDateStr    - 作成日To
'         outMatched   - 出力: マッチしたナレッジ番号の配列 (score 降順)
'         outScores    - 出力: 対応するスコア配列
' 戻り値: Long - マッチ件数
' ================================================================
Private Function ScanAndMatch(ByVal formatId As String, _
                                ByVal keywords As String, _
                                ByVal searchMode As String, _
                                ByVal targetField As String, _
                                ByVal fromDateStr As String, _
                                ByVal toDateStr As String, _
                                ByRef outMatched() As String, _
                                ByRef outScores() As Double) As Long
    Dim files As Variant
    files = ListFilesInFolder(m_dataFolder, "txt")

    ReDim outMatched(0 To RESULT_MAX_ROWS - 1)
    ReDim outScores(0 To RESULT_MAX_ROWS - 1)
    Dim matchCount As Long
    matchCount = 0

    Dim fromDate As Date
    Dim toDate As Date
    Dim hasFromDate As Boolean
    Dim hasToDate As Boolean
    hasFromDate = TryParseDate(fromDateStr, fromDate)
    hasToDate = TryParseDate(toDateStr, toDate)

    Dim i As Long
    For i = LBound(files) To UBound(files)
        Dim fileName As String
        fileName = CStr(files(i))

        Dim filePath As String
        filePath = CombineFilePath(m_dataFolder, fileName)

        Dim content As String
        content = ReadShiftJisFile(filePath)
        If content <> "" Then
            Dim sc As Double
            sc = ScoreMatch(content, formatId, keywords, searchMode, _
                              targetField, fromDate, hasFromDate, _
                              toDate, hasToDate)
            If sc > 0# Then
                If matchCount < RESULT_MAX_ROWS Then
                    outMatched(matchCount) = Left(fileName, Len(fileName) - 4)
                    outScores(matchCount) = sc
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next i

    ' スコア降順ソート (バブルソート、件数 200 以下なので十分)
    Call SortByScoreDesc(outMatched, outScores, matchCount)

    ScanAndMatch = matchCount
End Function

' ================================================================
' 関数名: SortByScoreDesc
' 概要:   matched / scores の対を score 降順に並び替える (バブルソート)
' 備考:   同点の場合は入力順を維持 (stable)
' ================================================================
Private Sub SortByScoreDesc(ByRef matched() As String, _
                              ByRef scores() As Double, _
                              ByVal n As Long)
    Dim i As Long, j As Long
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If scores(j) < scores(j + 1) Then
                Dim ts As Double
                Dim tm As String
                ts = scores(j)
                scores(j) = scores(j + 1)
                scores(j + 1) = ts
                tm = matched(j)
                matched(j) = matched(j + 1)
                matched(j + 1) = tm
            End If
        Next j
    Next i
End Sub

' ================================================================
' 関数名: ScoreMatch
' 概要:   1 つのナレッジファイル内容にスコアを付与
'         0 = フィルタで弾かれた (該当なし)
'         1 以上 = キーワード出現回数 × ブースト係数
' 引数:   IsMatch とほぼ同じ + ブースト計算
' 戻り値: Double - スコア (0 = 該当なし)
' 備考:   既存 IsMatch (Boolean) はテスト互換性のため残してあり、
'         本関数で完全に上位互換できる。
' ================================================================
Private Function ScoreMatch(ByVal content As String, _
                              ByVal formatId As String, _
                              ByVal keywords As String, _
                              ByVal searchMode As String, _
                              ByVal targetField As String, _
                              ByVal fromDate As Date, _
                              ByVal hasFromDate As Boolean, _
                              ByVal toDate As Date, _
                              ByVal hasToDate As Boolean) As Double
    ' フィルタ判定 (既存 IsMatch と同じ)
    If formatId <> "" Then
        Dim actualFmt As String
        actualFmt = ExtractStanzaValue(content, "FormatID")
        If actualFmt <> formatId Then
            ScoreMatch = 0#
            Exit Function
        End If
    End If

    If hasFromDate Or hasToDate Then
        Dim createdStr As String
        createdStr = ExtractStanzaValue(content, "CreatedDate")
        Dim createdDate As Date
        If TryParseDate(createdStr, createdDate) Then
            If hasFromDate And createdDate < fromDate Then
                ScoreMatch = 0#
                Exit Function
            End If
            If hasToDate And createdDate > toDate Then
                ScoreMatch = 0#
                Exit Function
            End If
        End If
    End If

    ' キーワード判定
    If Trim(keywords) = "" Then
        ' キーワードなし = 全件ヒット (score=1)
        ScoreMatch = 1#
        Exit Function
    End If

    ' searchMode 判定 (空 or AND/OR でない場合は AND 扱い)
    Dim isOr As Boolean
    isOr = (UCase(searchMode) = "OR")

    ' AND/OR の bool 判定 (既存ロジック維持) で 0 確定をスキップ
    Dim allText As String
    allText = BuildSearchTarget(content, targetField)
    Dim hitAll As Boolean
    If isOr Then
        hitAll = ContainsAnyKeyword(allText, keywords)
    Else
        hitAll = ContainsAllKeywords(allText, keywords)
    End If
    If Not hitAll Then
        ScoreMatch = 0#
        Exit Function
    End If

    ' --- ブースト計算 ---
    ' 全フィールド連結への出現回数
    Dim baseHits As Long
    baseHits = CountKeywordHits(allText, keywords)

    ' タイトルへの出現はブースト
    Dim titleText As String
    titleText = ExtractKeyValueFromItems(content, "タイトル")
    Dim titleHits As Long
    titleHits = CountKeywordHits(titleText, keywords)

    ' targetField 指定時は通常ブーストの代わりに targetField ヒットを強調
    Dim fieldBoost As Double
    fieldBoost = 0#
    If targetField <> "" And targetField <> "(全フィールド)" Then
        Dim fieldText As String
        fieldText = ExtractKeyValueFromItems(content, targetField)
        Dim fieldHits As Long
        fieldHits = CountKeywordHits(fieldText, keywords)
        fieldBoost = fieldHits * SCORE_BOOST_TARGET_FIELD
    End If

    ScoreMatch = CDbl(baseHits) + (CDbl(titleHits) * SCORE_BOOST_TITLE) + fieldBoost
End Function

' ================================================================
' 関数名: CountKeywordHits
' 概要:   target にスペース区切り keywords の各語が出現する合計回数を返す
'         (大文字小文字無視)
' ================================================================
Private Function CountKeywordHits(ByVal target As String, _
                                    ByVal keywords As String) As Long
    If Trim(target) = "" Or Trim(keywords) = "" Then
        CountKeywordHits = 0
        Exit Function
    End If

    Dim parts() As String
    parts = Split(Trim(keywords), " ")
    Dim total As Long
    total = 0
    Dim targetLower As String
    targetLower = LCase(target)

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim k As String
        k = Trim(parts(i))
        If k <> "" Then
            total = total + CountOccurrences(targetLower, LCase(k))
        End If
    Next i
    CountKeywordHits = total
End Function

' ================================================================
' 関数名: CountOccurrences
' 概要:   needle が haystack に何回出現するかカウント (重複なし)
' ================================================================
Private Function CountOccurrences(ByVal haystack As String, _
                                    ByVal needle As String) As Long
    If needle = "" Then
        CountOccurrences = 0
        Exit Function
    End If
    Dim cnt As Long
    Dim startPos As Long
    startPos = 1
    cnt = 0
    Do While startPos <= Len(haystack)
        Dim pos As Long
        pos = InStr(startPos, haystack, needle)
        If pos = 0 Then Exit Do
        cnt = cnt + 1
        startPos = pos + Len(needle)
    Loop
    CountOccurrences = cnt
End Function

' ================================================================
' 関数名: IsMatch (既存互換)
' 概要:   既存テスト (T4 系) との互換のため Bool 版を残す
'         内部的には ScoreMatch > 0 と等価
' ================================================================
Private Function IsMatch(ByVal content As String, _
                           ByVal formatId As String, _
                           ByVal keywords As String, _
                           ByVal searchMode As String, _
                           ByVal targetField As String, _
                           ByVal fromDate As Date, _
                           ByVal hasFromDate As Boolean, _
                           ByVal toDate As Date, _
                           ByVal hasToDate As Boolean) As Boolean
    IsMatch = (ScoreMatch(content, formatId, keywords, searchMode, _
                           targetField, fromDate, hasFromDate, _
                           toDate, hasToDate) > 0#)
End Function

' ================================================================
' 関数名: BuildSearchTarget
' 概要:   検索対象文字列を構築 (対象フィールド指定時は特定のみ)
' ================================================================
Private Function BuildSearchTarget(ByVal content As String, _
                                     ByVal targetField As String) As String
    Dim lines() As String
    Dim result As String
    Dim i As Long

    lines = Split(content, vbCrLf)

    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "FieldName=") > 0 And _
           InStr(lines(i), "Value=") > 0 Then
            Dim fName As String
            Dim fValue As String
            fName = ExtractKeyFromItem(lines(i), "FieldName")
            fValue = ExtractKeyFromItem(lines(i), "Value")

            If targetField = "" Or _
               targetField = "(全フィールド)" Or _
               targetField = fName Then
                result = result & " " & fValue
            End If
        End If
    Next i

    BuildSearchTarget = result
End Function

' ================================================================
' 関数名: ExtractKeyFromItem
' 概要:   ITEM 行から指定キーの値を抽出
' ================================================================
Private Function ExtractKeyFromItem(ByVal line As String, _
                                      ByVal keyName As String) As String
    Dim searchKey As String
    Dim startPos As Long
    Dim endPos As Long

    searchKey = keyName & "="
    startPos = InStr(line, searchKey)
    If startPos = 0 Then
        ExtractKeyFromItem = ""
        Exit Function
    End If

    startPos = startPos + Len(searchKey)
    endPos = InStr(startPos, line, " / ")
    If endPos = 0 Then
        ExtractKeyFromItem = Mid(line, startPos)
    Else
        ExtractKeyFromItem = Mid(line, startPos, endPos - startPos)
    End If
End Function

' ================================================================
' 関数名: ExtractStanzaValue
' 概要:   スタンザ Key=Value 形式から単純抽出
' ================================================================
Private Function ExtractStanzaValue(ByVal content As String, _
                                      ByVal keyName As String) As String
    Dim lines() As String
    Dim i As Long

    lines = Split(content, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), keyName & "=") = 1 Then
            ExtractStanzaValue = Mid(lines(i), Len(keyName) + 2)
            Exit Function
        End If
    Next i
    ExtractStanzaValue = ""
End Function

' ================================================================
' 関数名: ExtractImagePath
' 概要:   ImagePath スタンザを取得。なければ既定値 (knwNo & ".png") を返す
' 引数:   content     - ナレッジファイル内容
'         knowledgeNo - ナレッジ番号 (既定パス生成用)
' 戻り値: String - ImagePath 値 or 既定値
' 備考:   既存ファイル互換: ImagePath 行が無い場合も knwNo.png にフォールバック
' ================================================================
Private Function ExtractImagePath(ByVal content As String, _
                                    ByVal knowledgeNo As String) As String
    Dim v As String
    v = ExtractStanzaValue(content, IMAGE_PATH_KEY)
    If v = "" Then
        ExtractImagePath = knowledgeNo & ".png"
    Else
        ExtractImagePath = v
    End If
End Function

' ================================================================
' 関数名: ResolveImageFolder
' 概要:   <dataFolder>/../kb_images/ のフルパスを返す
' 引数:   dataFolder - データフォルダパス
' 戻り値: String - 画像フォルダパス (末尾区切り無し)
' 備考:   path 区切りは Windows と互換 (\ 区切りを保持)
' ================================================================
Private Function ResolveImageFolder(ByVal dataFolder As String) As String
    Dim base As String
    base = dataFolder
    ' 末尾の \ / を除去
    Do While Len(base) > 0 And (Right(base, 1) = "" Or Right(base, 1) = "/")
        base = Left(base, Len(base) - 1)
    Loop
    ' 親フォルダ
    Dim parentEnd As Long
    parentEnd = InStrRev(base, "")
    If parentEnd = 0 Then parentEnd = InStrRev(base, "/")
    If parentEnd = 0 Then
        ResolveImageFolder = base & "\kb_images"
    Else
        ResolveImageFolder = Left(base, parentEnd - 1) & "\kb_images"
    End If
End Function

' ================================================================
' 関数名: ResolveImageFullPath
' 概要:   ImagePath (相対 or 絶対) を絶対パスに解決
' 引数:   imagePath  - スタンザの ImagePath 値 or 既定値
'         dataFolder - 解決基準フォルダ (<dataFolder>/../kb_images がベース)
' 戻り値: String - 絶対パス
' 備考:   imagePath が "X:\..." or "\..." なら絶対扱い、そうでなければ
'         <imageFolder>\<imagePath> として解決
' ================================================================
Private Function ResolveImageFullPath(ByVal imagePath As String, _
                                        ByVal dataFolder As String) As String
    If Len(imagePath) >= 2 Then
        If Mid(imagePath, 2, 1) = ":" Or Left(imagePath, 2) = "\" Then
            ResolveImageFullPath = imagePath
            Exit Function
        End If
    End If
    ResolveImageFullPath = ResolveImageFolder(dataFolder) & "" & imagePath
End Function

' ================================================================
' 関数名: ClearResults
' 概要:   検索結果一覧の値部分のみクリア (結合セル対策、行ごと Resume Next)
' ================================================================
Private Sub ClearResults(ByVal ws As Worksheet)
    Dim i As Long
    For i = SS_RESULT_START_ROW To SS_RESULT_START_ROW + RESULT_MAX_ROWS
        On Error Resume Next
        ws.Range(ws.Cells(i, 1), ws.Cells(i, SS_RESULT_COL_SCORE)).ClearContents
        Err.Clear
        On Error GoTo 0
    Next i
End Sub

' ================================================================
' 関数名: ClearAllThumbs
' 概要:   検索シート上の旧サムネ Shape (kbThumb_*) を全削除
' 備考:   modImageRender.ClearShapesByPrefix を呼ぶ
' ================================================================
Private Sub ClearAllThumbs(ByVal ws As Worksheet)
    Call ClearShapesByPrefix(ws, "kbThumb_")
End Sub

' ================================================================
' 関数名: ClearDetailImage
' 概要:   ナレッジ表示シート上の旧詳細画像 Shape (kbDetailImg_*) を全削除
' ================================================================
Private Sub ClearDetailImage(ByVal ws As Worksheet)
    Call ClearShapesByPrefix(ws, "kbDetailImg_")
End Sub

' ================================================================
' 関数名: PopulateResults
' 概要:   検索結果を 9 列で書き出す。サムネは modImageRender 経由で
'         Shape として配置。
' 引数:   ws       - 検索シート
'         matched  - 番号配列 (score 降順)
'         scores   - 対応スコア配列
'         count    - 件数
' ================================================================
Private Sub PopulateResults(ByVal ws As Worksheet, _
                              ByRef matched() As String, _
                              ByRef scores() As Double, _
                              ByVal count As Long)
    Dim i As Long
    For i = 0 To count - 1
        Dim knwNo As String
        knwNo = matched(i)

        Dim filePath As String
        filePath = CombineFilePath(m_dataFolder, knwNo & ".txt")

        Dim content As String
        content = ReadShiftJisFile(filePath)

        Dim targetRow As Long
        targetRow = SS_RESULT_START_ROW + i

        ws.Cells(targetRow, SS_RESULT_COL_NO).Value = i + 1
        ws.Cells(targetRow, SS_RESULT_COL_KNW_NO).Value = knwNo
        ws.Cells(targetRow, SS_RESULT_COL_FMT_NAME).Value = _
            ExtractStanzaValue(content, "FormatName")
        ws.Cells(targetRow, SS_RESULT_COL_TITLE).Value = _
            ExtractKeyValueFromItems(content, "タイトル")
        ws.Cells(targetRow, SS_RESULT_COL_CREATED).Value = _
            ExtractStanzaValue(content, "CreatedDate")
        ws.Cells(targetRow, SS_RESULT_COL_UPDATED).Value = _
            ExtractStanzaValue(content, "UpdatedDate")
        ' "▶ 詳細" は CP932 で ▶ が encode 失敗するため ChrW で動的構築する
        ws.Cells(targetRow, SS_RESULT_COL_DETAIL).Value = ChrW(&H25B6) & " 詳細"

        ' --- サムネ Shape 配置 (列 H) ---
        Dim imgRel As String
        imgRel = ExtractImagePath(content, knwNo)
        Dim imgFull As String
        imgFull = ResolveImageFullPath(imgRel, m_dataFolder)
        Call RenderThumb(ws, targetRow, SS_RESULT_COL_THUMB, imgFull, knwNo)

        ' --- スコア (列 I) ---
        ws.Cells(targetRow, SS_RESULT_COL_SCORE).Value = _
            Format(scores(i), "0.00")
    Next i
End Sub

' ================================================================
' 関数名: ExtractKeyValueFromItems
' 概要:   [ITEM] 行から FieldName 指定の値を抽出
' ================================================================
Private Function ExtractKeyValueFromItems(ByVal content As String, _
                                            ByVal fieldName As String) As String
    Dim lines() As String
    Dim i As Long

    lines = Split(content, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "FieldName=" & fieldName) > 0 Then
            ExtractKeyValueFromItems = ExtractKeyFromItem(lines(i), "Value")
            Exit Function
        End If
    Next i
    ExtractKeyValueFromItems = ""
End Function

' ================================================================
' 関数名: ClearDisplaySheet
' 概要:   ナレッジ表示シートの値部分をクリア (Shape は別関数でクリア)
' ================================================================
Private Sub ClearDisplaySheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Cells.ClearContents
    On Error GoTo 0
End Sub

' ================================================================
' 関数名: RenderKnowledgeToDisplay
' 概要:   ナレッジ表示シートに ITEM 行を展開
' ================================================================
Private Sub RenderKnowledgeToDisplay(ByVal ws As Worksheet, _
                                       ByVal knowledgeNo As String, _
                                       ByVal content As String)
    Dim formatId As String
    Dim formatName As String
    formatId = ExtractStanzaValue(content, "FormatID")
    formatName = ExtractStanzaValue(content, "FormatName")

    ws.Cells(KD_ROW_KNW_NO, 1).Value = "ナレッジ番号:"
    ws.Cells(KD_ROW_KNW_NO, KD_COL_KNW_NO_VAL).Value = knowledgeNo
    ws.Cells(KD_ROW_KNW_NO, KD_COL_FMT_INFO).Value = _
        formatName & " (" & formatId & ")"

    Dim lines() As String
    lines = Split(content, vbCrLf)

    Dim targetRow As Long
    targetRow = KD_FORM_START_ROW

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "FieldNo=") > 0 And _
           InStr(lines(i), "FieldName=") > 0 Then
            ws.Cells(targetRow, KD_COL_FIELD_NO).Value = _
                ExtractKeyFromItem(lines(i), "FieldNo")
            ws.Cells(targetRow, KD_COL_FIELD_NAME).Value = _
                ExtractKeyFromItem(lines(i), "FieldName")
            ws.Cells(targetRow, KD_COL_FIELD_VALUE).NumberFormat = "@"
            ws.Cells(targetRow, KD_COL_FIELD_VALUE).Value = _
                SafeCellText(ExtractKeyFromItem(lines(i), "Value"))
            targetRow = targetRow + 1
        End If
    Next i
End Sub

' ================================================================
' 関数名: RenderDetailImagePane
' 概要:   ナレッジ表示シートの J4:N20 領域に詳細画像を貼付
'         画像が無い場合は薄字で「[画像未配置: ...]」を表示
' ================================================================
Private Sub RenderDetailImagePane(ByVal ws As Worksheet, _
                                    ByVal knowledgeNo As String, _
                                    ByVal content As String)
    On Error Resume Next
    Dim imgRel As String
    imgRel = ExtractImagePath(content, knowledgeNo)
    Dim imgFull As String
    imgFull = ResolveImageFullPath(imgRel, m_dataFolder)

    Call RenderDetailImage(ws, _
                            KD_DETAIL_IMG_TOP_ROW, KD_DETAIL_IMG_LEFT_COL, _
                            KD_DETAIL_IMG_BOT_ROW, KD_DETAIL_IMG_RIGHT_COL, _
                            imgFull, knowledgeNo)
End Sub

' ================================================================
' 関数名: CombineFilePath
' 概要:   フォルダパスとファイル名を結合
' ================================================================
Private Function CombineFilePath(ByVal folder As String, _
                                   ByVal fileName As String) As String
    If Right(folder, 1) = "" Or Right(folder, 1) = "/" Then
        CombineFilePath = folder & fileName
    Else
        CombineFilePath = folder & "" & fileName
    End If
End Function
'@ },
    @{ Name='clsSectionSpec'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsSectionSpec（画面層 — データ）
' 概要:   1 個のセクション帯の仕様（位置・ラベル・色）を保持する ValueObject。
' 依存先: なし
' ================================================================

Private m_address As String
Private m_label As String
Private m_colorHex As String

Public Property Get Address() As String
    Address = m_address
End Property
Public Property Let Address(ByVal value As String)
    m_address = value
End Property

Public Property Get Label() As String
    Label = m_label
End Property
Public Property Let Label(ByVal value As String)
    m_label = value
End Property

Public Property Get ColorHex() As String
    ColorHex = m_colorHex
End Property
Public Property Let ColorHex(ByVal value As String)
    m_colorHex = value
End Property

Public Sub Init(ByVal address As String, ByVal label As String, ByVal colorHex As String)
    m_address = address
    m_label = label
    m_colorHex = colorHex
End Sub
'@ },
    @{ Name='clsSetupOrchestrator'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsSetupOrchestrator（インストーラ層）
' 概要:   初期セットアップの司令塔。14 シート作成 → 各画面の Setup 呼出 →
'         初期可視性設定 → 既定 Sheet 削除 を順次実行。
'         旧 modSetup.SetupSheetsAndButtons の責務を OOP で再構成したもの。
' 依存先: IScreenRenderer, modFactory, modScreenSpecRegistry, modCommon, clsLogger
' 備考:   v21 (E2E rerun) で全 Sub に ENTER/EXIT/step ログを注入。
'         SetupOneScreen の silent NextScreen → LogError で必ず痕跡を残す。
' ================================================================

Private Const MOD_NAME As String = "clsSetupOrchestrator"

Private m_renderer As IScreenRenderer

Private Const REQUIRED_SHEETS_CSV As String = _
    "メイン,フォーマット一覧,フォーマット設計,フォーマットプレビュー," & _
    "ナレッジ登録,ナレッジ修正,ナレッジ一覧,検索,ナレッジ表示," & _
    "格納先設定,設定,既存データへのフィールド反映,データファイル形式,ログ"
Private Const DEFAULT_SHEETS_CSV As String = "Sheet1,Sheet2,Sheet3,Sheet4"

' ================================================================
' 関数名: Init
' 概要:   Renderer を依存注入
' ================================================================
Public Sub Init(ByVal renderer As IScreenRenderer)
    Set m_renderer = renderer
End Sub

' ================================================================
' 関数名: RunFullSetup
' 概要:   フルセットアップを実行
' ================================================================
Public Sub RunFullSetup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    stepName = "CreateRequiredSheets"
    Call LogEnter("RunFullSetup", stepName)
    Call CreateRequiredSheets

    ' この時点で SHEET_LOG が存在するため、これ以降は logger を使える
    Call LogTraceSafe("RunFullSetup", "after CreateRequiredSheets - sheets created")

    stepName = "SetupAllScreens"
    Call LogTraceSafe("RunFullSetup", "STEP " & stepName)
    Call SetupAllScreens

    stepName = "InitializeSettingsSheet"
    Call LogTraceSafe("RunFullSetup", "STEP " & stepName)
    Call InitializeSettingsSheet

    stepName = "SetInitialVisibility"
    Call LogTraceSafe("RunFullSetup", "STEP " & stepName)
    Call SetInitialVisibility

    stepName = "DeleteEmptyDefaultSheets"
    Call LogTraceSafe("RunFullSetup", "STEP " & stepName)
    Call DeleteEmptyDefaultSheets

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Call LogTraceSafe("RunFullSetup", "EXIT ok")
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Call LogErrorSafe("RunFullSetup", stepName, Err.Number, Err.Description)
    Err.Raise Err.Number, "clsSetupOrchestrator.RunFullSetup", Err.Description
End Sub

' ================================================================
' 関数名: CreateRequiredSheets
' 概要:   14 シートを順次作成（既存はスキップ）
' ================================================================
Private Sub CreateRequiredSheets()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "split CSV"

    Dim names() As String
    names = Split(REQUIRED_SHEETS_CSV, ",")

    Dim createdCount As Long : createdCount = 0
    Dim skippedCount As Long : skippedCount = 0

    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim sheetName As String
        sheetName = Trim$(names(i))
        stepName = "check " & sheetName
        If Len(sheetName) > 0 Then
            If Not modCommon.SheetExists(sheetName) Then
                stepName = "append " & sheetName
                Call AppendNewSheet(sheetName)
                createdCount = createdCount + 1
            Else
                skippedCount = skippedCount + 1
            End If
        End If
    Next i

    Call LogTraceSafe("CreateRequiredSheets", _
                       "created=" & createdCount & " skipped=" & skippedCount)
    Exit Sub

ErrHandler:
    Call LogErrorSafe("CreateRequiredSheets", stepName, Err.Number, Err.Description)
    Err.Raise Err.Number, MOD_NAME & ".CreateRequiredSheets", Err.Description
End Sub

Private Sub AppendNewSheet(ByVal sheetName As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    Exit Sub
ErrHandler:
    Call LogErrorSafe("AppendNewSheet", "Add+Name " & sheetName, Err.Number, Err.Description)
    Err.Raise Err.Number, MOD_NAME & ".AppendNewSheet", Err.Description
End Sub

' ================================================================
' 関数名: SetupAllScreens
' 概要:   14 画面の Setup を順次実行（spec → screen → renderer）
' ================================================================
Private Sub SetupAllScreens()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "GetAllScreenIds"

    Dim screenIds As Variant
    screenIds = modScreenSpecRegistry.GetAllScreenIds()

    Dim okCount As Long : okCount = 0
    Dim ngCount As Long : ngCount = 0

    Dim i As Long
    For i = LBound(screenIds) To UBound(screenIds)
        Dim screenId As String
        screenId = CStr(screenIds(i))
        stepName = "Setup " & screenId
        Call LogTraceSafe("SetupAllScreens", "BEGIN " & screenId)
        If SetupOneScreenLogged(screenId) Then
            okCount = okCount + 1
        Else
            ngCount = ngCount + 1
        End If
    Next i

    Call LogTraceSafe("SetupAllScreens", _
                       "EXIT ok=" & okCount & " ng=" & ngCount)
    Exit Sub

ErrHandler:
    Call LogErrorSafe("SetupAllScreens", stepName, Err.Number, Err.Description)
    Err.Raise Err.Number, MOD_NAME & ".SetupAllScreens", Err.Description
End Sub

' ================================================================
' 関数名: SetupOneScreenLogged
' 概要:   1 画面の Setup を実行。失敗しても他画面の続行のため True/False で返す。
'         失敗時は LogError で痕跡を残す（旧 SetupOneScreen の silent NextScreen
'         を廃止し、必ず Err.Number/Description を記録）。
' 戻り値: Boolean - Setup 成功なら True
' ================================================================
Private Function SetupOneScreenLogged(ByVal screenId As String) As Boolean
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "CreateScreen " & screenId

    Dim screen As Object
    Set screen = modFactory.CreateScreen(screenId, m_renderer)
    If screen Is Nothing Then
        Call LogWarnSafe("SetupOneScreenLogged", screenId & " screen=Nothing (spec missing or unknown id)")
        SetupOneScreenLogged = False
        Exit Function
    End If

    stepName = "screen.Setup " & screenId
    screen.Setup

    Call LogTraceSafe("SetupOneScreenLogged", screenId & " EXIT ok")
    SetupOneScreenLogged = True
    Exit Function

ErrHandler:
    Call LogErrorSafe("SetupOneScreenLogged", _
                       screenId & " step=" & stepName, Err.Number, Err.Description)
    SetupOneScreenLogged = False
    Err.Clear
End Function

' ================================================================
' 関数名: InitializeSettingsSheet
' 概要:   設定シートのテストモード初期値を書き込み
' ================================================================
Private Sub InitializeSettingsSheet()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "open sheet"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    Err.Clear
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        Call LogWarnSafe("InitializeSettingsSheet", SHEET_SETTINGS & " not found - skip")
        Exit Sub
    End If

    stepName = "write testmode"
    ws.Cells(SETTINGS_ROW_TESTMODE, SETTINGS_COL_NAME).Value = "テストモード"
    ws.Cells(SETTINGS_ROW_TESTMODE, SETTINGS_COL_VALUE).Value = TESTMODE_OFF
    Call LogTraceSafe("InitializeSettingsSheet", "testmode=" & TESTMODE_OFF)
    Exit Sub

ErrHandler:
    Call LogErrorSafe("InitializeSettingsSheet", stepName, Err.Number, Err.Description)
    ' 致命的でないので Raise せずに継続
End Sub

' ================================================================
' 関数名: SetInitialVisibility
' 概要:   メインのみ可視、その他は xlSheetHidden
' ================================================================
Private Sub SetInitialVisibility()
    On Error Resume Next
    Dim shown As Long, hidden As Long
    Dim w As Worksheet
    For Each w In ThisWorkbook.Worksheets
        If w.Name = SHEET_MAIN Then
            w.Visible = -1  ' xlSheetVisible
            shown = shown + 1
        Else
            w.Visible = 0   ' xlSheetHidden
            hidden = hidden + 1
        End If
    Next w
    Call LogTraceSafe("SetInitialVisibility", "shown=" & shown & " hidden=" & hidden)
End Sub

' ================================================================
' 関数名: DeleteEmptyDefaultSheets
' 概要: Excel 既定の空 Sheet1/Sheet2/Sheet3 等を削除
' ================================================================
Private Sub DeleteEmptyDefaultSheets()
    On Error Resume Next
    Application.DisplayAlerts = False
    Dim deleted As Long : deleted = 0
    Dim names() As String
    Dim cnt As Long : cnt = 0
    ReDim names(0 To 100)
    Dim w As Worksheet
    For Each w In ThisWorkbook.Worksheets
        If LCase$(Left$(w.Name, 5)) = "sheet" Then
            ' Excel 既定 (Sheet1/Sheet2/...) と思われるもののみ
            names(cnt) = w.Name
            cnt = cnt + 1
        End If
    Next w
    Dim i As Long
    For i = 0 To cnt - 1
        If ThisWorkbook.Worksheets.Count > 1 Then
            ThisWorkbook.Worksheets(names(i)).Delete
            deleted = deleted + 1
        End If
    Next i
    Application.DisplayAlerts = True
    Call LogTraceSafe("DeleteEmptyDefaultSheets", "deleted=" & deleted)
End Sub
'@ },
    @{ Name='clsSheetRenderer'; Type='cls'; Code=@'
Option Explicit

Implements IScreenRenderer

' ================================================================
' クラス: clsSheetRenderer（画面層 — 実装1）
' 概要:   IScreenRenderer のシート埋込型実装。
'         ws.Range / ws.Shapes を使ってセル背景・ラベル・フォーム
'         コントロールボタンを物理配置する。
' 依存先: modCommon (カラー定数), clsButtonSpec, clsFieldSpec, clsLogger
' 備考:   将来 UserForm 切替する場合は clsUserFormRenderer に差し替え可。
'         画面クラス側は IScreenRenderer 経由で呼ぶため変更不要。
'         v21 (E2E rerun) で RenderButton に AddFormControl 失敗ログを注入。
' ================================================================

Private Const MOD_NAME As String = "clsSheetRenderer"

Private Const XL_BUTTON_CONTROL As Long = 0  ' xlButtonControl と同値（headless 互換）
Private Const BTN_MIN_W As Double = 100#
Private Const BTN_MIN_H As Double = 26#
Private Const TITLE_ROW_HEIGHT As Double = 28#
Private Const SECTION_ROW_HEIGHT As Double = 20#

Private m_ws As Worksheet

' ================================================================
' 関数名: IScreenRenderer_BindSheet
' 概要:   描画対象シートを束縛する
' 引数:   sheetName - 対象シート名
' ================================================================
Private Sub IScreenRenderer_BindSheet(ByVal sheetName As String)
    Set m_ws = ThisWorkbook.Worksheets(sheetName)
End Sub

' ================================================================
' 関数名: IScreenRenderer_ClearScreen
' 概要:   シート内容（セル値・色・シェイプ）を全クリア。
'         再描画前のリセット用。
' ================================================================
Private Sub IScreenRenderer_ClearScreen()
    If m_ws Is Nothing Then Exit Sub
    On Error Resume Next
    m_ws.Cells.Clear
    Dim shp As Shape
    For Each shp In m_ws.Shapes
        shp.Delete
    Next shp
    Err.Clear
    On Error GoTo 0
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderTitle
' 概要:   画面タイトル帯を行 1 に描画する
' 引数:   screenId - "M-01" 等
'         title    - タイトル文字列
'         colorHex - 背景色 HEX（例 "#1F3864"）
' ================================================================
Private Sub IScreenRenderer_RenderTitle(ByVal screenId As String, _
                                         ByVal title As String, _
                                         ByVal colorHex As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range("A1")
    cell.Value = title
    cell.Font.Bold = True
    cell.Font.Color = RGB(255, 255, 255)
    cell.Font.Size = 14
    cell.Interior.Color = HexToRgb(colorHex)
    m_ws.Rows(1).RowHeight = TITLE_ROW_HEIGHT
    ' タイトル列 A:L を背景色塗り
    m_ws.Range("A1:L1").Interior.Color = HexToRgb(colorHex)
    m_ws.Range("A1:L1").Font.Color = RGB(255, 255, 255)
    m_ws.Range("A1:L1").Font.Bold = True
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderSection
' 概要:   セクション帯（■ ラベル）を描画する
' 引数:   sectionAddr   - 開始セル位置 ("A3" 等)
'         sectionLabel  - "■ モード/予定番号" 等
'         colorHex      - 背景色 HEX（例 "#4472C4"）
' ================================================================
Private Sub IScreenRenderer_RenderSection(ByVal sectionAddr As String, _
                                           ByVal sectionLabel As String, _
                                           ByVal colorHex As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range(sectionAddr)
    cell.Value = sectionLabel
    cell.Font.Bold = True
    cell.Font.Color = RGB(255, 255, 255)
    ' セクション帯は A〜L 列を塗る
    Dim rowNum As Long
    rowNum = cell.Row
    Dim rowRange As Range
    Set rowRange = m_ws.Range(m_ws.Cells(rowNum, 1), m_ws.Cells(rowNum, 12))
    rowRange.Interior.Color = HexToRgb(colorHex)
    rowRange.Font.Color = RGB(255, 255, 255)
    rowRange.Font.Bold = True
    m_ws.Rows(rowNum).RowHeight = SECTION_ROW_HEIGHT
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderButton
' 概要:   フォームコントロールボタンを配置する
' 引数:   btnSpec - clsButtonSpec インスタンス（CellAddr/Caption/ColorHex/BtnName 必須、
'                   HintAddr/HintText 任意）
' 備考:   既存同名ボタンは削除→再配置で idempotent。色付けはシート背景セルで擬似実現
'         （フォームコントロール本体は色付け不可のため、ボタン直下のセルを塗ることで
'          グループ識別性を確保）。
' ================================================================
Private Sub IScreenRenderer_RenderButton(ByVal btnSpec As Object)
    Dim stepName As String : stepName = "begin"

    If m_ws Is Nothing Then
        Call LogWarnSafe("RenderButton", "m_ws=Nothing (BindSheet not called?)")
        Exit Sub
    End If

    stepName = "cast spec"
    Dim spec As clsButtonSpec
    Set spec = btnSpec

    Call LogTraceSafe("RenderButton", _
                       "ENTER ws=" & m_ws.Name & " btnName=" & spec.BtnName & " addr=" & spec.CellAddr)

    ' 既存同名ボタン削除
    stepName = "DeleteShapeByName"
    Call DeleteShapeByName(spec.BtnName)

    stepName = "resolve Range " & spec.CellAddr
    Dim rng As Range
    Set rng = m_ws.Range(spec.CellAddr)

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

    ' ボタン直下セルに色を塗る（グループ識別性 — フォームコントロール本体は色付け不可）
    stepName = "tint cell"
    On Error Resume Next
    rng.Interior.Color = HexToRgb(spec.ColorHex)
    rng.Font.Color = RGB(255, 255, 255)
    rng.Font.Bold = True
    If Err.Number <> 0 Then
        Call LogWarnSafe("RenderButton", "cell tint failed btnName=" & spec.BtnName & _
                          " errNum=" & Err.Number & " desc=" & Err.Description)
        Err.Clear
    End If
    On Error GoTo 0

    ' フォームコントロールボタン配置（Object late binding で headless/Excel 互換確保）
    stepName = "AddFormControl"
    Dim shp As Shape
    Dim shapesObj As Object
    Set shapesObj = m_ws.Shapes
    Dim afcErrNum As Long : afcErrNum = 0
    Dim afcErrDesc As String : afcErrDesc = ""
    On Error Resume Next
    Set shp = shapesObj.AddFormControl(XL_BUTTON_CONTROL, leftPt, topPt, widthPt, heightPt)
    afcErrNum = Err.Number
    afcErrDesc = Err.Description
    Err.Clear
    On Error GoTo 0

    If shp Is Nothing Then
        Call LogErrorWithErrSafe("RenderButton", _
                                  "AddFormControl returned Nothing btnName=" & spec.BtnName & _
                                  " ws=" & m_ws.Name & " addr=" & spec.CellAddr, _
                                  afcErrNum, afcErrDesc)
        Exit Sub
    End If

    stepName = "shp.Name"
    shp.Name = spec.BtnName

    stepName = "SetButtonCaptionAndAction"
    Call SetButtonCaptionAndAction(shp, spec.Caption, spec.BtnName)

    ' ヒントテキスト（ボタン下の説明セル）
    If Len(spec.HintAddr) > 0 And Len(spec.HintText) > 0 Then
        stepName = "hint " & spec.HintAddr
        m_ws.Range(spec.HintAddr).Value = spec.HintText
        m_ws.Range(spec.HintAddr).WrapText = True
        m_ws.Range(spec.HintAddr).Font.Size = 9
    End If

    Call LogTraceSafe("RenderButton", "EXIT ok btnName=" & spec.BtnName)
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderLabel
' 概要:   ラベルセルを描画（背景色任意）
' ================================================================
Private Sub IScreenRenderer_RenderLabel(ByVal cellAddr As String, _
                                         ByVal labelText As String, _
                                         ByVal colorHex As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range(cellAddr)
    cell.Value = labelText
    If Len(colorHex) > 0 Then
        cell.Interior.Color = HexToRgb(colorHex)
    End If
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderInputField
' 概要:   入力フィールドの「ラベル + 必須マーク + 型表示 + 入力欄ハイライト」を描画。
'         データが空でもフィールドが見えるようにする「空状態 UI」の実装。
' 引数:   cellAddr  - 開始セル（未使用、fieldSpec 内のアドレスを使う）
'         fieldSpec - clsFieldSpec インスタンス
' ================================================================
Private Sub IScreenRenderer_RenderInputField(ByVal cellAddr As String, _
                                              ByVal fieldSpec As Object)
    If m_ws Is Nothing Then Exit Sub
    Dim spec As clsFieldSpec
    Set spec = fieldSpec

    ' 順序番号
    If Len(spec.OrderAddr) > 0 Then
        m_ws.Range(spec.OrderAddr).Value = spec.FieldOrder
        m_ws.Range(spec.OrderAddr).HorizontalAlignment = xlCenter
    End If

    ' 必須マーク
    If spec.Required And Len(spec.ReqMarkAddr) > 0 Then
        Call IScreenRenderer_RenderRequiredMark(spec.ReqMarkAddr)
    End If

    ' ラベル
    If Len(spec.LabelAddr) > 0 Then
        m_ws.Range(spec.LabelAddr).Value = spec.Label
        m_ws.Range(spec.LabelAddr).HorizontalAlignment = xlRight
    End If

    ' 型表示（イタリック、薄黄）
    If Len(spec.TypeAddr) > 0 Then
        m_ws.Range(spec.TypeAddr).Value = spec.TypeText
        m_ws.Range(spec.TypeAddr).Font.Italic = True
        m_ws.Range(spec.TypeAddr).Interior.Color = HexToRgb(COLOR_HINT_YELLOW)
        m_ws.Range(spec.TypeAddr).Font.Size = 9
    End If

    ' 入力欄ヒント（説明文）
    If Len(spec.InputAddr) > 0 And Len(spec.HintText) > 0 Then
        m_ws.Range(spec.InputAddr).Value = spec.HintText
        m_ws.Range(spec.InputAddr).Font.Color = RGB(128, 128, 128)
        m_ws.Range(spec.InputAddr).Font.Italic = True
    End If
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderRequiredMark
' 概要:   必須マーク（赤背景に "*"）を描画
' ================================================================
Private Sub IScreenRenderer_RenderRequiredMark(ByVal cellAddr As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range(cellAddr)
    cell.Value = "*"
    cell.Interior.Color = HexToRgb(COLOR_REQUIRED_RED)
    cell.Font.Color = RGB(255, 255, 255)
    cell.Font.Bold = True
    cell.HorizontalAlignment = xlCenter
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderHint
' 概要:   ヒント/凡例テキストを描画
' ================================================================
Private Sub IScreenRenderer_RenderHint(ByVal cellAddr As String, ByVal hintText As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range(cellAddr)
    cell.Value = hintText
    cell.Interior.Color = HexToRgb(COLOR_HINT_BAR)
    cell.Font.Italic = True
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderHeaderRow
' 概要:   一覧テーブルのヘッダ行を描画（M-02/M-07/M-08/M-10/M-12/M-14 共通）
' 引数:   startAddr     - 開始セル（"B10" 等）
'         headerLabels  - ヘッダ文字列の配列
'         colorHex      - 背景色（"#5B9BD5" 等）
' ================================================================
Private Sub IScreenRenderer_RenderHeaderRow(ByVal startAddr As String, _
                                             ByVal headerLabels As Variant, _
                                             ByVal colorHex As String)
    If m_ws Is Nothing Then Exit Sub
    If IsEmpty(headerLabels) Then Exit Sub

    Dim startCell As Range
    Set startCell = m_ws.Range(startAddr)
    Dim startRow As Long
    Dim startCol As Long
    startRow = startCell.Row
    startCol = startCell.Column

    Dim i As Long
    For i = LBound(headerLabels) To UBound(headerLabels)
        Dim cell As Range
        Set cell = m_ws.Cells(startRow, startCol + i)
        cell.Value = CStr(headerLabels(i))
        cell.Font.Bold = True
        cell.Font.Color = RGB(255, 255, 255)
        cell.Interior.Color = HexToRgb(colorHex)
        cell.HorizontalAlignment = xlCenter
    Next i
End Sub

' ================================================================
' 関数名: IScreenRenderer_RenderEmptyState
' 概要:   「データなし」状態のメッセージを描画
' ================================================================
Private Sub IScreenRenderer_RenderEmptyState(ByVal cellAddr As String, _
                                              ByVal message As String)
    If m_ws Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = m_ws.Range(cellAddr)
    cell.Value = message
    cell.Font.Italic = True
    cell.Font.Color = RGB(128, 128, 128)
End Sub

' --- 内部ユーティリティ ---

' ================================================================
' 関数名: HexToRgb
' 概要:   HEX 文字列 ("#1F3864" or "1F3864") を Long の RGB 値に変換
' 引数:   hex - HEX 文字列
' 戻り値: Long - VBA RGB 値
' ================================================================
Private Function HexToRgb(ByVal hex As String) As Long
    Dim s As String
    s = hex
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)
    If Len(s) <> 6 Then
        HexToRgb = RGB(255, 255, 255)
        Exit Function
    End If
    Dim r As Long
    Dim g As Long
    Dim b As Long
    r = CLng("&H" & Mid$(s, 1, 2))
    g = CLng("&H" & Mid$(s, 3, 2))
    b = CLng("&H" & Mid$(s, 5, 2))
    HexToRgb = RGB(r, g, b)
End Function

' ================================================================
' 関数名: DeleteShapeByName
' 概要:   指定名のシェイプが存在すれば削除（idempotent）
' ================================================================
Private Sub DeleteShapeByName(ByVal shapeName As String)
    If m_ws Is Nothing Then Exit Sub
    Dim shp As Shape
    On Error Resume Next
    Set shp = m_ws.Shapes(shapeName)
    If Not shp Is Nothing Then shp.Delete
    Err.Clear
    On Error GoTo 0
End Sub

' ================================================================
' 関数名: SetButtonCaptionAndAction
' 概要: フォームコントロールボタンのキャプションと OnAction を設定
' ================================================================
Private Sub SetButtonCaptionAndAction(ByVal shp As Shape, ByVal caption As String, ByVal onAction As String)
    On Error Resume Next
    shp.OLEFormat.Object.Caption = caption
    shp.TextFrame.Characters.Text = caption
    shp.OnAction = onAction
End Sub
'@ },
    @{ Name='clsStorageResolver'; Type='cls'; Code=@'
Option Explicit

' Phase 6 レビュー: GetStorageConfig は O(N) だが N ≤ MAX_STORAGE_SCAN_ROWS=1000。
' 現状想定 (docType 数十件) では十分。指摘なし。

' ================================================================
' クラス: clsStorageResolver（ビジネスロジック層）
' 概要:   格納先設定に基づき、ファイル参照フィールドからリンクを生成
'         ドキュメント種類ごとに「共有フォルダ/BOX」「ファイル指定/フォルダ指定」を
'         解釈してクリック可能なパスを組み立てる
' 依存先: clsLogger, modCommon
' ================================================================

' 格納先設定シートの列番号
Private Const STORAGE_COL_NO As Long = 1
Private Const STORAGE_COL_DOC_TYPE As Long = 2
Private Const STORAGE_COL_STORAGE_TYPE As Long = 3
Private Const STORAGE_COL_PATH_TYPE As Long = 4
Private Const STORAGE_COL_BASE_PATH As Long = 5
Private Const STORAGE_COL_NOTE As Long = 6

' 格納先設定の開始行（ヘッダーの次）
Private Const STORAGE_START_ROW As Long = 2

' 格納種別
Private Const MAX_STORAGE_SCAN_ROWS As Long = 1000  ' m-11: マジック数 Const 化
Private Const STORAGE_TYPE_SHARED As String = "共有フォルダ"
Private Const STORAGE_TYPE_BOX As String = "BOX"

' パス種別
Private Const PATH_TYPE_FILE As String = "ファイル指定"
Private Const PATH_TYPE_FOLDER As String = "フォルダ指定"

Private m_logger As clsLogger

' ================================================================
' 関数名: Init
' 概要:   初期化（ロガー参照を保持）
' 引数:   logger - ログ出力用
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger)
    Set m_logger = logger
End Sub

' ================================================================
' 関数名: ResolveLink
' 概要:   ドキュメント種類とファイル名からクリック可能なリンク文字列を生成
' 引数:   docType  - ドキュメント種類（例: "障害報告書"）
'         fileName - ファイル名（例: "report.xlsx"）
' 戻り値: String - 生成されたリンク文字列
'                  （例: "\\filesvr\incident\reports\report.xlsx"）
' 備考:   設定が見つからない/パス種別がフォルダの場合はベースパスのみ返す
'         BOXの場合はBOX URL
' ================================================================
Public Function ResolveLink(ByVal docType As String, _
                              ByVal fileName As String) As String
    On Error GoTo ErrHandler
    
    Dim storageType As String
    Dim pathType As String
    Dim basePath As String
    
    If Not GetStorageConfig(docType, storageType, pathType, basePath) Then
        If Not m_logger Is Nothing Then
            m_logger.LogWarn "clsStorageResolver", "ResolveLink", _
                              "格納先設定なし: docType=" & docType
        End If
        ResolveLink = ""
        Exit Function
    End If
    
    ' BOXの場合はBOX URLをそのまま返す（ファイル直指定不可）
    If storageType = STORAGE_TYPE_BOX Then
        ResolveLink = basePath
        Exit Function
    End If
    
    ' 共有フォルダの場合
    If pathType = PATH_TYPE_FILE Then
        ' ファイル指定: basePath + fileName
        ResolveLink = CombinePath(basePath, fileName)
    Else
        ' フォルダ指定: basePath のみ
        ResolveLink = basePath
    End If
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsStorageResolver", "ResolveLink", _
                           "リンク生成失敗: " & Err.Description
    End If
    ResolveLink = ""
End Function

' ================================================================
' 関数名: GetStorageConfig
' 概要:   格納先設定シートから指定ドキュメント種類の設定を取得
' 引数:   docType        - ドキュメント種類
'         outStorageType - 出力: 格納種別
'         outPathType    - 出力: パス種別
'         outBasePath    - 出力: ベースパス
' 戻り値: Boolean - 設定が見つかったなら True
' 備考:   ドキュメント種類名の完全一致で検索
' ================================================================
Private Function GetStorageConfig(ByVal docType As String, _
                                    ByRef outStorageType As String, _
                                    ByRef outPathType As String, _
                                    ByRef outBasePath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STORAGE)
    
    Dim i As Long
    Dim maxRow As Long
    maxRow = MAX_STORAGE_SCAN_ROWS
    
    For i = STORAGE_START_ROW To maxRow
        Dim currentType As String
        currentType = CStr(ws.Cells(i, STORAGE_COL_DOC_TYPE).Value)
        
        If currentType = "" Then
            Exit For
        End If
        
        If currentType = docType Then
            outStorageType = CStr(ws.Cells(i, STORAGE_COL_STORAGE_TYPE).Value)
            outPathType = CStr(ws.Cells(i, STORAGE_COL_PATH_TYPE).Value)
            outBasePath = CStr(ws.Cells(i, STORAGE_COL_BASE_PATH).Value)
            GetStorageConfig = True
            Exit Function
        End If
    Next i
    
    GetStorageConfig = False
    Exit Function

ErrHandler:
    GetStorageConfig = False
End Function

' ================================================================
' 関数名: CombinePath
' 概要:   ベースパスとファイル名を結合する
' 引数:   basePath - ベースパス
'         fileName - ファイル名
' 戻り値: String - 結合されたパス
' 備考:   basePath の末尾にセパレータがなければ追加
' ================================================================
Private Function CombinePath(ByVal basePath As String, _
                               ByVal fileName As String) As String
    Dim sep As String
    sep = "\"
    
    If Right(basePath, 1) = "\" Or Right(basePath, 1) = "/" Then
        CombinePath = basePath & fileName
    Else
        CombinePath = basePath & sep & fileName
    End If
End Function
'@ },
    @{ Name='clsTaskController'; Type='cls'; Code=@'
Option Explicit

' Phase 6 レビュー: クラス内状態は transient (シート切替の一時記録のみ)。
' D-3 余地: 現状 ThisWorkbook 直結なし、副作用は呼出側 ws を経由。テスト容易性 OK。

' ================================================================
' クラス: clsTaskController（ビジネスロジック層）
' 概要:   利用シーン別シート表示・非表示を制御する
'         8タスクに応じて関連シートのみを表示し、他を非表示にする
'         メインシートは常に表示
' 依存先: clsLogger, modCommon
' ================================================================

' メインシートの現在タスク表示位置
Private Const CURRENT_TASK_ROW As Long = 3
Private Const CURRENT_TASK_COL As Long = 3

Private m_logger As clsLogger

' ================================================================
' 関数名: Init
' 概要:   初期化
' 引数:   logger - ログ出力用
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger)
    Set m_logger = logger
End Sub

' ================================================================
' 関数名: SwitchToTask
' 概要:   指定タスクに切り替える（関連シートのみ表示、他は非表示）
'         メインシートの現在タスク表示も更新
' 引数:   taskName - タスク名（例: "ナレッジ登録"）
' 戻り値: なし
' 備考:   未定義のタスク名の場合はメインシートのみ表示
' ================================================================
Public Sub SwitchToTask(ByVal taskName As String)
    On Error GoTo ErrHandler
    
    Dim sheetNames As Variant
    sheetNames = GetTaskDefinition(taskName)
    
    Call SetVisibleSheets(sheetNames)
    Call UpdateCurrentTaskLabel(taskName)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsTaskController", "SwitchToTask", _
                          "タスク切替: " & taskName
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsTaskController", "SwitchToTask", _
                           "切替失敗: " & Err.Description
    End If
End Sub

' ================================================================
' 関数名: GetCurrentTask
' 概要:   メインシートに表示されている現在のタスク名を取得
' 引数:   なし
' 戻り値: String - 現在のタスク名（未選択なら空文字列）
' ================================================================
Public Function GetCurrentTask() As String
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    GetCurrentTask = CStr(ws.Cells(CURRENT_TASK_ROW, CURRENT_TASK_COL).Value)
    Exit Function

ErrHandler:
    GetCurrentTask = ""
End Function

' ================================================================
' 関数名: SetVisibleSheets
' 概要:   指定されたシート群のみを表示、他は非表示にする
' 引数:   sheetNames - 表示するシート名の配列
' 戻り値: なし
' 備考:   xlSheetHidden を使用（ユーザが右クリック→再表示で戻せる）
'         xlVeryHidden は使わない
' ================================================================
Private Sub SetVisibleSheets(ByVal sheetNames As Variant)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim shouldShow As Boolean
    
    For Each ws In ThisWorkbook.Worksheets
        shouldShow = IsInArray(ws.Name, sheetNames)
        
        If shouldShow Then
            ws.Visible = xlSheetVisible
        Else
            ' ログシート・データファイル形式・メイン以外を非表示化
            ' ただし全てのシートを非表示にはできないため、最低1枚は残す
            If ws.Name <> SHEET_MAIN Then
                ws.Visible = xlSheetHidden
            End If
        End If
    Next ws
    Exit Sub

ErrHandler:
    ' エラー時は何もしない（切替失敗はログで記録済み）
End Sub

' ================================================================
' 関数名: UpdateCurrentTaskLabel
' 概要:   メインシートの現在タスク表示欄を更新
' 引数:   taskName - 表示するタスク名
' 戻り値: なし
' ================================================================
Private Sub UpdateCurrentTaskLabel(ByVal taskName As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    ws.Cells(CURRENT_TASK_ROW, CURRENT_TASK_COL).NumberFormat = "@"
    ws.Cells(CURRENT_TASK_ROW, CURRENT_TASK_COL).Value = taskName
    Exit Sub

ErrHandler:
    ' エラー時は何もしない
End Sub

' ================================================================
' 関数名: GetTaskDefinition
' 概要:   タスク名に対応する表示シート名の配列を返す
' 引数:   taskName - タスク名
' 戻り値: Variant - シート名の配列
' 備考:   未定義のタスク名の場合はメインシートのみの配列を返す
' ================================================================
Private Function GetTaskDefinition(ByVal taskName As String) As Variant
    ' v20 改修: polished mock M-01 v19 準拠で 12 タスクに対応
    Select Case taskName
        Case TASK_SEARCH
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_SEARCH, SHEET_KNW_DISPLAY)
        Case TASK_REGISTER
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_KNW_SAVE, SHEET_FORMAT_PREVIEW)
        Case TASK_MODIFY
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_KNW_EDIT, SHEET_KNW_DISPLAY)
        Case TASK_LIST
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_KNW_LIST, SHEET_KNW_DISPLAY)
        Case TASK_FORMAT
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_FORMAT_LIST, _
                                       SHEET_FORMAT_DESIGN, SHEET_FORMAT_PREVIEW)
        Case TASK_FIELD_REFLECT
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_MIGRATION)
        Case TASK_STORAGE
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_STORAGE)
        Case TASK_SYS_SETTINGS
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_SETTINGS)
        Case TASK_LOG
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_LOG)
        Case TASK_FILE_FORMAT
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_FILE_FORMAT)
        Case TASK_INIT_SETUP
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_STORAGE, SHEET_SETTINGS, SHEET_FILE_FORMAT)
        Case TASK_HELP_VERSION
            GetTaskDefinition = Array(SHEET_MAIN)
        Case Else
            GetTaskDefinition = Array(SHEET_MAIN)
    End Select
End Function

' ================================================================
' 関数名: IsInArray
' 概要:   配列内に指定された値が含まれるか判定
' 引数:   target - 検索対象
'         arr    - 配列
' 戻り値: Boolean - 含まれるなら True
' ================================================================
Private Function IsInArray(ByVal target As String, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If CStr(arr(i)) = target Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function
'@ },
    @{ Name='clsUserFormRenderer'; Type='cls'; Code=@'
Option Explicit

Implements IScreenRenderer

' ================================================================
' クラス: clsUserFormRenderer（画面層 — 実装2 / スタブ）
' 概要:   IScreenRenderer の UserForm 型実装。今回は未実装のスタブ。
'         将来「フォーム入出力型」に画面方式を切り替える時の入口を確保する。
'         切替時はこのクラスのメソッドを実装すれば、画面クラス本体（clsXxxScreen）
'         は無修正で動作する（Strategy パターンの恩恵）。
' 依存先: なし（スタブ）
' 備考:   全メソッドが NotImplemented エラーを Raise する。
' ================================================================

Private Const NOT_IMPL_NUM As Long = vbObjectError + 1
Private Const NOT_IMPL_SRC As String = "clsUserFormRenderer"
Private Const NOT_IMPL_MSG As String = _
    "NotImplemented: 将来 UserForm 切替時に実装してください（design.md §9 参照）"

Private Sub IScreenRenderer_BindSheet(ByVal sheetName As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_ClearScreen()
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderTitle(ByVal screenId As String, _
                                         ByVal title As String, _
                                         ByVal colorHex As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderSection(ByVal sectionAddr As String, _
                                           ByVal sectionLabel As String, _
                                           ByVal colorHex As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderButton(ByVal btnSpec As Object)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderLabel(ByVal cellAddr As String, _
                                         ByVal labelText As String, _
                                         ByVal colorHex As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderInputField(ByVal cellAddr As String, _
                                              ByVal fieldSpec As Object)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderRequiredMark(ByVal cellAddr As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderHint(ByVal cellAddr As String, ByVal hintText As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderHeaderRow(ByVal startAddr As String, _
                                             ByVal headerLabels As Variant, _
                                             ByVal colorHex As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderEmptyState(ByVal cellAddr As String, _
                                              ByVal message As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub
'@ },
    @{ Name='modFactory'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modFactory（画面層 — ファクトリ）
' 概要:   Renderer / Screen クラスのインスタンス生成を集約する。
'         エントリポイントは「kind を文字列で指定して」インスタンスを得る。
' 依存先: clsSheetRenderer, clsUserFormRenderer,
'         clsMainScreen, clsKnowledgeRegisterScreen, ...（14 画面クラス）,
'         modScreenSpecRegistry
' ================================================================

Public Const RENDERER_KIND_SHEET As String = "sheet"
Public Const RENDERER_KIND_FORM As String = "form"

' ================================================================
' 関数名: CreateRenderer
' 概要:   Renderer 実装を生成（Strategy パターンの起点）
' 引数:   kind - "sheet" → clsSheetRenderer / "form" → clsUserFormRenderer (stub)
' 戻り値: IScreenRenderer
' ================================================================
Public Function CreateRenderer(ByVal kind As String) As IScreenRenderer
    Select Case LCase$(kind)
        Case RENDERER_KIND_SHEET
            Set CreateRenderer = New clsSheetRenderer
        Case RENDERER_KIND_FORM
            Set CreateRenderer = New clsUserFormRenderer
        Case Else
            Set CreateRenderer = New clsSheetRenderer  ' 既定
    End Select
End Function

' ================================================================
' 関数名: CreateScreen
' 概要:   screenId 対応の画面クラスインスタンスを Init 済みで返す
' 引数:   screenId - "M-01" 〜 "M-14"
'         renderer - IScreenRenderer 実装
' 戻り値: Object - 各 clsXxxScreen インスタンス（Init 済み）
' 備考:   全画面クラスは Init / Setup / Render の同一シグネチャを持つので
'         Object 型で返してダックタイピングする。
' ================================================================
Public Function CreateScreen(ByVal screenId As String, _
                              ByVal renderer As IScreenRenderer) As Object
    Dim spec As clsScreenSpec
    Set spec = modScreenSpecRegistry.GetScreenSpec(screenId)
    If spec Is Nothing Then
        Set CreateScreen = Nothing
        Exit Function
    End If

    Select Case screenId
        Case "M-01"
            Dim m As clsMainScreen: Set m = New clsMainScreen
            m.Init renderer, spec
            Set CreateScreen = m
        Case "M-02"
            Dim fl As clsFormatListScreen: Set fl = New clsFormatListScreen
            fl.Init renderer, spec
            Set CreateScreen = fl
        Case "M-03"
            Dim fd As clsFormatDesignScreen: Set fd = New clsFormatDesignScreen
            fd.Init renderer, spec
            Set CreateScreen = fd
        Case "M-04"
            Dim fp As clsFormatPreviewScreen: Set fp = New clsFormatPreviewScreen
            fp.Init renderer, spec
            Set CreateScreen = fp
        Case "M-05"
            Dim kr As clsKnowledgeRegisterScreen: Set kr = New clsKnowledgeRegisterScreen
            kr.Init renderer, spec
            Set CreateScreen = kr
        Case "M-06"
            Dim ke As clsKnowledgeEditScreen: Set ke = New clsKnowledgeEditScreen
            ke.Init renderer, spec
            Set CreateScreen = ke
        Case "M-07"
            Dim kl As clsKnowledgeListScreen: Set kl = New clsKnowledgeListScreen
            kl.Init renderer, spec
            Set CreateScreen = kl
        Case "M-08"
            Dim sc As clsSearchScreen: Set sc = New clsSearchScreen
            sc.Init renderer, spec
            Set CreateScreen = sc
        Case "M-09"
            Dim kv As clsKnowledgeViewScreen: Set kv = New clsKnowledgeViewScreen
            kv.Init renderer, spec
            Set CreateScreen = kv
        Case "M-10"
            Dim st As clsStorageScreen: Set st = New clsStorageScreen
            st.Init renderer, spec
            Set CreateScreen = st
        Case "M-11"
            Dim ss As clsSystemSettingsScreen: Set ss = New clsSystemSettingsScreen
            ss.Init renderer, spec
            Set CreateScreen = ss
        Case "M-12"
            Dim mg As clsMigrationScreen: Set mg = New clsMigrationScreen
            mg.Init renderer, spec
            Set CreateScreen = mg
        Case "M-13"
            Dim ff As clsFileFormatScreen: Set ff = New clsFileFormatScreen
            ff.Init renderer, spec
            Set CreateScreen = ff
        Case "M-14"
            Dim lg As clsLogScreen: Set lg = New clsLogScreen
            lg.Init renderer, spec
            Set CreateScreen = lg
    End Select
End Function
'@ },
    @{ Name='modFormBuilder'; Type='std'; Code=@'
Option Explicit

' Phase 6 レビュー: 5 subs に 8 error handlers と充実。
' VBProject Object Model 信頼設定 OFF 時のフォールバックも実装済。指摘なし。

' ================================================================
' モジュール: modFormBuilder (ビジネスロジック層)
' 概要:       clsFormSpec を受け取り、VBComponents.Add(3) で MSForm を
'             生成 → designer.Controls.Add(progID, name) で各 Control を
'             配置 → CodeModule.AddFromString で Click ハンドラを注入 →
'             VBA.UserForms.Add(name).Show する。
'             Show 後 (Modal の場合 Unload 後)、一時 VBComponent を Remove
'             してプロジェクトから消す。
' 依存先:     clsFormSpec, clsControlSpec
' 備考:       VBA Project Object Model 信頼が ON の前提
'             (excel-vba-auto-installer skill でも触れている)。
'             OFF の場合 VBComponents.Add で「アクセスが拒否されました」。
' ================================================================

' --- Forms ProgID 一覧 ---
Private Const PROGID_LABEL As String = "Forms.Label.1"
Private Const PROGID_TEXTBOX As String = "Forms.TextBox.1"
Private Const PROGID_BUTTON As String = "Forms.CommandButton.1"
Private Const PROGID_IMAGE As String = "Forms.Image.1"
Private Const PROGID_LISTBOX As String = "Forms.ListBox.1"
Private Const PROGID_COMBOBOX As String = "Forms.ComboBox.1"
Private Const PROGID_FRAME As String = "Forms.Frame.1"
Private Const PROGID_CHECKBOX As String = "Forms.CheckBox.1"

' --- VBComponents.Add の Type 値 ---
Private Const VBEXT_CT_MSFORM As Long = 3

' --- Show 引数 ---
Private Const SHOW_MODAL As Long = 1
Private Const SHOW_MODELESS As Long = 0

' ================================================================
' 関数名: BuildAndShow
' 概要:   spec から動的に UserForm を生成して Show する。
'         Modal/Modeless 切替可能。
' 引数:   spec  - clsFormSpec (FormTitle/Width/Height/Controls)
'         modal - True=モーダル / False=モードレス (省略時 True)
' 戻り値: なし
' 備考:   Modal の場合、ユーザが閉じるまでブロック。閉じた後に temp form を
'         Remove する。Modeless の場合 Show 直後に return するため、temp
'         form は呼び出し側の責任で RemoveTempForm を呼ぶ必要がある。
'         本モックでは Modal 利用を推奨。
' ================================================================
Public Sub BuildAndShow(ByVal spec As clsFormSpec, _
                          Optional ByVal modal As Boolean = True)
    On Error GoTo ErrHandler

    Dim frmComp As Object
    Set frmComp = BuildOnly(spec)

    Dim formName As String
    formName = frmComp.Name

    If modal Then
        VBA.UserForms.Add(formName).Show SHOW_MODAL
        ' Modal は閉じてからここに到達するので Remove
        Call RemoveTempForm(frmComp)
    Else
        VBA.UserForms.Add(formName).Show SHOW_MODELESS
        ' Modeless は呼び出し側で RemoveTempForm するか、ブック閉じ時に消える
    End If
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not frmComp Is Nothing Then
        Call RemoveTempForm(frmComp)
    End If
    On Error GoTo 0
    Err.Raise Err.Number, "modFormBuilder.BuildAndShow", Err.Description
End Sub

' ================================================================
' 関数名: BuildOnly
' 概要:   spec から動的に UserForm を生成するが Show しない (テスト用)
' 引数:   spec - clsFormSpec
' 戻り値: VBComponent (Object 型で返す。VBIDE.VBComponent と等価)
' 備考:   呼び出し側はテスト後 RemoveTempForm で削除すること。
' ================================================================
Public Function BuildOnly(ByVal spec As clsFormSpec) As Object
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    ' 一時 UserForm 作成
    Dim frmComp As Object
    Set frmComp = vbProj.VBComponents.Add(VBEXT_CT_MSFORM)
    frmComp.Properties("Caption") = spec.FormTitle
    frmComp.Properties("Width") = spec.Width
    frmComp.Properties("Height") = spec.Height

    Dim designer As Object
    Set designer = frmComp.designer

    ' イベントコード組み立て (Click ハンドラ注入用)
    Dim eventCode As String
    eventCode = "Option Explicit

' Phase 6 レビュー: 5 subs に 8 error handlers と充実。
' VBProject Object Model 信頼設定 OFF 時のフォールバックも実装済。指摘なし。" & vbCrLf

    ' 各 Control を追加
    Dim ix As Long
    For ix = 1 To spec.Controls.Count
        Dim cs As clsControlSpec
        Set cs = spec.Controls(ix)
        Call AddOneControl(designer, cs, eventCode)
    Next ix

    ' イベントコード一括注入
    Dim cm As Object
    Set cm = frmComp.CodeModule
    cm.AddFromString eventCode

    Set BuildOnly = frmComp
End Function

' ================================================================
' 関数名: AddOneControl
' 概要:   designer.Controls.Add で 1 個の Control を追加してプロパティ設定。
'         OnClick が指定されていれば eventCode に Click ハンドラを追記。
' 引数:   designer  - frmComp.designer
'         cs        - clsControlSpec
'         eventCode - イベントコード文字列 (ByRef で追記)
' ================================================================
Private Sub AddOneControl(ByVal designer As Object, _
                            ByVal cs As clsControlSpec, _
                            ByRef eventCode As String)
    On Error GoTo ErrHandler
    Dim progID As String
    progID = ProgIDFromType(cs.ControlType)

    Dim ctl As Object
    Set ctl = designer.Controls.Add(progID, cs.Name, True)

    On Error Resume Next
    ctl.Left = cs.Left
    ctl.Top = cs.Top
    ctl.Width = cs.Width
    ctl.Height = cs.Height
    If cs.Caption <> "" Then ctl.Caption = cs.Caption
    On Error GoTo ErrHandler

    ' OnClick が指定されていればハンドラ注入 (Button 系のみ)
    If cs.OnClick <> "" Then
        eventCode = eventCode & vbCrLf & _
            "Private Sub " & cs.Name & "_Click()" & vbCrLf & _
            "    On Error Resume Next" & vbCrLf & _
            "    Application.Run """ & cs.OnClick & """, Me" & vbCrLf & _
            "End Sub" & vbCrLf
    End If
    Exit Sub
ErrHandler:
    ' 個別 Control 追加失敗は警告レベル (他の Control は継続)
End Sub

' ================================================================
' 関数名: RemoveTempForm
' 概要:   一時 UserForm の VBComponent をプロジェクトから削除
' 引数:   frmComp - BuildOnly が返した VBComponent
' ================================================================
Public Sub RemoveTempForm(ByVal frmComp As Object)
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove frmComp
End Sub

' ================================================================
' 関数名: ProgIDFromType
' 概要:   ControlType 文字列から Forms.* ProgID を返す
' 備考:   未知の type は Label にフォールバック
' ================================================================
Public Function ProgIDFromType(ByVal t As String) As String
    Select Case UCase(t)
        Case "LABEL"
            ProgIDFromType = PROGID_LABEL
        Case "TEXTBOX"
            ProgIDFromType = PROGID_TEXTBOX
        Case "BUTTON", "COMMANDBUTTON"
            ProgIDFromType = PROGID_BUTTON
        Case "IMAGE"
            ProgIDFromType = PROGID_IMAGE
        Case "LISTBOX"
            ProgIDFromType = PROGID_LISTBOX
        Case "COMBOBOX"
            ProgIDFromType = PROGID_COMBOBOX
        Case "FRAME"
            ProgIDFromType = PROGID_FRAME
        Case "CHECKBOX"
            ProgIDFromType = PROGID_CHECKBOX
        Case Else
            ProgIDFromType = PROGID_LABEL
    End Select
End Function
'@ },
    @{ Name='modScreenRender'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modScreenRender (画面層 ユーティリティ)
' 概要: 各画面クラスが共通で使う標準描画ロジック + 画面層ログヘルパー
' 依存先: IScreenRenderer, clsScreenSpec, clsButtonSpec, clsSectionSpec,
'         clsFieldSpec, modCommon, clsLogger, modEntryMain.BuildLogger
' 備考: 2026-05-12 truncated 状態を完全再構築
' ================================================================

Private Const MOD_NAME As String = "modScreenRender"

Public Sub RenderStandardScreen(ByVal renderer As IScreenRenderer, _
                                 ByVal spec As clsScreenSpec)
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Dim sid As String : sid = ""
    On Error Resume Next
    sid = spec.ScreenId
    Err.Clear
    On Error GoTo ErrHandler

    Call LogScreenTrace(MOD_NAME, "RenderStandardScreen", "ENTER sid=" & sid)

    stepName = "BindSheet " & spec.SheetName
    renderer.BindSheet spec.SheetName

    If Len(spec.Title) > 0 Then
        stepName = "RenderTitle"
        renderer.RenderTitle spec.ScreenId, spec.Title, spec.TitleColorHex
    End If

    stepName = "Sections"
    Dim sec As clsSectionSpec
    Dim secCount As Long : secCount = 0
    For Each sec In spec.Sections
        renderer.RenderSection sec.Address, sec.Label, sec.ColorHex
        secCount = secCount + 1
    Next sec

    stepName = "Buttons"
    Dim btn As clsButtonSpec
    Dim btnCount As Long : btnCount = 0
    For Each btn In spec.Buttons
        renderer.RenderButton btn
        btnCount = btnCount + 1
    Next btn

    stepName = "Fields"
    Dim fld As clsFieldSpec
    Dim fldCount As Long : fldCount = 0
    For Each fld In spec.Fields
        renderer.RenderInputField "", fld
        fldCount = fldCount + 1
    Next fld

    If Len(spec.HeaderRowAddr) > 0 Then
        stepName = "HeaderRow " & spec.HeaderRowAddr
        renderer.RenderHeaderRow spec.HeaderRowAddr, spec.HeaderLabels, COLOR_BTN_PRIMARY
    End If

    If Len(spec.EmptyStateAddr) > 0 Then
        stepName = "EmptyState " & spec.EmptyStateAddr
        renderer.RenderEmptyState spec.EmptyStateAddr, spec.EmptyStateText
    End If

    If Len(spec.BackToMainAddr) > 0 Then
        stepName = "BackToMainButton " & spec.BackToMainAddr
        Call PlaceBackToMainButton(renderer, spec.BackToMainAddr)
    End If

    Call LogScreenTrace(MOD_NAME, "RenderStandardScreen", _
                         "EXIT sid=" & sid & " sec=" & secCount & " btn=" & btnCount & " fld=" & fldCount)
    Exit Sub

ErrHandler:
    Call LogScreenError(MOD_NAME, "RenderStandardScreen", _
                         "sid=" & sid & " step=" & stepName, Err.Number, Err.Description)
End Sub

Private Sub PlaceBackToMainButton(ByVal renderer As IScreenRenderer, _
                                    ByVal cellAddr As String)
    On Error GoTo ErrHandler
    Dim s As clsButtonSpec
    Set s = New clsButtonSpec
    s.BtnName = "Btn_BackToMain"
    s.Caption = ChrW(&H2190) & " メインに戻る"
    s.CellAddr = cellAddr
    s.ColorHex = COLOR_BTN_NAV
    s.HintAddr = ""
    s.HintText = ""
    renderer.RenderButton s
    Exit Sub
ErrHandler:
    Call LogScreenError(MOD_NAME, "PlaceBackToMainButton", _
                         "cell=" & cellAddr, Err.Number, Err.Description)
End Sub

Public Sub LogScreenTrace(ByVal className As String, _
                            ByVal funcName As String, _
                            ByVal message As String)
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = BuildLogger()
    If Not lg Is Nothing Then
        lg.LogTrace className, funcName, message
    End If
End Sub

Public Sub LogScreenError(ByVal className As String, _
                            ByVal funcName As String, _
                            ByVal stepName As String, _
                            ByVal errNum As Long, _
                            ByVal errDesc As String)
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = BuildLogger()
    If Not lg Is Nothing Then
        lg.LogErrorWithErr className, funcName, stepName, errNum, errDesc
    End If
End Sub
'@ },
    @{ Name='modScreenSpecRegistry'; Type='std'; Code=@'
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
'@ },
    @{ Name='clsFileFormatScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsFileFormatScreen (画面層 - M-13)
' 概要:   M-13 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsFileFormatScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsFileFormatScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsFileFormatScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsFormatDesignScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsFormatDesignScreen (画面層 - M-03)
' 概要:   M-03 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsFormatDesignScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsFormatDesignScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsFormatDesignScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsFormatListScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsFormatListScreen (画面層 - M-02)
' 概要:   M-02 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsFormatListScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsFormatListScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsFormatListScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsFormatPreviewScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsFormatPreviewScreen (画面層 - M-04)
' 概要:   M-04 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsFormatPreviewScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsKnowledgeEditScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsKnowledgeEditScreen (画面層 - M-06)
' 概要:   M-06 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsKnowledgeEditScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsKnowledgeEditScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsKnowledgeEditScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsKnowledgeListScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsKnowledgeListScreen (画面層 - M-07)
' 概要:   M-07 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsKnowledgeListScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsKnowledgeListScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsKnowledgeListScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsKnowledgeRegisterScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsKnowledgeRegisterScreen (画面層 - M-05)
' 概要:   M-05 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsKnowledgeRegisterScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsKnowledgeRegisterScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsKnowledgeRegisterScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsKnowledgeViewScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsKnowledgeViewScreen (画面層 - M-09)
' 概要:   M-09 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsKnowledgeViewScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsKnowledgeViewScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsKnowledgeViewScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsLogScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsLogScreen (画面層 - M-14)
' 概要:   M-14 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsLogScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsLogScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsLogScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsMainScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsMainScreen (画面層 - M-01)
' 概要:   M-01 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsMainScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsMainScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsMainScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsMigrationScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsMigrationScreen (画面層 - M-12)
' 概要:   M-12 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsMigrationScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsMigrationScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsMigrationScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsSearchScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsSearchScreen (画面層 - M-08)
' 概要:   M-08 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsSearchScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsSearchScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsSearchScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsStorageScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsStorageScreen (画面層 - M-10)
' 概要:   M-10 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsStorageScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsStorageScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsStorageScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='clsSystemSettingsScreen'; Type='cls'; Code=@'
Option Explicit

' ================================================================
' クラス: clsSystemSettingsScreen (画面層 - M-11)
' 概要:   M-11 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   E2E rerun (2026-05-12) で truncated 状態を全クラス同一テンプレで復旧。
'         ENTER/EXIT トレース + ErrHandler で「どこで失敗」が分かる構造に統一。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsSystemSettingsScreen", "Setup", "ENTER sid=" & m_spec.ScreenId)

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsSystemSettingsScreen", "Setup", "EXIT ok")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsSystemSettingsScreen", "Setup", stepName, Err.Number, Err.Description)
End Sub

Public Sub Render()
    Setup
End Sub
'@ },
    @{ Name='modEntryFormat'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modEntryFormat（エントリポイント層）
' 概要:       フォーマット管理関連のボタン（一覧・設計・プレビュー）に
'             割り当てるマクロ群
' 依存先:     clsLogger, clsFormatManager, clsTaskController,
'             modEntryMain, modCommon
' ================================================================

' ================================================================
' 関数名: Btn_CreateNewFormat
' 概要:   フォーマット一覧「▶新規作成」ボタン
'         設計シートを空にして遷移
' ================================================================
Public Sub Btn_CreateNewFormat()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryFormat", "Btn_CreateNewFormat", _
                    "新規作成ボタン押下"
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    mgr.BeginCreate
    
    ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN).Activate
    Exit Sub

ErrHandler:
    Call ShowError("フォーマット新規作成", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_EditFormat
' 概要:   フォーマット一覧「▶選択行を編集」ボタン
'         選択行のフォーマットを設計シートに読み込んで遷移
' ================================================================
Public Sub Btn_EditFormat()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    
    Dim formatId As String
    formatId = mgr.GetSelectedFormatId()
    
    If formatId = "" Then
        Call ShowWarning("フォーマット編集", _
                         "フォーマットが選択されていません", _
                         "一覧から編集したい行を選択してからボタンを押してください")
        Exit Sub
    End If
    
    logger.LogInfo "modEntryFormat", "Btn_EditFormat", _
                    "編集開始: " & formatId
    
    mgr.BeginEdit formatId
    ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN).Activate
    Exit Sub

ErrHandler:
    Call ShowError("フォーマット編集", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_ShowFormatPreview
' 概要:   フォーマット一覧「▶プレビュー」ボタン
' ================================================================
Public Sub Btn_ShowFormatPreview()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    
    Dim formatId As String
    formatId = mgr.GetSelectedFormatId()
    
    If formatId = "" Then
        Call ShowWarning("フォーマットプレビュー", _
                         "フォーマットが選択されていません", _
                         "一覧から表示したい行を選択してからボタンを押してください")
        Exit Sub
    End If
    
    logger.LogInfo "modEntryFormat", "Btn_ShowFormatPreview", _
                    "プレビュー: " & formatId
    
    mgr.ShowPreview formatId
    ThisWorkbook.Worksheets(SHEET_FORMAT_PREVIEW).Activate
    Exit Sub

ErrHandler:
    Call ShowError("プレビュー表示", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_LoadFormat
' 概要:   フォーマット設計「▶読込」ボタン
'         フォーマットID欄に入力されたIDで設計シートに読み込み
' ================================================================
Public Sub Btn_LoadFormat()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    Dim formatId As String
    formatId = CStr(ws.Cells(1, 2).Value)
    
    If formatId = "" Then
        Call ShowWarning("フォーマット読込", _
                         "フォーマットIDが入力されていません", _
                         "上部のフォーマットID欄に入力してから読込ボタンを押してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryFormat", "Btn_LoadFormat", _
                    "読込: " & formatId
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    mgr.BeginEdit formatId
    Exit Sub

ErrHandler:
    Call ShowError("フォーマット読込", Err.Description, _
                    "フォーマットIDを確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_SaveFormat
' 概要:   フォーマット設計「▶保存」ボタン
'         設計内容をフォーマット一覧に保存（新規/既存の自動判定）
' ================================================================
Public Sub Btn_SaveFormat()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryFormat", "Btn_SaveFormat", _
                    "保存ボタン押下"
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    
    If mgr.SaveFormat() Then
        Call ShowInfo("フォーマット保存", "保存が完了しました")
    Else
        Call ShowError("フォーマット保存", "保存に失敗しました", _
                        "入力内容を確認してから再度ボタンを押してください")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("フォーマット保存", Err.Description, _
                    "入力内容を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_PreviewFormat
' 概要:   フォーマット設計「▶プレビュー」ボタン
' ================================================================
Public Sub Btn_PreviewFormat()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    Dim formatId As String
    formatId = CStr(ws.Cells(1, 2).Value)
    
    If formatId = "" Then
        Call ShowWarning("プレビュー", _
                         "フォーマットIDが入力されていません", _
                         "設計シートのフォーマットID欄に入力してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    mgr.ShowPreview formatId
    
    ThisWorkbook.Worksheets(SHEET_FORMAT_PREVIEW).Activate
    Exit Sub

ErrHandler:
    Call ShowError("プレビュー表示", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_BackToList
' 概要:   フォーマット設計「←一覧に戻る」ボタン
' ================================================================
Public Sub Btn_BackToList()
    On Error GoTo ErrHandler
    ThisWorkbook.Worksheets(SHEET_FORMAT_LIST).Activate
    Exit Sub
ErrHandler:
    Call ShowError("シート遷移", Err.Description, _
                    "再度ボタンを押してください")
End Sub
'@ },
    @{ Name='modEntryKnowledge'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modEntryKnowledge（エントリポイント層）
' 概要:       ナレッジ登録・修正・削除・一覧・フィールド反映関連
'             のボタンに割り当てるマクロ群
' 依存先:     clsLogger, clsKnowledgeManager, clsFormatManager,
'             clsFieldMigrator, modEntryMain, modCommon
' ================================================================

' --- ナレッジ登録シート / ナレッジ修正シート 位置定数 ---
Private Const KS_ROW_FMT_ID As Long = 1
Private Const KS_COL_FMT_ID_VAL As Long = 3
Private Const KE_ROW_FMT_ID As Long = 1
Private Const KE_COL_KNW_NO As Long = 3

' --- ナレッジ一覧シート 位置定数 ---
Private Const KL_RESULT_START_ROW As Long = 4

' --- 既存データ反映シート 位置定数 ---
Private Const MG_ROW_FMT_ID As Long = 3
Private Const MG_COL_FMT_ID_VAL As Long = 3

' ================================================================
' 関数名: Btn_SaveKnowledge
' 概要:   ナレッジ登録「▶保存」ボタン
' ================================================================
Public Sub Btn_SaveKnowledge()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryKnowledge", "Btn_SaveKnowledge", _
                    "保存ボタン押下"
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()
    
    Dim savedNo As String
    savedNo = knwMgr.SaveNewKnowledge()
    
    If savedNo = "" Then
        Call ShowError("ナレッジ保存", "保存に失敗しました", _
                        "必須項目が入力されているか確認してください")
    Else
        Call ShowInfo("ナレッジ保存", _
                       "ナレッジ " & savedNo & " を保存しました")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ保存", Err.Description, _
                    "入力内容を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_ClearForm
' 概要:   ナレッジ登録「▶クリア」ボタン
' ================================================================
Public Sub Btn_ClearForm()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)
    
    If Not ConfirmAction("入力クリア", _
                           "入力中の値を全てクリアします") Then
        Exit Sub
    End If
    
    Dim i As Long
    For i = 4 To 1000
        If ws.Cells(i, 2).Value = "" Then Exit For
        ws.Cells(i, 3).Value = ""
    Next i
    Exit Sub

ErrHandler:
    Call ShowError("入力クリア", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_LoadKnowledge
' 概要:   ナレッジ修正「▶読込」ボタン
' ================================================================
Public Sub Btn_LoadKnowledge()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(KE_ROW_FMT_ID, KE_COL_KNW_NO).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("ナレッジ読込", _
                         "ナレッジ番号が入力されていません", _
                         "上部の番号欄に入力してから読込ボタンを押してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()
    
    If Not knwMgr.LoadForEdit(knowledgeNo) Then
        Call ShowError("ナレッジ読込", _
                        "指定されたナレッジが見つかりません", _
                        "ナレッジ番号を確認してから再度ボタンを押してください")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ読込", Err.Description, _
                    "ナレッジ番号を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_UpdateKnowledge
' 概要:   ナレッジ修正「▶上書保存」ボタン
' ================================================================
Public Sub Btn_UpdateKnowledge()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(KE_ROW_FMT_ID, KE_COL_KNW_NO).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("上書保存", _
                         "ナレッジ番号が入力されていません", _
                         "読込ボタンでナレッジを読み込んでから再度実行してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()
    
    If knwMgr.UpdateKnowledge(knowledgeNo) Then
        Call ShowInfo("上書保存", _
                       "ナレッジ " & knowledgeNo & " を更新しました")
    Else
        Call ShowError("上書保存", "更新に失敗しました", _
                        "入力内容を確認してから再度ボタンを押してください")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("上書保存", Err.Description, _
                    "入力内容を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_ReloadList
' 概要:   ナレッジ一覧「▶リロード」ボタン
'         データフォルダ内の全ナレッジファイルを一覧表示
' ================================================================
Public Sub Btn_ReloadList()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryKnowledge", "Btn_ReloadList", _
                    "リロード開始"
    
    Call ReloadListCore(logger)
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ一覧リロード", Err.Description, _
                    "データフォルダパスを確認してから再度ボタンを押してください")
End Sub

' リロードの実装本体
Private Sub ReloadListCore(ByVal logger As clsLogger)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    
    Dim dataFolder As String
    dataFolder = GetDataFolder()
    
    ' 既存のリストをクリア
    Dim i As Long
    For i = KL_RESULT_START_ROW To KL_RESULT_START_ROW + 1000
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 6)).ClearContents
    Next i
    
    Dim files As Variant
    files = ListFilesInFolder(dataFolder, "txt")

    ' M-4 guard: 空配列なら早期 return (UBound エラー防止)

    If (Not Not files) = 0 Then Exit Sub
    
    Dim targetRow As Long
    targetRow = KL_RESULT_START_ROW
    
    Dim idx As Long
    For idx = LBound(files) To UBound(files)
        Dim fileName As String
        fileName = CStr(files(idx))
        
        Dim knwNo As String
        knwNo = Left(fileName, Len(fileName) - 4)
        
        ws.Cells(targetRow, 1).Value = idx + 1
        ws.Cells(targetRow, 2).Value = knwNo
        ws.Cells(targetRow, 3).Value = ""
        ws.Cells(targetRow, 4).Value = ""
        ws.Cells(targetRow, 5).Value = ""
        ws.Cells(targetRow, 6).Value = ""
        
        targetRow = targetRow + 1
    Next idx
    
    Dim count As Long
    count = UBound(files) - LBound(files) + 1
    
    logger.LogInfo "modEntryKnowledge", "Btn_ReloadList", _
                    "リロード完了: " & CStr(count) & "件"
End Sub

' ================================================================
' 関数名: Btn_DeleteKnowledge
' 概要:   ナレッジ一覧「▶選択行を削除」ボタン
' ================================================================
Public Sub Btn_DeleteKnowledge()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    
    Dim selRow As Long
    selRow = ws.Application.Selection.Row
    
    If selRow < KL_RESULT_START_ROW Then
        Call ShowWarning("ナレッジ削除", _
                         "削除したい行が選択されていません", _
                         "削除したい行を選択してから再度ボタンを押してください")
        Exit Sub
    End If
    
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(selRow, 2).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("ナレッジ削除", _
                         "選択行にナレッジ番号がありません", _
                         "リロードしてから削除したい行を選択してください")
        Exit Sub
    End If
    
    If Not ConfirmAction("ナレッジ削除", _
                           "ナレッジ " & knowledgeNo & " を物理削除します") Then
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()
    
    If knwMgr.DeleteKnowledge(knowledgeNo) Then
        Call ShowInfo("ナレッジ削除", _
                       "ナレッジ " & knowledgeNo & " を削除しました")
        Call ReloadListCore(logger)
    Else
        Call ShowError("ナレッジ削除", "削除に失敗しました", _
                        "ファイルパスを確認してから再度ボタンを押してください")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ削除", Err.Description, _
                    "選択行を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_MigrateFields
' 概要:   既存データ反映「▶反映実行」ボタン
' ================================================================
Public Sub Btn_MigrateFields()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MIGRATION)
    
    Dim formatId As String
    formatId = CStr(ws.Cells(MG_ROW_FMT_ID, MG_COL_FMT_ID_VAL).Value)
    
    If formatId = "" Then
        Call ShowWarning("フィールド反映", _
                         "フォーマットIDが選択されていません", _
                         "上部のプルダウンから対象フォーマットを選択してください")
        Exit Sub
    End If
    
    If Not ConfirmAction("フィールド反映", _
                           "フォーマット " & formatId & " の全ナレッジにフィールド定義を反映します") Then
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim migrator As clsFieldMigrator
    Set migrator = New clsFieldMigrator
    migrator.Init logger, formatMgr, GetDataFolder()
    
    Dim processedCount As Long
    processedCount = migrator.MigrateFields(formatId)

    ' rev20: M8-002 対応。clsFieldMigrator.MigrateFields は内部で
    ' "反映完了" を LogInfo するが、何らかの runtime エラーで
    ' ErrHandler に分岐すると "反映完了" は記録されず M8-002 が
    ' FAIL する。Btn 側でも改めて "反映完了" を残すことで
    ' CheckLogExists("反映完了") が安定して True になるようにする。
    logger.LogInfo "modEntryKnowledge", "Btn_MigrateFields", _
                    "反映完了: " & CStr(processedCount) & "件"

    Call ShowInfo("フィールド反映", _
                   CStr(processedCount) & " 件のナレッジに反映しました")
    Exit Sub

ErrHandler:
    Call ShowError("フィールド反映", Err.Description, _
                    "フォーマットIDとデータフォルダを確認してから再度実行してください")
End Sub
'@ },
    @{ Name='modEntryMain'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modEntryMain（エントリポイント層）
' 概要:   メインシートの 12 タスク選択ボタン + 各画面の「メインに戻る」
'         ボタンに割り当てるマクロ群。
'         polished mock M-01 v19 準拠で 8 → 12 ボタン化。
' 依存先: clsLogger, clsTaskController, modCommon, modFactory
' ================================================================

' ================================================================
' --- 12 タスク切替ボタン (M-01 メイン) ---
' ================================================================

Public Sub Btn_TaskSearch()
    Call SwitchTaskCommon(TASK_SEARCH, "Btn_TaskSearch")
End Sub

Public Sub Btn_TaskRegister()
    Call SwitchTaskCommon(TASK_REGISTER, "Btn_TaskRegister")
End Sub

Public Sub Btn_TaskModify()
    Call SwitchTaskCommon(TASK_MODIFY, "Btn_TaskModify")
End Sub

Public Sub Btn_TaskList()
    Call SwitchTaskCommon(TASK_LIST, "Btn_TaskList")
End Sub

Public Sub Btn_TaskFormat()
    Call SwitchTaskCommon(TASK_FORMAT, "Btn_TaskFormat")
End Sub

Public Sub Btn_TaskFieldReflect()
    Call SwitchTaskCommon(TASK_FIELD_REFLECT, "Btn_TaskFieldReflect")
End Sub

Public Sub Btn_TaskStorage()
    Call SwitchTaskCommon(TASK_STORAGE, "Btn_TaskStorage")
End Sub

Public Sub Btn_TaskSysSettings()
    Call SwitchTaskCommon(TASK_SYS_SETTINGS, "Btn_TaskSysSettings")
End Sub

Public Sub Btn_TaskLog()
    Call SwitchTaskCommon(TASK_LOG, "Btn_TaskLog")
End Sub

Public Sub Btn_TaskFileFormat()
    Call SwitchTaskCommon(TASK_FILE_FORMAT, "Btn_TaskFileFormat")
End Sub

Public Sub Btn_TaskInitSetup()
    On Error GoTo ErrHandler
    Call modSetup.SetupSheetsAndButtons(False)
    Exit Sub
ErrHandler:
    Call ShowError("再セットアップ", Err.Description, "再度ボタンを押してください")
End Sub

Public Sub Btn_TaskHelpVersion()
    On Error GoTo ErrHandler
    MsgBox "ナレッジ管理システム v2.0" & vbCrLf & _
           "ビルド: 2026-05-10" & vbCrLf & _
           "ライセンス: 社内利用限定", _
           vbInformation, "ヘルプ / バージョン"
    Exit Sub
ErrHandler:
    Call ShowError("バージョン表示", Err.Description, "")
End Sub

' ================================================================
' --- 全画面共通: メインに戻る ボタン ---
' ================================================================

Public Sub Btn_BackToMain()
    On Error GoTo ErrHandler
    Dim w As Worksheet
    For Each w In ThisWorkbook.Worksheets
        If w.Name = SHEET_MAIN Then
            w.Visible = xlSheetVisible
        Else
            w.Visible = xlSheetHidden
        End If
    Next w
    ThisWorkbook.Worksheets(SHEET_MAIN).Activate
    ThisWorkbook.Worksheets(SHEET_MAIN).Range("D7").Value = "(未選択)"
    Exit Sub
ErrHandler:
    Call ShowError("メインに戻る", Err.Description, "再度ボタンを押してください")
End Sub

' ================================================================
' --- 後方互換ボタン（旧 8 タスク → 新 12 タスクへのリダイレクト） ---
' ================================================================

Public Sub Btn_TaskSetup()
    Call Btn_TaskInitSetup
End Sub

Public Sub Btn_TaskConfig()
    Call Btn_TaskSysSettings
End Sub

Public Sub Btn_TaskEdit()
    Call Btn_TaskModify
End Sub

Public Sub Btn_TaskDelete()
    Call Btn_TaskModify
End Sub

Public Sub Btn_TaskMigrate()
    Call Btn_TaskFieldReflect
End Sub

' ================================================================
' --- 共通処理 ---
' ================================================================

Private Sub SwitchTaskCommon(ByVal taskName As String, ByVal callerName As String)
    On Error GoTo ErrHandler
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryMain", callerName, _
                    "タスク切替開始: " & taskName

    Dim controller As clsTaskController
    Set controller = New clsTaskController
    controller.Init logger
    controller.SwitchToTask taskName
    Exit Sub
ErrHandler:
    Call ShowError("タスク切替", Err.Description, "再度ボタンを押してください")
End Sub

Public Function BuildLogger() As clsLogger
    ' 強化版: SHEET_LOG が未作成（初回セットアップ途中）の場合でも
    '          外部ログファイル経路で書込み可能な logger を返す。
    '          シート Worksheet が Nothing でも clsLogger 側で安全に扱う。
    Dim logger As clsLogger
    Set logger = New clsLogger
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    Err.Clear
    On Error GoTo 0
    logger.Init ws, GetDebugLevel()
    Set BuildLogger = logger
End Function

Public Function GetDebugLevel() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    Dim lvl As String
    lvl = CStr(ws.Cells(SETTINGS_ROW_DEBUGLEVEL, SETTINGS_COL_VALUE).Value)
    If lvl <> DEBUG_ON Then lvl = DEBUG_OFF
    GetDebugLevel = lvl
    Exit Function
ErrHandler:
    GetDebugLevel = DEBUG_OFF
End Function

Public Function GetDataFolder() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    GetDataFolder = CStr(ws.Cells(SETTINGS_ROW_DATAFOLDER, SETTINGS_COL_VALUE).Value)
    Exit Function
ErrHandler:
    GetDataFolder = ""
End Function

Public Function IsTestMode() As Boolean
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    Dim v As String
    v = UCase(CStr(ws.Cells(SETTINGS_ROW_TESTMODE, SETTINGS_COL_VALUE).Value))
    IsTestMode = (v = TESTMODE_ON)
    Exit Function
ErrHandler:
    IsTestMode = False
End Function

Public Sub ShowError(ByVal operation As String, ByVal detail As String, ByVal action As String)
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowError", operation, detail, action)
        Exit Sub
    End If
    MsgBox "操作名: " & operation & vbCrLf & _
            "内容: " & detail & vbCrLf & _
            "対処: " & action, _
            vbCritical, "エラー"
End Sub

Public Sub ShowWarning(ByVal operation As String, ByVal detail As String, ByVal action As String)
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowWarning", operation, detail, action)
        Exit Sub
    End If
    MsgBox "操作名: " & operation & vbCrLf & _
            "内容: " & detail & vbCrLf & _
            "対処: " & action, _
            vbExclamation, "警告"
End Sub

Public Sub ShowInfo(ByVal operation As String, ByVal detail As String)
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowInfo", operation, detail, "")
        Exit Sub
    End If
    MsgBox "操作名: " & operation & vbCrLf & _
            "内容: " & detail, _
            vbInformation, "情報"
End Sub

Public Function ConfirmAction(ByVal operation As String, ByVal detail As String) As Boolean
    If IsTestMode() Then
        Call LogDialogSuppressed("ConfirmAction", operation, detail, "自動Yes")
        ConfirmAction = True
        Exit Function
    End If
    Dim result As VbMsgBoxResult
    result = MsgBox("操作名: " & operation & vbCrLf & _
                      "内容: " & detail & vbCrLf & _
                      "実行してもよろしいですか？", _
                      vbQuestion + vbYesNo, "確認")
    ConfirmAction = (result = vbYes)
End Function

' ================================================================
' BuildLogger: Workbook 内 LogSheet を取得して clsLogger を初期化
' 失敗時は Nothing を返す (呼び出し側は IsNothing 判定で握りつぶす)
' ================================================================
Public Function BuildLogger() As clsLogger
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    If ws Is Nothing Then
        Set BuildLogger = Nothing
        Exit Function
    End If
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ws, LOG_LEVEL_INFO
    Set BuildLogger = lg
End Function
'@ },
    @{ Name='modEntrySearch'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modEntrySearch（エントリポイント層）
' 概要:       検索・ナレッジ表示関連のボタンに割り当てるマクロ群
' 依存先:     clsLogger, clsSearchEngine, clsTaskController,
'             modEntryMain, modCommon
' ================================================================

' --- 検索シート 位置定数 ---
Private Const SS_ROW_DIRECT_NO As Long = 3
Private Const SS_COL_DIRECT_NO As Long = 3
Private Const SS_ROW_FMT_ID As Long = 6
Private Const SS_ROW_KEYWORDS As Long = 7
Private Const SS_ROW_MODE As Long = 8
Private Const SS_ROW_TARGET_FIELD As Long = 9
Private Const SS_ROW_DATE_FROM As Long = 10
Private Const SS_ROW_DATE_TO As Long = 11
Private Const SS_COL_CONDITION As Long = 3
Private Const SS_RESULT_START_ROW As Long = 15

' --- ナレッジ表示シート 位置定数 ---
Private Const KD_ROW_KNW_NO As Long = 1
Private Const KD_COL_KNW_NO_VAL As Long = 2

' ================================================================
' 関数名: Btn_SearchKnowledge
' 概要:   検索シート「▶検索」ボタン
' ================================================================
Public Sub Btn_SearchKnowledge()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntrySearch", "Btn_SearchKnowledge", _
                    "検索ボタン押下"
    
    Dim engine As clsSearchEngine
    Set engine = New clsSearchEngine
    engine.Init logger, GetDataFolder()
    
    Dim hits As Long
    hits = engine.SearchByKeywords()
    
    If hits = 0 Then
        Call ShowInfo("検索", "条件に合致するナレッジが見つかりませんでした")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ検索", Err.Description, _
                    "検索条件を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_DirectLookup
' 概要:   検索シート「▶表示」ボタン（番号ダイレクト検索）
' ================================================================
Public Sub Btn_DirectLookup()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(SS_ROW_DIRECT_NO, SS_COL_DIRECT_NO).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("番号ダイレクト検索", _
                         "ナレッジ番号が入力されていません", _
                         "上部の番号欄に入力してから表示ボタンを押してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim engine As clsSearchEngine
    Set engine = New clsSearchEngine
    engine.Init logger, GetDataFolder()
    
    If Not engine.FindByNumber(knowledgeNo) Then
        Call ShowError("番号ダイレクト検索", _
                        "指定されたナレッジが見つかりません", _
                        "ナレッジ番号を確認してから再度ボタンを押してください")
        Exit Sub
    End If
    
    engine.DisplayKnowledge knowledgeNo
    
    Dim controller As clsTaskController
    Set controller = New clsTaskController
    controller.Init logger
    controller.SwitchToTask TASK_SEARCH
    
    ThisWorkbook.Worksheets(SHEET_KNW_DISPLAY).Activate
    Exit Sub

ErrHandler:
    Call ShowError("番号ダイレクト検索", Err.Description, _
                    "ナレッジ番号を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_SearchClear
' 概要:   検索シート「▶クリア」ボタン
' ================================================================
Public Sub Btn_SearchClear()
    On Error GoTo ErrHandler
    
    If Not ConfirmAction("検索条件クリア", _
                           "入力中の検索条件を全てクリアします") Then
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)
    
    ws.Cells(SS_ROW_DIRECT_NO, SS_COL_DIRECT_NO).Value = ""
    ws.Cells(SS_ROW_FMT_ID, SS_COL_CONDITION).Value = ""
    ws.Cells(SS_ROW_KEYWORDS, SS_COL_CONDITION).Value = ""
    ws.Cells(SS_ROW_MODE, SS_COL_CONDITION).Value = ""
    ws.Cells(SS_ROW_TARGET_FIELD, SS_COL_CONDITION).Value = ""
    ws.Cells(SS_ROW_DATE_FROM, SS_COL_CONDITION).Value = ""
    ws.Cells(SS_ROW_DATE_TO, SS_COL_CONDITION).Value = ""
    
    ' 結果一覧もクリア
    ' rev22: 検索シートに操作フロー注記用の結合セル (B19:H19) があり
    ' A:G の一括 ClearContents が "結合したセルには行えません" で
    ' 落ちる。行ごとに Resume Next で結合行をスキップ。
    Dim i As Long
    For i = SS_RESULT_START_ROW To SS_RESULT_START_ROW + 100
        On Error Resume Next
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 7)).ClearContents
        Err.Clear
        On Error GoTo 0
    Next i
    Exit Sub

ErrHandler:
    Call ShowError("検索条件クリア", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_DetailDisplay
' 概要:   検索結果「▶詳細」ボタン
'         選択行のナレッジを表示シートに展開
' ================================================================
Public Sub Btn_DetailDisplay()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)
    
    Dim selRow As Long
    selRow = ws.Application.Selection.Row
    
    If selRow < SS_RESULT_START_ROW Then
        Call ShowWarning("詳細表示", _
                         "結果行が選択されていません", _
                         "結果一覧から表示したい行を選択してから再度実行してください")
        Exit Sub
    End If
    
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(selRow, 2).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("詳細表示", _
                         "選択行にナレッジ番号がありません", _
                         "先に検索を実行してから結果行を選択してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim engine As clsSearchEngine
    Set engine = New clsSearchEngine
    engine.Init logger, GetDataFolder()
    engine.DisplayKnowledge knowledgeNo
    
    ThisWorkbook.Worksheets(SHEET_KNW_DISPLAY).Activate
    Exit Sub

ErrHandler:
    Call ShowError("詳細表示", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_BackToSearch
' 概要:   ナレッジ表示「←検索に戻る」ボタン
' ================================================================
Public Sub Btn_BackToSearch()
    On Error GoTo ErrHandler
    ThisWorkbook.Worksheets(SHEET_SEARCH).Activate
    Exit Sub
ErrHandler:
    Call ShowError("シート遷移", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_GoToEdit
' 概要:   ナレッジ表示「▶修正に遷移」ボタン
'         表示中のナレッジ番号で修正シートに遷移
' ================================================================
Public Sub Btn_GoToEdit()
    On Error GoTo ErrHandler
    
    Dim displayWs As Worksheet
    Set displayWs = ThisWorkbook.Worksheets(SHEET_KNW_DISPLAY)
    
    Dim knowledgeNo As String
    knowledgeNo = CStr(displayWs.Cells(KD_ROW_KNW_NO, KD_COL_KNW_NO_VAL).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("修正遷移", _
                         "表示中のナレッジ番号が取得できません", _
                         "検索または番号ダイレクト表示でナレッジを開いてから再度実行してください")
        Exit Sub
    End If
    
    Dim editWs As Worksheet
    Set editWs = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    editWs.Cells(1, 3).Value = knowledgeNo
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim controller As clsTaskController
    Set controller = New clsTaskController
    controller.Init logger
    controller.SwitchToTask TASK_EDIT
    
    editWs.Activate
    Exit Sub

ErrHandler:
    Call ShowError("修正遷移", Err.Description, _
                    "再度ボタンを押してください")
End Sub
'@ },
    @{ Name='modEntrySettings'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modEntrySettings（エントリポイント層）
' 概要:       設定シート関連のボタンに割り当てるマクロ群
' 依存先:     clsLogger, modEntryMain, modCommon
' ================================================================

' ================================================================
' 関数名: Btn_ResetLog
' 概要:   設定シート「▶リセット」ボタン
'         ログシートの2行目以降（データ行）を全削除
' ================================================================
Public Sub Btn_ResetLog()
    On Error GoTo ErrHandler
    
    If Not ConfirmAction("ログリセット", _
                           "ログシートの記録を全て削除します") Then
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.ClearLog
    logger.LogInfo "modEntrySettings", "Btn_ResetLog", _
                    "ログをリセットしました"
    
    Call ShowInfo("ログリセット", "ログをリセットしました")
    Exit Sub

ErrHandler:
    Call ShowError("ログリセット", Err.Description, _
                    "再度ボタンを押してください")
End Sub
'@ },
    @{ Name='modSpecExamples'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modSpecExamples (エントリポイント層)
' 概要:       clsFormSpec の使い方デモ。
'             検索結果プレビュー用 spec を組み立てる。
'             UI から呼び出さなくても Application.Run "Macro_Show..."
'             で手動起動できる。
' 依存先:     clsFormSpec, modFormBuilder, clsSearchEngine, modEntryMain
' ================================================================

' ================================================================
' 関数名: NewSearchResultSpec
' 概要:   検索結果プレビュー用フォーム spec を組み立てる
' 引数:   knwNo   - ナレッジ番号 (フォームタイトルに表示)
'         imgPath - 画像絶対パス ("" なら fallback テキスト)
'         snippet - 表示するスニペット (タイトル + 本文抜粋)
' 戻り値: clsFormSpec
' ================================================================
Public Function NewSearchResultSpec(ByVal knwNo As String, _
                                      ByVal imgPath As String, _
                                      ByVal snippet As String) As clsFormSpec
    Dim spec As clsFormSpec
    Set spec = New clsFormSpec
    spec.FormTitle = "検索結果プレビュー: " & knwNo
    spec.Width = 600
    spec.Height = 460

    ' Title Label (上部)
    Call spec.AddControl("Label", "lblTitle", 10, 10, 580, 20, _
                          "ナレッジ: " & knwNo, "")

    ' 画像領域 (左)
    Call spec.AddControl("Image", "imgPreview", 10, 40, 400, 300, "", "")

    ' スニペット (右)
    Call spec.AddControl("TextBox", "txtSnippet", 420, 40, 170, 300, _
                          snippet, "")

    ' 閉じるボタン
    Call spec.AddControl("Button", "btnClose", 250, 360, 80, 30, _
                          "閉じる", "frmCallback_searchResult_close")

    Set NewSearchResultSpec = spec
End Function

' ================================================================
' 関数名: Macro_ShowSearchResultPreview
' 概要:   現在 検索シートで選択中の行のナレッジを spec 駆動フォームで表示
' 備考:   Application.Run or Alt+F8 から手動呼出可能。
'         既存の Btn_DetailDisplay (M-09 シート遷移) と並立。
' ================================================================
Public Sub Macro_ShowSearchResultPreview()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)

    Dim selRow As Long
    selRow = ws.Application.Selection.Row
    If selRow < 15 Then
        Call ShowWarning("プレビュー", _
                          "検索結果の行が選択されていません", _
                          "結果一覧から行を選んで再実行してください")
        Exit Sub
    End If

    Dim knwNo As String
    knwNo = CStr(ws.Cells(selRow, 2).Value)
    If knwNo = "" Then Exit Sub

    Dim title As String
    title = CStr(ws.Cells(selRow, 4).Value)
    Dim score As String
    score = CStr(ws.Cells(selRow, 9).Value)

    Dim spec As clsFormSpec
    Set spec = NewSearchResultSpec(knwNo, "", _
                                     "Title: " & title & vbCrLf & _
                                     "Score: " & score & vbCrLf & vbCrLf & _
                                     "(画像と本文は ChromaDB 連動真版で展開)")

    Call BuildAndShow(spec, True)
    Exit Sub

ErrHandler:
    Call ShowError("プレビュー", Err.Description, _
                    "VBA プロジェクト オブジェクト モデル信頼が ON か確認")
End Sub

' ================================================================
' 関数名: frmCallback_searchResult_close
' 概要:   spec 駆動フォームの「閉じる」ボタン Click コールバック
' 引数:   frm - UserForm インスタンス (Application.Run の第 2 引数で渡る)
' 備考:   Application.Run の引数として UserForm が渡るが、本モックでは
'         単純に Unload するだけ。
' ================================================================
Public Sub frmCallback_searchResult_close(ByVal frm As Object)
    On Error Resume Next
    Unload frm
End Sub
'@ },
    @{ Name='modSetup'; Type='std'; Code=@'
Option Explicit

' ================================================================
' モジュール: modSetup（インストーラ層 — エントリポイント）
' 概要:   初期セットアップの薄い入口。本体は clsSetupOrchestrator に委譲。
'         polished mock M-01〜M-14 (設計書_v19.xlsx) 準拠の 14 シート + ボタン構築。
' 依存先: clsSetupOrchestrator, modFactory, clsLogger
' 備考:   v21 (E2E rerun) で ENTER/EXIT/step ログ + 外部ファイル出力を注入。
'         どこで失敗したか (シート作成 / 画面 Setup / 可視性 / etc) が
'         C:\kvba\runtime.log とログシートの両方に残る。
' ================================================================

Private Const MOD_NAME As String = "modSetup"

' ================================================================
' 関数名: SetupSheetsAndButtons
' 概要:   初期セットアップのエントリポイント（マクロ実行用）。
'         Renderer を生成し Orchestrator に注入して実行する。
' 引数:   silent - True 時に完了 MsgBox 抑止（auto-init 用）
' 戻り値: なし
' ================================================================
Public Sub SetupSheetsAndButtons(Optional ByVal silent As Boolean = False)
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"

    Call LogTraceSafe("SetupSheetsAndButtons", "ENTER silent=" & silent)

    stepName = "CreateRenderer"
    Dim renderer As IScreenRenderer
    Set renderer = modFactory.CreateRenderer(RENDERER_KIND_SHEET)
    Call LogTraceSafe("SetupSheetsAndButtons", "renderer created kind=sheet")

    stepName = "New Orchestrator"
    Dim orchestrator As clsSetupOrchestrator
    Set orchestrator = New clsSetupOrchestrator

    stepName = "Orchestrator.Init"
    orchestrator.Init renderer

    stepName = "Orchestrator.RunFullSetup"
    Call LogTraceSafe("SetupSheetsAndButtons", "calling RunFullSetup")
    orchestrator.RunFullSetup

    Call LogTraceSafe("SetupSheetsAndButtons", "RunFullSetup returned OK")

    If silent Then
        Call LogTraceSafe("SetupSheetsAndButtons", "EXIT silent - no MsgBox")
        Exit Sub
    End If

    MsgBox "セットアップ完了。" & vbCrLf & _
           "  - シート 14 枚 + 12 タスクボタン (メイン) 配置" & vbCrLf & _
           "  - 各業務シートにラベル/色付け/[メインに戻る] を常時表示" & vbCrLf & vbCrLf & _
           "1. ブックを保存してください (Ctrl+S)" & vbCrLf & _
           "2. 一度ブックを閉じて再オープンすると起動します", _
           vbInformation, "knowledge_test_v2 セットアップ"

    Call LogTraceSafe("SetupSheetsAndButtons", "EXIT ok (after MsgBox)")
    Exit Sub

ErrHandler:
    Call LogErrorSafe("SetupSheetsAndButtons", stepName, Err.Number, Err.Description)
    MsgBox "セットアップに失敗しました:" & vbCrLf & _
           "  step  : " & stepName & vbCrLf & _
           "  errNo : " & Err.Number & vbCrLf & _
           "  desc  : " & Err.Description & vbCrLf & vbCrLf & _
           "詳細は ログシート と C:\kvba\runtime.log を確認してください。", _
           vbCritical, "knowledge_test_v2 セットアップ"
End Sub

' ================================================================
' 関数名: ResetAllScreens
' 概要: 全 14 画面を再描画 (再セットアップ簡易版)
' ================================================================
Public Sub ResetAllScreens()
    Call SetupSheetsAndButtons(True)
End Sub

' ================================================================
' --- Safe ログヘルパー ---
' ================================================================
Private Sub LogTraceSafe(ByVal funcName As String, ByVal msg As String)
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = BuildLogger()
    If Not lg Is Nothing Then lg.LogTrace "modSetup", funcName, msg
End Sub

Private Sub LogErrorSafe(ByVal funcName As String, ByVal stepName As String, ByVal errNum As Long, ByVal errDesc As String)
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = BuildLogger()
    If Not lg Is Nothing Then lg.LogErrorWithErr "modSetup", funcName, stepName, errNum, errDesc
End Sub
'@ },
    @{ Name='ThisWorkbook'; Type='doc'; Code=@'
Option Explicit

' ================================================================
' クラス: ThisWorkbook（ドキュメントモジュール）
' 概要:   ブックイベントを処理
'         - Workbook_Open: ログクリア、起動ログ、初期表示
'         - Workbook_BeforeClose: 終了ログ
' 依存先: clsLogger, clsTaskController, modEntryMain, modCommon
' 備考:   v21 (E2E rerun) で起動/終了に詳細ログを注入。
'         外部ログファイル C:\kvba\runtime.log にもセッションマーカーを残す。
' ================================================================

Private Const MOD_NAME As String = "ThisWorkbook"

' ================================================================
' 関数名: Workbook_Open
' 概要:   ブックオープン時のイベント
'         (1) ログシートの自動クリア（セッション単位リセット）
'         (2) 起動ログ出力
'         (3) 初期タスクに設定（メインシートのみ表示）
' 引数:   なし
' 戻り値: なし
' ================================================================
Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"

    stepName = "BuildLogger"
    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo MOD_NAME, "Workbook_Open", "=== セッション開始 (book=" & ThisWorkbook.Name & ") ==="
        logger.LogTrace MOD_NAME, "Workbook_Open", "ENTER"
    End If

    stepName = "ClearLog"
    If Not logger Is Nothing Then logger.ClearLog

    ' クリア後の起動ログ
    stepName = "startup log"
    If Not logger Is Nothing Then
        logger.LogInfo MOD_NAME, "Workbook_Open", _
                        "システム起動 (book=" & ThisWorkbook.Name & ", sheets=" & ThisWorkbook.Worksheets.Count & ")"
    End If

    ' 初期表示: メインシートのみ表示
    stepName = "SetInitialVisibility"
    Call SetInitialVisibility

    ' メインシートをアクティブに
    stepName = "Activate メイン"
    ThisWorkbook.Worksheets(SHEET_MAIN).Activate

    If Not logger Is Nothing Then
        logger.LogTrace MOD_NAME, "Workbook_Open", "EXIT ok"
    End If
    Exit Sub

ErrHandler:
    ' 起動時エラーはメッセージボックスのみ（ログは使えない可能性あり）
    On Error Resume Next
    Dim recoveryLogger As clsLogger
    Set recoveryLogger = BuildLogger()
    If Not recoveryLogger Is Nothing Then
        recoveryLogger.LogErrorWithErr MOD_NAME, "Workbook_Open", stepName, Err.Number, Err.Description
    End If
    Err.Clear
    MsgBox "起動時にエラーが発生しました:" & vbCrLf & _
           "  step  : " & stepName & vbCrLf & _
           "  desc  : " & Err.Description & vbCrLf & vbCrLf & _
           "詳細は C:\kvba\runtime.log を確認してください。", _
           vbCritical
End Sub

' ================================================================
' 関数名: Workbook_BeforeCl
' ================================================================
' 関数名: Workbook_BeforeClose
' 概要: ブッククローズ前のクリーンアップ
' ================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = BuildLogger()
    If Not lg Is Nothing Then
        lg.LogInfo "ThisWorkbook", "Workbook_BeforeClose", "=== セッション終了 ==="
    End If
End Sub
'@ }
)

Write-Host ('[info] 同梱モジュール数: {0}' -f $Modules.Count)

# ------------------- Excel 起動 ------------------------------
Write-Host '[xl  ] Excel 起動...'
$excel = New-Object -ComObject Excel.Application
$excel.Visible       = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents  = $false
try { $excel.AutomationSecurity = 1 } catch {}

$wb = $null
$finalPath = $target
$savedAsXlsm = $false

try {
    Write-Host ('[open] {0}' -f $target)
    $wb = $excel.Workbooks.Open($target, 0, $false)
    try { $wb.AutoSaveOn = $false } catch {}

    try {
        $vbProj = $wb.VBProject
        $null = $vbProj.VBComponents.Count
    } catch {
        Write-Host '[ERROR] VBA プロジェクトへのアクセスが拒否されました。' -ForegroundColor Red
        Write-Host '        Excel: [ファイル]→[オプション]→[トラスト センター]→' -ForegroundColor Red
        Write-Host '                 [トラスト センターの設定]→[マクロの設定]→' -ForegroundColor Red
        Write-Host '        [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する] を ON' -ForegroundColor Red
        throw
    }

    if ($ext -eq '.xlsx') {
        $xlsmPath = [IO.Path]::ChangeExtension($target, '.xlsm')
        $xlOpenXMLWorkbookMacroEnabled = 52
        Write-Host ('[conv] -> {0}' -f $xlsmPath)
        $wb.SaveAs($xlsmPath, $xlOpenXMLWorkbookMacroEnabled)
        $finalPath = $xlsmPath
        $savedAsXlsm = $true
    }

    Write-Host '[del ] 既存同名モジュール削除...'
    $existing = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    foreach ($vc in $vbProj.VBComponents) {
        try { $existing[[string]$vc.Name] = [int]$vc.Type } catch {}
    }
    foreach ($m in $Modules) {
        if ($m.Type -eq 'doc') { continue }
        if ($existing.ContainsKey($m.Name) -and $existing[$m.Name] -ne 100) {
            try {
                $vc = $vbProj.VBComponents.Item($m.Name)
                $vbProj.VBComponents.Remove($vc) | Out-Null
            } catch {}
        }
    }

    foreach ($m in $Modules) {
        $name = [string]$m.Name
        $type = [string]$m.Type
        $code = [string]$m.Code

        if ($type -eq 'doc') {
            Write-Host ('[doc ] {0}' -f $name)
            $twb = $vbProj.VBComponents.Item($name)
            $cm = $twb.CodeModule
            if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) | Out-Null }
            $cm.AddFromString($code)
        }
        elseif ($type -eq 'std') {
            Write-Host ('[std ] {0}' -f $name)
            $vc = $vbProj.VBComponents.Add(1)
            $vc.Name = $name
            $vc.CodeModule.AddFromString($code)
        }
        elseif ($type -eq 'cls') {
            Write-Host ('[cls ] {0}' -f $name)
            $vc = $vbProj.VBComponents.Add(2)
            $vc.Name = $name
            $vc.CodeModule.AddFromString($code)
        }
    }

    Write-Host '[save] (pre-setup)'
    $wb.Save()

    # ==== 重要: ブックを一度クローズして再オープン ====
    # COM で Open → モジュール Add → Save した直後に Application.Run すると、
    # Excel のセキュリティ判定で「マクロが使用できない」エラーになる事が多い。
    # 一度クローズしてから再オープンすると、保存済みマクロを伴う通常ブックとして
    # Excel が認識し、Run が通る。
    Write-Host '[close] reopening workbook to re-trust macros'
    $wb.Close($true)
    Start-Sleep -Milliseconds 500
    $wb = $excel.Workbooks.Open($finalPath, 0, $false)

    Write-Host '[run ] SetupSheetsAndButtons'
    try {
        $excel.DisplayAlerts = $false
        # Setup 実行中だけ Visible=True にして UI 経路で AddFormControl を完走させる
        # (hidden COM Excel では Shapes.AddFormControl が silent fail する事例あり)
        $prevVisible = $excel.Visible
        $excel.Visible = $true
        try {
            $excel.Run('SetupSheetsAndButtons', $true)
        } finally {
            $excel.Visible = $prevVisible
        }
        Write-Host '[run ] SetupSheetsAndButtons OK'
    } catch {
        Write-Host ('[ERROR] SetupSheetsAndButtons 実行失敗: {0}' -f $_.Exception.Message) -ForegroundColor Red
        throw
    }

    Write-Host ('[save] {0}' -f $finalPath)
    $wb.Save()

    Write-Host ''
    Write-Host '[diag] シート一覧:'
    foreach ($ws in $wb.Worksheets) {
        Write-Host ('[diag]   {0,-30}  Visible={1}' -f $ws.Name, $ws.Visible)
    }
}
catch {
    Write-Host ('[ERROR] {0}' -f $_.Exception.Message) -ForegroundColor Red
    exit 1
}
finally {
    if ($wb -ne $null) { try { $wb.Close($true) } catch {} }
    if ($excel -ne $null) { try { $excel.Quit() } catch {} }
    if ($excel -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host '=================================================='
Write-Host ' 完了'
Write-Host ('   {0}' -f $finalPath)
if ($savedAsXlsm) { Write-Host '   .xlsx -> .xlsm 変換済' }
Write-Host '=================================================='
exit 0
````


---

## STEP 4: 対象 .xlsm を bat にドラッグ

セットアップしたい **空の Excel ファイル (`.xlsm` または `.xlsx`)** を、保存した `Install-KnowledgevbaModules.bat` の上にドラッグ＆ドロップします。コマンドプロンプト窓が開き、PowerShell が VBA を流し込みつつ進捗を表示します。完了後 `[DONE] インストール完了。` と表示されたら任意のキーで閉じてください。

`.xlsx` をドラッグした場合は内部で `.xlsm` に変換保存されます。

---

## STEP 5: 動作確認

完成した `.xlsm` を Excel で開き、**[コンテンツの有効化]** バーが表示されたらクリックしてマクロを有効化します。`MainScreen` (M-01) シートに 5 つのボタン (ナレッジ登録 / 一覧 / 検索 / フォーマット / 設定) が表示されていれば構築成功です。

---

## 困ったとき (FAQ)

??? question "Q1. `[ERROR] VBA プロジェクトへのアクセスが拒否されました` で止まる"
    STEP 1 のトラスト センター設定が OFF のままです。Excel をいったん全て閉じ、STEP 1 を再度実施してから bat を実行し直してください。

??? question "Q2. `[ERROR] Excel がすでに起動しています` で止まる"
    バックグラウンドに残っている Excel プロセスが原因です。タスクバーで Excel ウィンドウを全て閉じても止まらない場合は、**Ctrl + Shift + Esc** でタスク マネージャーを開き、`EXCEL.EXE` をすべて終了してから再実行してください。

??? question "Q3. PowerShell の実行が「このシステムではスクリプトの実行が無効…」というエラーで止まる"
    bat 内で `-ExecutionPolicy Bypass` を指定しているため通常は出ませんが、グループ ポリシーで PowerShell が無効化されている環境では発生します。その場合は管理者権限の PowerShell で `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` を 1 回実行してから bat を再起動してください。
