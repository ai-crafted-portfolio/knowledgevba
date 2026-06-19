---
title: インストール手順
description: knowledgevba をはじめて使うときのセットアップ手順
---

# インストール

knowledgevba は次の 2 つの Excel ブックで構成されます。

| ブック | 役割 |
|---|---|
| `検索.xlsm` | ナレッジの新規登録・修正・削除・検索・表示 |
| `管理.xlsm` | フォーマット定義・保存先フォルダ・ログ等の管理 |

各ブックを 1 つずつ「インストール」することで、ボタン・シート・自動処理がブック内に展開されます。本ページの手順を上から順に行ってください。

---

## 必要なもの

- Microsoft Excel 2019 / Microsoft 365 等、VBA が動作するバージョン
- Windows 10 / 11
- Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」設定が ON になっていること（[STEP 1](#step-1) 参照）

---

## 全体の流れ

1. **STEP 1**: Excel の事前設定（最初の 1 回だけ）
2. **STEP 2**: 作業フォルダを 1 か所決める（例: `C:\KnowledgeMgr\`）
3. **STEP 3**: 作業フォルダの直下に 2 つの空ブックを作る（`検索.xlsm` / `管理.xlsm`）
4. **STEP 4**: 作業フォルダの直下にデータ用 4 フォルダを作る
6. **STEP 5**: `installer\` フォルダを作り、2 つの bat と 1 つの ps1 を保存する
7. **STEP 6**: VBA モジュールのフォルダ（`installer\vba_modules\`）を用意する
8. **STEP 7**: 2 つの bat を順番にダブルクリックして実行する
9. **STEP 8**: 動作確認

---

## STEP 1: Excel の事前設定 {#step-1}

VBA を Excel COM 経由でブックへ書き込むため、Excel 側で「VBA プロジェクトへの外部アクセス」を 1 度だけ許可しておきます。

1. Excel を起動し、空のブックを 1 つ開きます
2. **[ファイル]** → **[オプション]** を開きます
3. 左メニューから **[トラスト センター]** を選び、**[トラスト センターの設定]** ボタンを押します
4. 左メニューから **[マクロの設定]** を選びます
5. 右側の **[VBA プロジェクト オブジェクト モデルへのアクセスを信頼する]** にチェックを入れます
6. **[OK]** を押してすべてのダイアログを閉じます
7. Excel をいったん全て閉じます

この設定が OFF のままだと、STEP 7 で `[ERROR] VBA プロジェクトへのアクセスが拒否されました` と出てインストールが止まります。

---

## STEP 2: 作業フォルダを決める

エクスプローラで作業フォルダを 1 つ作ります。本書では例として `C:\KnowledgeMgr\` を使います。別の場所でも構いません（ただしパスにスペースを含まない場所を推奨）。

---

## STEP 3: 2 つの空ブックを作る

作業フォルダ（例: `C:\KnowledgeMgr\`）の直下に、次の 2 つの空ブックを作成します。Excel で新規ブックを開き、ファイル名を変えて `.xlsm` 形式で保存してください。

```
C:\KnowledgeMgr\
  ├─ 検索.xlsm
  └─ 管理.xlsm
```

!!! warning "保存形式に注意"
    「Excel マクロ有効ブック (\*.xlsm)」を選んでください。`.xlsx` だとマクロが保存されません。

---

## STEP 4: データ用 4 フォルダを作る

作業フォルダの直下に、ナレッジ本体・フォーマット定義・画面定義・バックアップを保存するための 4 つのフォルダを作ります。

```
C:\KnowledgeMgr\
  ├─ data\       ← ナレッジ本体（.txt）が入る
  ├─ formats\    ← フォーマット定義（.txt）が入る
  ├─ ui\         ← 画面定義（.txt）が入る
  └─ backup\     ← 修正・削除時のバックアップが入る
```

中身は空のままで構いません。インストール後にナレッジを登録すると `data\` に、フォーマットを追加すると `formats\` に、それぞれ自動でファイルが作られます。

---

## STEP 4.5: 保存先の設定（格納先設定シート）

v2.3 では、保存先フォルダやログの設定を **テキストの設定ファイルで管理する方式は廃止**しました。設定は各ブック内の **「格納先設定」シート** にまとまり、ブックと一緒に保存されます。手作業での設定ファイル作成は不要です。

### 4.5.1 初回起動時の自動設定

インストール後に各ブックを初めて開くと、保存先の親フォルダをたずねるダイアログが表示されます（既定は `C:\KnowledgeMgr`）。そのまま **[OK]** を押すと、次の 6 つの保存先が親フォルダの下に自動で設定され、ブックに保存されます。

| 設定 | 既定の場所 |
|---|---|
| `data_dir`（ナレッジ本体） | `C:\KnowledgeMgr\data\` |
| `format_dir`（フォーマット定義） | `C:\KnowledgeMgr\formats\` |
| `ui_dir`（画面定義） | `C:\KnowledgeMgr\ui\` |
| `log_dir`（操作ログ） | `C:\KnowledgeMgr\log\` |
| `backup_dir`（バックアップ） | `C:\KnowledgeMgr\backup\` |
| `config_dir`（設定の基準フォルダ） | `C:\KnowledgeMgr\` |

あわせて、動作設定（`debugLevel` = `INFO`、`logRotationRows` = `10000`、`uiSchemaFailMode` = `safeDefault`）の初期値も設定されます。

### 4.5.2 別の場所にインストールしたいとき

保存先を `C:\KnowledgeMgr\` 以外にしたい場合は、STEP 5 で保存する install bat の先頭にある `BASE_DIR` の 1 行を、目的のフォルダに書き換えてから bat を実行してください。

```bat
REM === Fix-4 (ADR-0132): install path constants. Edit BASE_DIR only to retarget. ===
set "BASE_DIR=C:\KnowledgeMgr"
```

`BASE_DIR` を変えると、その下の `data` / `formats` / `ui` / `log` / `backup` がまとめて新しい場所を指します。

### 4.5.3 あとから設定を変えるとき

インストール後に保存先や動作設定を変えたいときは、**管理.xlsm の M-10「格納先設定」シート**でセルの値を直し、**[Ctrl]+[S]** でブックを上書き保存します。専用の保存ボタンはありません（セルに直接入力 → 通常の保存、という流れです）。

M-10 シートは次の 9 項目を 1 枚にまとめています。

| 項目 | 内容 |
|---|---|
| `data_dir` / `format_dir` / `ui_dir` / `log_dir` / `backup_dir` / `config_dir` | 6 つの保存先フォルダ |
| `debugLevel` | ログ出力の詳細度（`INFO` / `DEBUG` / `TRACE`） |
| `logRotationRows` | ログのローテーション行数（既定 `10000`） |
| `uiSchemaFailMode` | 画面定義の読込失敗時の挙動（既定 `safeDefault`） |

`debugLevel` の各値（`INFO` / `DEBUG` / `TRACE`）の意味は、M-10 シート上の NOTE 欄に記載しています。各設定値の詳細は[設定値の変更](settings.md)を参照してください。

!!! note "「設定」シート（旧 M-11）は廃止しました"
    以前は `debugLevel` 等の動作設定を専用の「設定」シート（M-11）で管理していましたが、v2.3 で M-11 を廃止し、上記のとおり **M-10「格納先設定」シートに統合**しました。設定値は M-10 のセルに入り、ブックと一緒に保存されます。

!!! tip "画面更新ボタンを押しても設定値は保持されます"
    各シートの **[画面更新]** ボタンは全シートを再描画しますが、M-10 に入力した 9 項目の値はそのまま保持されます（再描画で消えることはありません）。

---

## STEP 5: installer フォルダを作る

作業フォルダの直下に `installer\` フォルダを作り、次の 3 ファイルを保存します。

```
C:\KnowledgeMgr\
  └─ installer\
       ├─ install_search.bat
       ├─ install_admin.bat
       └─ _auto_install.ps1
```

### 5.1 install_search.bat

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\`
- ファイル名: `install_search.bat`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（UTF-8 にすると日本語が文字化けします）

```bat
@echo off
chcp 932 > nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

REM ====================================================================
REM  knowledgevba v2.3  install_search.bat
REM  - 検索.xlsm に common\ + search\ の VBA モジュールを自動 Import し、
REM    Setup_search を実行する。
REM  - v2.3 注: M-07 ナレッジ一覧は廃止 (ADR-0072 §2.1)。
REM    シート構成は M-08 / M-09 / ログ の 3 シート。
REM ====================================================================

set "XLSM_PATH=%~dp0..\検索.xlsm"

if not exist "!XLSM_PATH!" (
    echo [ERROR] 検索.xlsm が見つかりません: !XLSM_PATH!
    echo         install bat と同じ階層の上 ^( C:\KnowledgeMgr\ ^) に
    echo         検索.xlsm を作成してください ^(本書 §16.4^)。
    timeout /t 10 /nobreak > nul
    exit /b 1
)

REM === EXCEL.EXE 自動 kill (起動中なら taskkill /F でクリーンアップ後 2 秒 wait) ===
REM v2.3 install bat:
REM   従来は EXCEL.EXE が起動中なら [ERROR] でユーザに手動停止を要求していたが、
REM   実機検証で「ユーザはタスクマネージャ操作に不慣れで失敗率が高い」ことが判明したため
REM   bat 側で自動 kill 後に再確認するフローへ変更 (2026-05-26)。
set "EXCEL_COUNT=0"
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I "EXCEL.EXE" >NUL
if not errorlevel 1 (
    for /f %%C in ('tasklist /FI "IMAGENAME eq EXCEL.EXE" /NH 2^>NUL ^| find /C /I "EXCEL.EXE"') do set "EXCEL_COUNT=%%C"
    echo [INFO] EXCEL.EXE が起動中です ^(!EXCEL_COUNT! 件^)。自動 kill します ^(taskkill /F /IM EXCEL.EXE^)。
    taskkill /F /IM EXCEL.EXE >NUL 2>&1
    echo [INFO] taskkill 完了。プロセス終了確定のため 2 秒待機します。
    timeout /T 2 /NOBREAK >NUL
    REM === 再確認: kill 後も残っていればエラー終了 ===
    tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I "EXCEL.EXE" >NUL
    if not errorlevel 1 (
        echo [ERROR] EXCEL.EXE が kill 後も残存しています。
        echo         手動で停止してから再実行してください ^(本書 §16.2.2^)。
        timeout /t 10 /nobreak > nul
        exit /b 1
    )
    echo [INFO] EXCEL.EXE 終了確認 OK。処理を続行します。
) else (
    echo [INFO] EXCEL.EXE は起動していません。処理を続行します。
)

if not exist "%~dp0vba_modules\common" (
    echo [ERROR] vba_modules\common フォルダが見つかりません: %~dp0vba_modules\common
    timeout /t 10 /nobreak > nul
    exit /b 1
)
if not exist "%~dp0vba_modules\search" (
    echo [ERROR] vba_modules\search フォルダが見つかりません: %~dp0vba_modules\search
    timeout /t 10 /nobreak > nul
    exit /b 1
)

echo [INFO] install_search.bat: PowerShell 起動 (Role=search)
REM === Fix-4 (ADR-0132): install path constants. Edit BASE_DIR only to retarget. ===
set "BASE_DIR=C:\KnowledgeMgr"
set "DATA_DIR=%BASE_DIR%\data"
set "FORMAT_DIR=%BASE_DIR%\formats"
set "UI_DIR=%BASE_DIR%\ui"
set "BACKUP_DIR=%BASE_DIR%\backup"
set "LOG_DIR=%BASE_DIR%\log"
set "CONFIG_DIR=%BASE_DIR%"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_auto_install.ps1" -XlsmPath "!XLSM_PATH!" -Role search -DataDir "%DATA_DIR%" -FormatDir "%FORMAT_DIR%" -UiDir "%UI_DIR%" -BackupDir "%BACKUP_DIR%" -LogDir "%LOG_DIR%" -ConfigDir "%CONFIG_DIR%"
set "PS_EXIT=!errorlevel!"

if !PS_EXIT! NEQ 0 (
    echo [FAIL] install_search failed. exit_code=!PS_EXIT!
    echo        トラブルシュート: 本書 §16.7
) else (
    echo [DONE] Install + setup succeeded.
    echo        Excel で 検索.xlsm を開くと、起動シート M-08 が表示されることを
    echo        確認してください ^(本書 §16.7 動作確認^)。
)
timeout /t 5 /nobreak > nul
exit /b !PS_EXIT!
```

### 5.2 install_admin.bat

同様に、次のコードを `install_admin.bat` として保存します（場所・文字コードは 5.1 と同じ）。

```bat
@echo off
chcp 932 > nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

REM ====================================================================
REM  knowledgevba v2.3  install_admin.bat
REM  - 管理.xlsm に common\ + admin\ の VBA モジュールを自動 Import し、
REM    Setup_admin を実行する。
REM  - v2.3 注: M-13 ファイル形式設定は廃止 (ADR-0072 §2.1)。
REM    2026-06-07: M-04 / M-14 retire 済み。2026-06-19: M-11 設定シート retire。
REM    シート構成は M-02 / M-03 / M-12 / M-10 / LOG の 5 シート
REM    (SSOT: clsSetupOrchestrator.cls SHEETS_KANRI)。
REM ====================================================================

set "XLSM_PATH=%~dp0..\管理.xlsm"

if not exist "!XLSM_PATH!" (
    echo [ERROR] 管理.xlsm が見つかりません: !XLSM_PATH!
    echo         install bat と同じ階層の上 ^( C:\KnowledgeMgr\ ^) に
    echo         管理.xlsm を作成してください ^(本書 §16.4^)。
    timeout /t 10 /nobreak > nul
    exit /b 1
)

REM === EXCEL.EXE 自動 kill (起動中なら taskkill /F でクリーンアップ後 2 秒 wait) ===
REM v2.3 install bat:
REM   従来は EXCEL.EXE が起動中なら [ERROR] でユーザに手動停止を要求していたが、
REM   実機検証で「ユーザはタスクマネージャ操作に不慣れで失敗率が高い」ことが判明したため
REM   bat 側で自動 kill 後に再確認するフローへ変更 (2026-05-26)。
set "EXCEL_COUNT=0"
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I "EXCEL.EXE" >NUL
if not errorlevel 1 (
    for /f %%C in ('tasklist /FI "IMAGENAME eq EXCEL.EXE" /NH 2^>NUL ^| find /C /I "EXCEL.EXE"') do set "EXCEL_COUNT=%%C"
    echo [INFO] EXCEL.EXE が起動中です ^(!EXCEL_COUNT! 件^)。自動 kill します ^(taskkill /F /IM EXCEL.EXE^)。
    taskkill /F /IM EXCEL.EXE >NUL 2>&1
    echo [INFO] taskkill 完了。プロセス終了確定のため 2 秒待機します。
    timeout /T 2 /NOBREAK >NUL
    REM === 再確認: kill 後も残っていればエラー終了 ===
    tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I "EXCEL.EXE" >NUL
    if not errorlevel 1 (
        echo [ERROR] EXCEL.EXE が kill 後も残存しています。
        echo         手動で停止してから再実行してください ^(本書 §16.2.2^)。
        timeout /t 10 /nobreak > nul
        exit /b 1
    )
    echo [INFO] EXCEL.EXE 終了確認 OK。処理を続行します。
) else (
    echo [INFO] EXCEL.EXE は起動していません。処理を続行します。
)

if not exist "%~dp0vba_modules\common" (
    echo [ERROR] vba_modules\common フォルダが見つかりません: %~dp0vba_modules\common
    timeout /t 10 /nobreak > nul
    exit /b 1
)
if not exist "%~dp0vba_modules\admin" (
    echo [ERROR] vba_modules\admin フォルダが見つかりません: %~dp0vba_modules\admin
    timeout /t 10 /nobreak > nul
    exit /b 1
)

echo [INFO] install_admin.bat: PowerShell 起動 (Role=admin)
REM === Fix-4 (ADR-0132): install path constants. Edit BASE_DIR only to retarget. ===
set "BASE_DIR=C:\KnowledgeMgr"
set "DATA_DIR=%BASE_DIR%\data"
set "FORMAT_DIR=%BASE_DIR%\formats"
set "UI_DIR=%BASE_DIR%\ui"
set "BACKUP_DIR=%BASE_DIR%\backup"
set "LOG_DIR=%BASE_DIR%\log"
set "CONFIG_DIR=%BASE_DIR%"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_auto_install.ps1" -XlsmPath "!XLSM_PATH!" -Role admin -DataDir "%DATA_DIR%" -FormatDir "%FORMAT_DIR%" -UiDir "%UI_DIR%" -BackupDir "%BACKUP_DIR%" -LogDir "%LOG_DIR%" -ConfigDir "%CONFIG_DIR%"
set "PS_EXIT=!errorlevel!"

if !PS_EXIT! NEQ 0 (
    echo [FAIL] install_admin failed. exit_code=!PS_EXIT!
    echo        トラブルシュート: 本書 §16.7
) else (
    echo [DONE] Install + setup succeeded.
    echo        Excel で 管理.xlsm を開くと、起動シート M-02 が表示されることを
    echo        確認してください ^(本書 §16.7 動作確認^)。
)
timeout /t 5 /nobreak > nul
exit /b !PS_EXIT!
```

### 5.3 _auto_install.ps1

2 つの bat から共通で呼ばれる PowerShell 本体です。下のコードを `_auto_install.ps1` として保存します。

- 場所: `C:\KnowledgeMgr\installer\`
- ファイル名: `_auto_install.ps1`
- ファイルの種類: **すべてのファイル**
- 文字コード: **UTF-8 (BOM 付き)**

```powershell
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$XlsmPath,

    [Parameter(Mandatory = $true)]
    [ValidateSet('search', 'admin')]
    [string]$Role,

    # Fix-4 (ADR-0132): install-injected path scheme. install bat passes these
    # from BASE_DIR; defaults preserve the historical C:\KnowledgeMgr\ layout.
    [string]$DataDir   = 'C:\KnowledgeMgr\data',
    [string]$FormatDir = 'C:\KnowledgeMgr\formats',
    [string]$UiDir     = 'C:\KnowledgeMgr\ui',
    [string]$BackupDir = 'C:\KnowledgeMgr\backup',
    [string]$LogDir    = 'C:\KnowledgeMgr\log',
    [string]$ConfigDir = 'C:\KnowledgeMgr'
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# Fix-4 (ADR-0132): normalize each dir to a single trailing backslash.
function Norm-Dir([string]$d){ if([string]::IsNullOrWhiteSpace($d)){ return $d }; return ($d.TrimEnd('\') + '\') }
$DataDir   = Norm-Dir $DataDir
$FormatDir = Norm-Dir $FormatDir
$UiDir     = Norm-Dir $UiDir
$BackupDir = Norm-Dir $BackupDir
$LogDir    = Norm-Dir $LogDir
$ConfigDir = Norm-Dir $ConfigDir

# ====================================================================
#  knowledgevba v2.3  _auto_install.ps1
#  - install_search.bat / install_admin.bat
#    から共通で呼ばれる本体スクリプト。
#  - 保存時の文字コード: UTF-8 (BOM 付き) / 改行: CRLF
#  - 配置: C:\KnowledgeMgr\installer\_auto_install.ps1
# ====================================================================

function Write-Log($msg) {
    Write-Host ('[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $msg)
}

# Fix-6 (ADR-0133): install-injected path-constant module generation removed (worksheet SSOT).

# ---- Step 1. パス検証 ----
if (-not (Test-Path -LiteralPath $XlsmPath)) {
    Write-Host "[ERROR] ターゲット .xlsm が見つかりません: $XlsmPath" -ForegroundColor Red
    exit 1
}
$target = (Resolve-Path -LiteralPath $XlsmPath).Path
$ext = [IO.Path]::GetExtension($target).ToLowerInvariant()
if ($ext -ne '.xlsm') {
    Write-Host "[ERROR] .xlsm ではありません (拡張子 $ext)。本書 §16.4 で .xlsm を作成してください。" -ForegroundColor Red
    exit 1
}

# ---- Step 1.5 (v2.3 2026-05-27): ui_seed を ui_dir パラメータの直下に自動配布 ----
# 配布物は dist_v2\ui_seed\{管理|検索}\M-NN.txt を含む。
# ui_dir パラメータ (例: C:\KnowledgeMgr\ui\) 直下へ <xlsmFolder>\M-NN.txt を
# コピーする。既存ファイルは上書き (旧 v2.2 stanza の置換が目的)。
$roleToJp = @{
    'admin'    = ([char]0x7BA1)+([char]0x7406)                     # 管理
    'search'   = ([char]0x691C)+([char]0x7D22)                     # 検索
}
$xlsmJp = $roleToJp[$Role]
$uiSeedSrc = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) ('..\ui_seed\' + $xlsmJp)
$uiSeedSrc = [IO.Path]::GetFullPath($uiSeedSrc)
# Fix-6 (ADR-0133): text-file config loader retired (worksheet SSOT); $seedUiDir falls through to $UiDir below.
$seedUiDir = $null
# v2.3.1 (2026-06-11): fresh 環境 fallback - 設定シート未生成時は default ui_dir に配布 (modConfigLoader DEFAULT_UI_DIR と同値)
if(-not $seedUiDir){ $seedUiDir = $UiDir }
if($seedUiDir -and (Test-Path -LiteralPath $uiSeedSrc)){
    $uiDst = Join-Path $seedUiDir $xlsmJp
    if(-not (Test-Path -LiteralPath $uiDst)){
        New-Item -ItemType Directory -Path $uiDst -Force | Out-Null
    }
    Get-ChildItem -LiteralPath $uiSeedSrc -File | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $uiDst -Force
    }
    Write-Log ("[OK] ui_seed 配布: {0} -> {1}" -f $uiSeedSrc, $uiDst)
} else {
    Write-Log ("[INFO] ui_seed 配布スキップ (uiDir={0} src={1})" -f $seedUiDir, $uiSeedSrc)
}

# ---- Step 2. モジュール配置 (common + <role>) を列挙 ----
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$commonDir = Join-Path $scriptDir 'vba_modules\common'
$roleDir   = Join-Path $scriptDir ('vba_modules\' + $Role)

foreach ($d in @($commonDir, $roleDir)) {
    if (-not (Test-Path -LiteralPath $d)) {
        Write-Host "[ERROR] フォルダが見つかりません: $d" -ForegroundColor Red
        exit 1
    }
}

$allFiles = @()
# Filter: only allow well-formed module names (modX/clsX/ClsX/IX + .bas/.cls)
# Reject: literal-path-as-filename buggy files, .bak.* / .bak_* artifacts
$validFilter = {
    $_.Extension -in @('.bas', '.cls') -and
    $_.Length -gt 100 -and
    $_.Name -match '^[A-Za-z][A-Za-z0-9_]*\.(bas|cls)$' -and
    $_.Name -notlike '*.bak.*' -and
    $_.Name -notlike '*.bak_*'
}
$allFiles += Get-ChildItem -LiteralPath $commonDir -Recurse -File | Where-Object $validFilter
$allFiles += Get-ChildItem -LiteralPath $roleDir   -Recurse -File | Where-Object $validFilter

# ThisWorkbook.cls は Document module なので Import せず CodeModule.AddFromString で本体だけ流し込む
$thisWorkbookFile = $allFiles | Where-Object { $_.Name -eq 'ThisWorkbook.cls' } | Select-Object -First 1
$importFiles      = $allFiles | Where-Object { $_.Name -ne 'ThisWorkbook.cls' }

Write-Log ('[OK] モジュール列挙: 全 {0} 本 (内 ThisWorkbook 1 / Import 対象 {1})' -f $allFiles.Count, $importFiles.Count)

# ---- Step 3. staging: source は CP932+CRLF (Phase E 規約)。
# v2.3 fix (2026-05-27): source が UTF-8 (BOM 有無問わず) の場合も対応。
# UTF-8 として valid なら UTF-8 として decode、そうでなければ CP932 として decode。
# 書き戻しは常に CP932 + CRLF (VBE Import が CP932 を期待するため)。
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
$cp932     = [System.Text.Encoding]::GetEncoding(932)
$cp932Repl = [System.Text.Encoding]::GetEncoding(932,
    [System.Text.EncoderFallback]::ReplacementFallback,
    [System.Text.DecoderFallback]::ReplacementFallback)
function Decode-Source([byte[]]$bytes){
    # Strip BOM if present
    if($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF){
        $bytes = $bytes[3..($bytes.Length-1)]
    }
    # Try strict UTF-8 first
    try {
        return $utf8Strict.GetString($bytes)
    } catch {
        # Fall back to CP932
        return $cp932.GetString($bytes)
    }
}
$staging   = Join-Path $env:TEMP ('vba_staging_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Path $staging -Force | Out-Null
$stagedPaths = @()
foreach ($f in $importFiles) {
    $destFp = Join-Path $staging $f.Name
    $bytes  = [IO.File]::ReadAllBytes($f.FullName)
    $text   = Decode-Source $bytes
    $text   = $text -replace "`r`n", "`n"
    $text   = $text -replace "`n", "`r`n"   # CRLF 強制
    # Encode CP932 with replacement fallback (any non-CP932 char becomes '?' but no crash)
    $outBytes = $cp932Repl.GetBytes($text)
    [IO.File]::WriteAllBytes($destFp, $outBytes)
    $stagedPaths += $destFp
}
Write-Log ('[OK] staging 完了: {0}' -f $staging)

# Fix-6 (ADR-0133): Step 3.5 (install path-constant module staging + import) removed.

# ---- Step 4. ThisWorkbook.cls 本体抽出 ----
$thisWorkbookBody = ''
if ($thisWorkbookFile -ne $null) {
    $twbBytes = [IO.File]::ReadAllBytes($thisWorkbookFile.FullName)
    $twbText  = Decode-Source $twbBytes
    $twbText  = $twbText -replace "`r`n", "`n"
    $twbText  = $twbText -replace "`n", "`r`n"
    $rawLines = $twbText -split "`r`n"
    $startIdx = 0
    $seenAttr = $false
    for ($i = 0; $i -lt $rawLines.Count; $i++) {
        $trim = $rawLines[$i].TrimStart()
        if ($trim.StartsWith('VERSION ') -or $trim -eq 'BEGIN' -or $trim -eq 'END' -or $trim.StartsWith('MultiUse ')) { continue }
        if ($trim.StartsWith('Attribute ')) { $seenAttr = $true; continue }
        if ($seenAttr -or $trim -ne '') { $startIdx = $i; break }
    }
    $thisWorkbookBody = ($rawLines[$startIdx..($rawLines.Count - 1)] -join "`r`n")
    Write-Log '[OK] ThisWorkbook 本体抽出 完了'
} else {
    Write-Host "[WARN] ThisWorkbook.cls が見つかりません (Role=$Role)。" -ForegroundColor Yellow
}

# ---- Step 5. Excel COM 起動 ----
Write-Log '[..] Excel COM 起動'
$excel = New-Object -ComObject Excel.Application
$excel.Visible        = $false
$excel.DisplayAlerts  = $false
$excel.EnableEvents   = $false   # Workbook_Open 暴発抑止
$excel.AskToUpdateLinks = $false
try { $excel.AutomationSecurity = 1 } catch {}   # msoAutomationSecurityLow

$wb = $null
try {
    $wb = $excel.Workbooks.Open($target, 0, $false)
    $vbProj = $wb.VBProject   # ← VBA project model 信頼が ON でないと「アクセスが拒否されました」

    # ---- Step 6. 既存同名モジュール削除 + 余分 module strict purge (Document module = Type 100 はスキップ) ----
    # iter23 (2026-06-01): canonical role と無関係な過去残留 module を全部 Remove する
    # strict mode 化。これがないと過去 install で混入した modEntrySettings 等が居座って
    # Application.Run 解決衝突を起こし、E2E が「マクロを実行できません」で失敗する。
    $importNameSet = @{}
    foreach ($f in $importFiles) {
        $importNameSet[[IO.Path]::GetFileNameWithoutExtension($f.Name)] = $true
    }
    $allComponentNames = @()
    foreach ($vc in $vbProj.VBComponents) { $allComponentNames += [string]$vc.Name }
    foreach ($modName in $allComponentNames) {
        try {
            $vc = $vbProj.VBComponents.Item($modName)
        } catch { continue }
        if ($vc.Type -eq 100) { continue }   # Document module (ThisWorkbook/Sheet*) は keep
        # canonical importFiles に含まれる name は再 Import で置換するため Remove
        # canonical importFiles に含まれない name は余分残留 → 同じく Remove (strict purge)
        try {
            $vbProj.VBComponents.Remove($vc) | Out-Null
        } catch {
            Write-Log ('[WARN] Remove FAIL: ' + $modName + ' (' + $_.Exception.Message + ')')
        }
    }

    # ---- Step 6.5. deferred Remove 確定 (feedback_cls_deferred_remove) ----
    Write-Log '[..] Step 6.5 wb.Save() after Remove (deferred remove flush)'
    $wb.Save()
    Start-Sleep -Seconds 1
    Write-Log '[OK] Step 6.5 Save + Sleep complete'

    # ---- Step 7. .bas / .cls を全件 Import ----
    foreach ($p in $stagedPaths) {
        $vbProj.VBComponents.Import($p) | Out-Null
    }
    Write-Log ('[OK] Import 完了 ({0} module)' -f $stagedPaths.Count)

    # ---- Step 8. ThisWorkbook 本体差し替え ----
    if ($thisWorkbookBody -ne '') {
        $twb = $vbProj.VBComponents.Item('ThisWorkbook')
        $cm  = $twb.CodeModule
        if ($cm.CountOfLines -gt 0) {
            $cm.DeleteLines(1, $cm.CountOfLines) | Out-Null
        }
        $cm.AddFromString($thisWorkbookBody)
        Write-Log '[OK] ThisWorkbook 本体注入 完了'
    }

    # ---- Step 9. 一旦保存 (コンパイル確定) ----
    $wb.Save()

    # ---- Step 10. セットアップマクロを Application.Run で起動 ----
    $setupName = 'Setup_' + $Role   # Setup_search / Setup_admin
    Write-Log ('[..] Application.Run "{0}"' -f $setupName)
    try {
        $excel.Run($setupName) | Out-Null
        Write-Log ('[OK] {0} 完了' -f $setupName)
    } catch {
        Write-Host ('[ERROR] {0} 実行に失敗: {1}' -f $setupName, $_.Exception.Message) -ForegroundColor Red
        throw
    }

    # ---- Step 11. セットアップ後の状態を保存 ----
    $wb.Save()
    Write-Log '[OK] xlsm 保存 完了'
}
catch {
    Write-Host ('[ERROR] {0}' -f $_.Exception.Message) -ForegroundColor Red
    exit 1
}
finally {
    if ($wb    -ne $null) { try { $wb.Close($true) } catch {} }
    if ($excel -ne $null) { try { $excel.Quit() }    catch {} }
    if ($excel -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    try { Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue } catch {}
}

Write-Log '[DONE] install + setup 完了'
exit 0
```

---

## STEP 6: VBA モジュールフォルダを用意する

`installer\vba_modules\` フォルダの直下に `common\` / `search\` / `admin\` の 3 つのサブフォルダを作り、各サブフォルダに該当する `.bas` / `.cls` ファイル一式を配置します。

```
C:\KnowledgeMgr\
  └─ installer\
       └─ vba_modules\
            ├─ common\   ← 検索.xlsm / 管理.xlsm 共通モジュール（50 ファイル）
            ├─ search\   ← 検索.xlsm 専用モジュール（8 ファイル）
            └─ admin\    ← 管理.xlsm 専用モジュール（9 ファイル）
```

各フォルダの中身（モジュールごとのソースコード）は **[モジュール一覧](modules/index.md)** ページに 1 ファイル 1 ページで全文掲載しています。1 モジュールずつメモ帳に貼り付け、下記の表のとおりのファイル名（例: `modCommon.bas`、`clsKnowledgeManager.cls` 等）で保存してください。

| 配置先サブフォルダ | ファイル数 | ファイル名一覧 |
|---|---|---|
| `vba_modules\common\` | 50 | 下記 6.1 |
| `vba_modules\search\` | 8 | 下記 6.2 |
| `vba_modules\admin\` | 8 | 下記 6.3 |

現在のバージョンでは **合計 66 ファイル** です。すべて保存し終わったら STEP 7 に進んでください。

### 6.1 common\ フォルダ（50 ファイル）

「検索.xlsm」と「管理.xlsm」が共通で使うモジュールです。基盤となる `clsSetupOrchestrator.cls` などを含む、下記 50 ファイルすべてを `vba_modules\common\` に保存してください。

| ファイル名 | 種類 |
|---|---|
| `clsButtonSpec.cls` | クラス |
| `clsCellAddrHelper.cls` | クラス |
| `clsCellBinding.cls` | クラス |
| `clsCellIO.cls` | クラス |
| `clsControlSpec.cls` | クラス |
| `clsFieldMigrator.cls` | クラス |
| `clsFieldSpec.cls` | クラス |
| `clsFormSpec.cls` | クラス |
| `clsFormatManager.cls` | クラス |
| `clsGridIO.cls` | クラス |
| `clsKnowledgeManager.cls` | クラス |
| `clsLogEntry.cls` | クラス |
| `clsLogger.cls` | クラス |
| `clsScreenSpec.cls` | クラス |
| `clsSearchEngine.cls` | クラス |
| `clsSectionSpec.cls` | クラス |
| `clsSetupOrchestrator.cls` | クラス |
| `clsSheetRenderer.cls` | クラス |
| `ClsStanzaSection.cls` | クラス |
| `ClsStanzaValidationIssue.cls` | クラス |
| `ClsStanzaValidationResult.cls` | クラス |
| `clsStorageResolver.cls` | クラス |
| `clsUserFormRenderer.cls` | クラス |
| `IScreenRenderer.cls` | クラス（インターフェース） |
| `modBtnGuard.bas` | 標準 |
| `modBtnMessages.bas` | 標準 |
| `modButtonWiring.bas` | 標準 |
| `modCommon.bas` | 標準 |
| `modConfigHolder.bas` | 標準 |
| `modConfigLoader.bas` | 標準 |
| `modDateUtil.bas` | 標準 |
| `modEntryUserForm.bas` | 標準 |
| `modFactory.bas` | 標準 |
| `modFileIO.bas` | 標準 |
| `modFormBuilder.bas` | 標準 |
| `modFormControlWiring.bas` | 標準 |
| `modFormatColumns.bas` | 標準 |
| `modFormatLoader.bas` | 標準 |
| `modKnowledgeFileIO.bas` | 標準 |
| `modLogIds.bas` | 標準 |
| `modPreviewForm.bas` | 標準 |
| `modRefresh.bas` | 標準 |
| `modScreenRender.bas` | 標準 |
| `modSetup.bas` | 標準 |
| `modSheetButtons.bas` | 標準 |
| `modSheetMap.bas` | 標準 |
| `modStanzaIO.bas` | 標準 |
| `modStringUtil.bas` | 標準 |
| `modUILoader.bas` | 標準 |
| `modUserFormCallback.bas` | 標準 |

### 6.2 search\ フォルダ（8 ファイル）

「検索.xlsm」専用のモジュールです。ナレッジの検索・表示に加え、新規登録・修正・削除の画面もこのフォルダに含まれます。下記 8 ファイルを `vba_modules\search\` に保存してください。

| ファイル名 | 種類 |
|---|---|
| `ThisWorkbook.cls` | クラス |
| `clsKnowledgeEditScreen.cls` | クラス |
| `clsKnowledgeRegisterScreen.cls` | クラス |
| `clsKnowledgeViewScreen.cls` | クラス |
| `clsSearchScreen.cls` | クラス |
| `modEntryKnowledge.bas` | 標準 |
| `modEntrySearch.bas` | 標準 |
| `modKnowledgeEntryHelpers.bas` | 標準 |

### 6.3 admin\ フォルダ（8 ファイル）

「管理.xlsm」専用のモジュールです。下記 8 ファイルを `vba_modules\admin\` に保存してください。

| ファイル名 | 種類 |
|---|---|
| `ThisWorkbook.cls` | クラス |
| `clsFormatDesignScreen.cls` | クラス |
| `clsFormatListScreen.cls` | クラス |
| `clsFormatPreviewScreen.cls` | クラス |
| `clsMigrationScreen.cls` | クラス |
| `clsStorageScreen.cls` | クラス |
| `modEntryFormat.bas` | 標準 |
| `modEntrySettings.bas` | 標準 |

!!! warning "モジュールファイルの保存形式"
    `.bas` / `.cls` の文字コードは **ANSI（Shift-JIS）** で保存してください。本インストーラは UTF-8 (BOM 有無いずれも) でも自動変換しますが、CP932（ANSI）が確実です。改行コードは **CRLF** にしてください。

---

## STEP 7: bat を順番に実行する

`C:\KnowledgeMgr\installer\` を開き、次の順番で bat を **ダブルクリック** します。

1. `install_search.bat`
2. `install_admin.bat`

各 bat は黒いコンソールウィンドウを開き、状況を表示しながら処理を進めます。最後に `[DONE] Install + setup succeeded.` と表示され、**「続行するには何かキーを押してください . . .」** で停止したら成功です。任意のキーを押してウィンドウを閉じてください。

!!! tip "実行中に Excel が一瞬開いて閉じることがあります"
    bat は内部で Excel を非表示モードで開き、VBA モジュールを取り込んでから保存・終了します。利用者の操作は不要です。

!!! warning "途中で Excel を起動しないでください"
    bat 実行中に手動で Excel を起動すると、インストーラが起動済み Excel を強制終了 (`taskkill /F`) して進行します。作業中のブックがあれば失われる可能性があります。

bat が `[ERROR]` 等で停止した場合は、[困ったとき](troubleshooting.md) を参照してください。

---

## STEP 8: 動作確認

2 つの bat がすべて成功したら、Excel で各ブックを開いて動作を確認します。

### 8.1 検索.xlsm

1. `C:\KnowledgeMgr\検索.xlsm` をダブルクリックで開く
2. 上部に黄色い「セキュリティの警告」バーが出たら **[コンテンツの有効化]** をクリック
3. 「検索」シートが表示されればインストール成功です（初回起動時は、保存先の親フォルダをたずねるダイアログが出ることがあります。STEP 4.5 のとおり `C:\KnowledgeMgr` のまま **[OK]** を押してください）
4. シートタブで **M-05 ナレッジ登録** を開き、フォーマットを選んで各フィールドを入力し **[登録]** を押すと、`data\` フォルダに `<FormatID>-0001.txt` の形でファイルが生成されることを確認します

### 8.2 管理.xlsm

1. `C:\KnowledgeMgr\管理.xlsm` をダブルクリックで開く
2. **[コンテンツの有効化]** をクリック
3. 起動シート **M-02** が表示されればインストール成功です
4. シートタブで **M-10 格納先設定** を開き、6 つの保存先フォルダのパスと動作設定（`debugLevel` ほか）が入っていることを確認します。値を変えたときはセルを直して **[Ctrl]+[S]** で上書き保存します（専用の保存ボタンはありません）。なお「設定」シート（旧 M-11）は v2.3 で廃止し、`debugLevel` は M-10 に統合しました

### 8.3 うまく動かないとき

次の点を順に確認してください。

- LOG シートにエラーの記録が残っていないか
- 画面定義ファイルが `ui\検索\` / `ui\管理\` に配置されているか
- モジュールの取り込み漏れや、STEP 1 の事前設定漏れが無いか
- 詳しくは [困ったとき](troubleshooting.md) を参照してください

---

## データの引き継ぎ

以前のバージョンや別の PC からナレッジを引き継ぐ場合は、`data\` フォルダのナレッジファイルと `formats\` フォルダのフォーマット定義ファイルを、新しい環境の同じフォルダにコピーします。コピー前には元のフォルダのバックアップを取ってください。

---

**更新日**: 2026-06-19 JST
