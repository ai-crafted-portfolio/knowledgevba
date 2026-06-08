---
title: インストール手順
description: knowledgevba をはじめて使うときのセットアップ手順
---

# インストール

knowledgevba は次の 3 つの Excel ブックで構成されます。

| ブック | 役割 |
|---|---|
| `登録修正.xlsm` | ナレッジを新しく記録する / 既存ナレッジを修正・削除する |
| `検索.xlsm` | ナレッジを検索して内容を表示する |
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
3. **STEP 3**: 作業フォルダの直下に 3 つの空ブックを作る（`登録修正.xlsm` / `検索.xlsm` / `管理.xlsm`）
4. **STEP 4**: 作業フォルダの直下にデータ用 4 フォルダを作る
5. **STEP 4.5**: 設定ファイル（`管理_config.txt` 等 3 ファイル）を手動で配置する
6. **STEP 5**: `installer\` フォルダを作り、3 つの bat と 1 つの ps1 を保存する
7. **STEP 6**: VBA モジュールのフォルダ（`installer\vba_modules\`）を用意する
8. **STEP 7**: 3 つの bat を順番にダブルクリックして実行する
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

## STEP 3: 3 つの空ブックを作る

作業フォルダ（例: `C:\KnowledgeMgr\`）の直下に、次の 3 つの空ブックを作成します。Excel で新規ブックを開き、ファイル名を変えて `.xlsm` 形式で保存してください。

```
C:\KnowledgeMgr\
  ├─ 登録修正.xlsm
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

## STEP 4.5: 設定ファイル（config.txt）を手動で配置する

各ブックは起動時に **同じ作業フォルダ直下** にある「ブック名_config.txt」を読み込みます。ブックごとに 1 ファイル、合計 3 ファイルを手動で作成します。

!!! info "なぜ手動なのか"
    インストーラ（bat / ps1）は、設定ファイルを自動生成しません。利用者ごとに保存先パスやログレベルを変えられるよう、テキストファイルを手で配置する方式に統一しています。

### 4.5.1 配置先と名前

作業フォルダ（例: `C:\KnowledgeMgr\`）の直下に、次の 3 ファイルを置きます。

```
C:\KnowledgeMgr\
  ├─ 管理_config.txt
  ├─ 登録修正_config.txt
  └─ 検索_config.txt
```

### 4.5.2 ファイル本体（3 ブック共通の標準値）

3 ファイルとも下の内容で問題ありません。後から保存先や `debugLevel` を変えたい場合は、メモ帳で直接書き換えてください（編集後はブックを開き直すと反映されます）。

```ini
[CONFIG]
debugLevel=INFO
logRotationRows=10000
data_dir=C:\KnowledgeMgr\data\
format_dir=C:\KnowledgeMgr\formats\
ui_dir=C:\KnowledgeMgr\ui\
backup_dir=C:\KnowledgeMgr\data\backup\
uiSchemaFailMode=safeDefault
systemSheetVisibility=Hidden
autoReloadOnStartup=TRUE
migrateBackupEnabled=TRUE
```

### 4.5.3 保存時の注意

- メモ帳で開き、**[名前を付けて保存]** から保存します
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（UTF-8 にすると日本語名のパスや値が文字化けします）
- 改行コード: 既定の CRLF のままで構いません

### 4.5.4 設定ファイルを置かなかった場合

設定ファイルが見つからない場合、ブックは上記の標準値で起動します（エラーにはなりません）。ただし、保存先パスをカスタマイズしたい場合や、`debugLevel` を変えてログを抑えたい場合は配置してください。

### 4.5.5 主な設定項目

| キー | 内容 | 既定値 |
|---|---|---|
| `debugLevel` | ログシートに出す情報量。`ERROR` / `WARN` / `INFO` / `DEBUG` から選択 | `INFO` |
| `logRotationRows` | LOG シートが何行を超えたら古い行を自動削除するか | `10000` |
| `data_dir` | ナレッジ本体（.txt）の保存先フォルダ | `C:\KnowledgeMgr\data\` |
| `format_dir` | フォーマット定義（.txt）の保存先フォルダ | `C:\KnowledgeMgr\formats\` |
| `ui_dir` | 画面定義（.txt）の保存先フォルダ | `C:\KnowledgeMgr\ui\` |
| `backup_dir` | 修正・削除時のバックアップ保存先 | `C:\KnowledgeMgr\data\backup\` |
| `uiSchemaFailMode` | 画面定義に不整合があった時の動作。`safeDefault` 推奨 | `safeDefault` |
| `systemSheetVisibility` | システム用シートの表示状態。`Hidden` / `Visible` | `Hidden` |
| `autoReloadOnStartup` | 起動時に自動で設定を再読み込みするか | `TRUE` |
| `migrateBackupEnabled` | フォーマット変更時にナレッジ自動マイグレーションを行うか | `TRUE` |

---

## STEP 5: installer フォルダを作る

作業フォルダの直下に `installer\` フォルダを作り、次の 4 ファイルを保存します。

```
C:\KnowledgeMgr\
  └─ installer\
       ├─ install_register.bat
       ├─ install_search.bat
       ├─ install_admin.bat
       └─ _auto_install.ps1
```

### 5.1 install_register.bat

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\`
- ファイル名: `install_register.bat`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（UTF-8 にすると日本語が文字化けします）

```bat
@echo off
chcp 932 > nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

REM ====================================================================
REM  knowledgevba v2.3  install_register.bat
REM  - 登録修正.xlsm に common\ + register\ の VBA モジュールを
REM    自動 Import し、続けてセットアップマクロ Setup_register を
REM    実行する。利用者はダブルクリック 1 回で完了。
REM  - 文字コード: CP932 (Shift_JIS) / 改行: CRLF
REM ====================================================================

REM === Step 1: ターゲット .xlsm パスを固定 (本 bat は登録修正.xlsm 専用) ===
set "XLSM_PATH=%~dp0..\登録修正.xlsm"

REM === Step 2: ファイル存在チェック ===
if not exist "!XLSM_PATH!" (
    echo [ERROR] 登録修正.xlsm が見つかりません: !XLSM_PATH!
    echo         install bat と同じ階層の上 ^( C:\KnowledgeMgr\ ^) に
    echo         登録修正.xlsm を作成してください ^(本書 §16.4^)。
    pause
    exit /b 1
)

REM === Step 3: Excel 起動チェック (残留 EXCEL.EXE と衝突しない) ===
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
        pause
        exit /b 1
    )
    echo [INFO] EXCEL.EXE 終了確認 OK。処理を続行します。
) else (
    echo [INFO] EXCEL.EXE は起動していません。処理を続行します。
)

REM === Step 4: vba_modules\common と register の存在チェック ===
if not exist "%~dp0vba_modules\common" (
    echo [ERROR] vba_modules\common フォルダが見つかりません: %~dp0vba_modules\common
    pause
    exit /b 1
)
if not exist "%~dp0vba_modules\register" (
    echo [ERROR] vba_modules\register フォルダが見つかりません: %~dp0vba_modules\register
    pause
    exit /b 1
)

REM === Step 5: PowerShell 本体を起動 ===
echo [INFO] install_register.bat: PowerShell 起動 (Role=register)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_auto_install.ps1" -XlsmPath "!XLSM_PATH!" -Role register
set "PS_EXIT=!errorlevel!"

if !PS_EXIT! NEQ 0 (
    echo [FAIL] install_register failed. exit_code=!PS_EXIT!
    echo        詳しいエラーは上の PowerShell 出力を参照。
    echo        トラブルシュート: 本書 §16.7
) else (
    echo [DONE] Install + setup succeeded.
    echo        Excel で 登録修正.xlsm を開くと、起動シート M-05 が表示されることを
    echo        確認してください ^(本書 §16.7 動作確認^)。
)
pause
exit /b !PS_EXIT!
```

### 5.2 install_search.bat

同様に、次のコードを `install_search.bat` として保存します（場所・文字コードは 5.1 と同じ）。

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
    pause
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
        pause
        exit /b 1
    )
    echo [INFO] EXCEL.EXE 終了確認 OK。処理を続行します。
) else (
    echo [INFO] EXCEL.EXE は起動していません。処理を続行します。
)

if not exist "%~dp0vba_modules\common" (
    echo [ERROR] vba_modules\common フォルダが見つかりません: %~dp0vba_modules\common
    pause
    exit /b 1
)
if not exist "%~dp0vba_modules\search" (
    echo [ERROR] vba_modules\search フォルダが見つかりません: %~dp0vba_modules\search
    pause
    exit /b 1
)

echo [INFO] install_search.bat: PowerShell 起動 (Role=search)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_auto_install.ps1" -XlsmPath "!XLSM_PATH!" -Role search
set "PS_EXIT=!errorlevel!"

if !PS_EXIT! NEQ 0 (
    echo [FAIL] install_search failed. exit_code=!PS_EXIT!
    echo        トラブルシュート: 本書 §16.7
) else (
    echo [DONE] Install + setup succeeded.
    echo        Excel で 検索.xlsm を開くと、起動シート M-08 が表示されることを
    echo        確認してください ^(本書 §16.7 動作確認^)。
)
pause
exit /b !PS_EXIT!
```

### 5.3 install_admin.bat

同様に、次のコードを `install_admin.bat` として保存します。

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
REM    シート構成は M-02 / M-03 / M-04 / M-10 / M-11 / M-12 / M-14 / ログ
REM    の 8 シート。M-12 は「フォーマット変更チェック」に改名・縮小
REM    (ADR-0072 §2.4)。
REM ====================================================================

set "XLSM_PATH=%~dp0..\管理.xlsm"

if not exist "!XLSM_PATH!" (
    echo [ERROR] 管理.xlsm が見つかりません: !XLSM_PATH!
    echo         install bat と同じ階層の上 ^( C:\KnowledgeMgr\ ^) に
    echo         管理.xlsm を作成してください ^(本書 §16.4^)。
    pause
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
        pause
        exit /b 1
    )
    echo [INFO] EXCEL.EXE 終了確認 OK。処理を続行します。
) else (
    echo [INFO] EXCEL.EXE は起動していません。処理を続行します。
)

if not exist "%~dp0vba_modules\common" (
    echo [ERROR] vba_modules\common フォルダが見つかりません: %~dp0vba_modules\common
    pause
    exit /b 1
)
if not exist "%~dp0vba_modules\admin" (
    echo [ERROR] vba_modules\admin フォルダが見つかりません: %~dp0vba_modules\admin
    pause
    exit /b 1
)

echo [INFO] install_admin.bat: PowerShell 起動 (Role=admin)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_auto_install.ps1" -XlsmPath "!XLSM_PATH!" -Role admin
set "PS_EXIT=!errorlevel!"

if !PS_EXIT! NEQ 0 (
    echo [FAIL] install_admin failed. exit_code=!PS_EXIT!
    echo        トラブルシュート: 本書 §16.7
) else (
    echo [DONE] Install + setup succeeded.
    echo        Excel で 管理.xlsm を開くと、起動シート M-02 が表示されることを
    echo        確認してください ^(本書 §16.7 動作確認^)。
)
pause
exit /b !PS_EXIT!
```

### 5.4 _auto_install.ps1

3 つの bat から共通で呼ばれる PowerShell 本体です。下のコードを `_auto_install.ps1` として保存します。

- 場所: `C:\KnowledgeMgr\installer\`
- ファイル名: `_auto_install.ps1`
- ファイルの種類: **すべてのファイル**
- 文字コード: **UTF-8 (BOM 付き)**

```powershell
�ｻｿ[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$XlsmPath,

    [Parameter(Mandatory = $true)]
    [ValidateSet('register', 'search', 'admin')]
    [string]$Role
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# ====================================================================
#  knowledgevba v2.3  _auto_install.ps1
#  - install_register.bat / install_search.bat / install_admin.bat
#    縺九ｉ蜈ｱ騾壹〒蜻ｼ縺ｰ繧後ｋ譛ｬ菴薙せ繧ｯ繝ｪ繝励ヨ縲�
#  - 菫晏ｭ俶凾縺ｮ譁�蟄励さ繝ｼ繝�: UTF-8 (BOM 莉倥″) / 謾ｹ陦�: CRLF
#  - 驟咲ｽｮ: C:\KnowledgeMgr\installer\_auto_install.ps1
# ====================================================================

function Write-Log($msg) {
    Write-Host ('[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $msg)
}

# ---- Step 1. 繝代せ讀懆ｨｼ ----
if (-not (Test-Path -LiteralPath $XlsmPath)) {
    Write-Host "[ERROR] 繧ｿ繝ｼ繧ｲ繝�繝� .xlsm 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ: $XlsmPath" -ForegroundColor Red
    exit 1
}
$target = (Resolve-Path -LiteralPath $XlsmPath).Path
$ext = [IO.Path]::GetExtension($target).ToLowerInvariant()
if ($ext -ne '.xlsm') {
    Write-Host "[ERROR] .xlsm 縺ｧ縺ｯ縺ゅｊ縺ｾ縺帙ｓ (諡｡蠑ｵ蟄� $ext)縲よ悽譖ｸ ﾂｧ16.4 縺ｧ .xlsm 繧剃ｽ懈�舌＠縺ｦ縺上□縺輔＞縲�" -ForegroundColor Red
    exit 1
}

# ---- Step 1.5 (v2.3 2026-05-27): ui_seed 繧� config 縺ｮ ui_dir 縺ｫ閾ｪ蜍暮�榊ｸ� ----
# 驟榊ｸ�迚ｩ縺ｯ dist_v2\ui_seed\{邂｡逅�|逋ｻ骭ｲ菫ｮ豁｣|讀懃ｴ｢}\M-NN.txt 繧貞性繧縲�
# config.txt 縺ｮ ui_dir (萓�: C:\KnowledgeMgr\ui\) 逶ｴ荳九∈ <xlsmFolder>\M-NN.txt 繧�
# 繧ｳ繝斐�ｼ縺吶ｋ縲よ里蟄倥ヵ繧｡繧､繝ｫ縺ｯ荳頑嶌縺� (譌ｧ v2.2 stanza 縺ｮ鄂ｮ謠帙′逶ｮ逧�)縲�
$roleToJp = @{
    'admin'    = ([char]0x7BA1)+([char]0x7406)                     # 邂｡逅�
    'register' = ([char]0x767B)+([char]0x9332)+([char]0x4FEE)+([char]0x6B63)  # 逋ｻ骭ｲ菫ｮ豁｣
    'search'   = ([char]0x691C)+([char]0x7D22)                     # 讀懃ｴ｢
}
$xlsmJp = $roleToJp[$Role]
$uiSeedSrc = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) ('..\ui_seed\' + $xlsmJp)
$uiSeedSrc = [IO.Path]::GetFullPath($uiSeedSrc)
# Resolve ui_dir from config.txt (look in known locations near xlsm)
$kanriRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Definition))
$cfgPath = $null
foreach($cfgTry in @(
    (Join-Path 'C:\KnowledgeMgr\config' ($xlsmJp + '_config.txt')),
    (Join-Path 'C:\KnowledgeMgr' ($xlsmJp + '_config.txt'))
)){
    if(Test-Path -LiteralPath $cfgTry){ $cfgPath = $cfgTry; break }
}
$uiDir = $null
if($cfgPath){
    foreach($l in Get-Content -LiteralPath $cfgPath -Encoding Default){
        if($l -match '^\s*ui_dir\s*=\s*(.+?)\s*$'){
            $uiDir = $matches[1].Trim().TrimEnd('\') + '\'
            break
        }
    }
}
if($uiDir -and (Test-Path -LiteralPath $uiSeedSrc)){
    $uiDst = Join-Path $uiDir $xlsmJp
    if(-not (Test-Path -LiteralPath $uiDst)){
        New-Item -ItemType Directory -Path $uiDst -Force | Out-Null
    }
    Get-ChildItem -LiteralPath $uiSeedSrc -File | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $uiDst -Force
    }
    Write-Log ("[OK] ui_seed 驟榊ｸ�: {0} -> {1}" -f $uiSeedSrc, $uiDst)
} else {
    Write-Log ("[INFO] ui_seed 驟榊ｸ�繧ｹ繧ｭ繝�繝� (uiDir={0} src={1})" -f $uiDir, $uiSeedSrc)
}

# ---- Step 2. 繝｢繧ｸ繝･繝ｼ繝ｫ驟咲ｽｮ (common + <role>) 繧貞�玲嫌 ----
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$commonDir = Join-Path $scriptDir 'vba_modules\common'
$roleDir   = Join-Path $scriptDir ('vba_modules\' + $Role)

foreach ($d in @($commonDir, $roleDir)) {
    if (-not (Test-Path -LiteralPath $d)) {
        Write-Host "[ERROR] 繝輔か繝ｫ繝縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ: $d" -ForegroundColor Red
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

# ThisWorkbook.cls 縺ｯ Document module 縺ｪ縺ｮ縺ｧ Import 縺帙★ CodeModule.AddFromString 縺ｧ譛ｬ菴薙□縺第ｵ√＠霎ｼ繧
$thisWorkbookFile = $allFiles | Where-Object { $_.Name -eq 'ThisWorkbook.cls' } | Select-Object -First 1
$importFiles      = $allFiles | Where-Object { $_.Name -ne 'ThisWorkbook.cls' }

Write-Log ('[OK] 繝｢繧ｸ繝･繝ｼ繝ｫ蛻玲嫌: 蜈ｨ {0} 譛ｬ (蜀� ThisWorkbook 1 / Import 蟇ｾ雎｡ {1})' -f $allFiles.Count, $importFiles.Count)

# ---- Step 3. staging: source 縺ｯ CP932+CRLF (Phase E 隕冗ｴ�)縲�
# v2.3 fix (2026-05-27): source 縺� UTF-8 (BOM 譛臥┌蝠上ｏ縺�) 縺ｮ蝣ｴ蜷医ｂ蟇ｾ蠢懊�
# UTF-8 縺ｨ縺励※ valid 縺ｪ繧� UTF-8 縺ｨ縺励※ decode縲√◎縺�縺ｧ縺ｪ縺代ｌ縺ｰ CP932 縺ｨ縺励※ decode縲�
# 譖ｸ縺肴綾縺励�ｯ蟶ｸ縺ｫ CP932 + CRLF (VBE Import 縺� CP932 繧呈悄蠕�縺吶ｋ縺溘ａ)縲�
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
    $text   = $text -replace "`n", "`r`n"   # CRLF 蠑ｷ蛻ｶ
    # Encode CP932 with replacement fallback (any non-CP932 char becomes '?' but no crash)
    $outBytes = $cp932Repl.GetBytes($text)
    [IO.File]::WriteAllBytes($destFp, $outBytes)
    $stagedPaths += $destFp
}
Write-Log ('[OK] staging 螳御ｺ�: {0}' -f $staging)

# ---- Step 4. ThisWorkbook.cls 譛ｬ菴捺歓蜃ｺ ----
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
    Write-Log '[OK] ThisWorkbook 譛ｬ菴捺歓蜃ｺ 螳御ｺ�'
} else {
    Write-Host "[WARN] ThisWorkbook.cls 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ (Role=$Role)縲�" -ForegroundColor Yellow
}

# ---- Step 5. Excel COM 襍ｷ蜍� ----
Write-Log '[..] Excel COM 襍ｷ蜍�'
$excel = New-Object -ComObject Excel.Application
$excel.Visible        = $true
$excel.DisplayAlerts  = $false
$excel.EnableEvents   = $false   # Workbook_Open 證ｴ逋ｺ謚第ｭ｢
$excel.AskToUpdateLinks = $false
try { $excel.AutomationSecurity = 1 } catch {}   # msoAutomationSecurityLow

$wb = $null
try {
    $wb = $excel.Workbooks.Open($target, 0, $false)
    $vbProj = $wb.VBProject   # 竊� VBA project model 菫｡鬆ｼ縺� ON 縺ｧ縺ｪ縺�縺ｨ縲後い繧ｯ繧ｻ繧ｹ縺梧拠蜷ｦ縺輔ｌ縺ｾ縺励◆縲�

    # ---- Step 6. 譌｢蟄伜酔蜷阪Δ繧ｸ繝･繝ｼ繝ｫ蜑企勁 + 菴吝�� module strict purge (Document module = Type 100 縺ｯ繧ｹ繧ｭ繝�繝�) ----
    # iter23 (2026-06-01): canonical role 縺ｨ辟｡髢｢菫ゅ↑驕主悉谿狗蕗 module 繧貞�ｨ驛ｨ Remove 縺吶ｋ
    # strict mode 蛹悶ゅ％繧後′縺ｪ縺�縺ｨ驕主悉 install 縺ｧ豺ｷ蜈･縺励◆ modEntrySettings 遲峨′螻�蠎ｧ縺｣縺ｦ
    # Application.Run 隗｣豎ｺ陦晉ｪ√ｒ襍ｷ縺薙＠縲・2E 縺後後�槭け繝ｭ繧貞ｮ溯｡後〒縺阪∪縺帙ｓ縲阪〒螟ｱ謨励☆繧九�
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
        if ($vc.Type -eq 100) { continue }   # Document module (ThisWorkbook/Sheet*) 縺ｯ keep
        # canonical importFiles 縺ｫ蜷ｫ縺ｾ繧後ｋ name 縺ｯ蜀� Import 縺ｧ鄂ｮ謠帙☆繧九◆繧� Remove
        # canonical importFiles 縺ｫ蜷ｫ縺ｾ繧後↑縺� name 縺ｯ菴吝��谿狗蕗 竊� 蜷後§縺� Remove (strict purge)
        try {
            $vbProj.VBComponents.Remove($vc) | Out-Null
        } catch {
            Write-Log ('[WARN] Remove FAIL: ' + $modName + ' (' + $_.Exception.Message + ')')
        }
    }

    # ---- Step 6.5. deferred Remove 遒ｺ螳� (feedback_cls_deferred_remove) ----
    Write-Log '[..] Step 6.5 wb.Save() after Remove (deferred remove flush)'
    $wb.Save()
    Start-Sleep -Seconds 1
    Write-Log '[OK] Step 6.5 Save + Sleep complete'

    # ---- Step 7. .bas / .cls 繧貞�ｨ莉ｶ Import ----
    foreach ($p in $stagedPaths) {
        $vbProj.VBComponents.Import($p) | Out-Null
    }
    Write-Log ('[OK] Import 螳御ｺ� ({0} module)' -f $stagedPaths.Count)

    # ---- Step 8. ThisWorkbook 譛ｬ菴灘ｷｮ縺玲崛縺� ----
    if ($thisWorkbookBody -ne '') {
        $twb = $vbProj.VBComponents.Item('ThisWorkbook')
        $cm  = $twb.CodeModule
        if ($cm.CountOfLines -gt 0) {
            $cm.DeleteLines(1, $cm.CountOfLines) | Out-Null
        }
        $cm.AddFromString($thisWorkbookBody)
        Write-Log '[OK] ThisWorkbook 譛ｬ菴捺ｳｨ蜈･ 螳御ｺ�'
    }

    # ---- Step 9. 荳譌ｦ菫晏ｭ� (繧ｳ繝ｳ繝代う繝ｫ遒ｺ螳�) ----
    $wb.Save()

    # ---- Step 10. 繧ｻ繝�繝医い繝�繝励�槭け繝ｭ繧� Application.Run 縺ｧ襍ｷ蜍� ----
    $setupName = 'Setup_' + $Role   # Setup_register / Setup_search / Setup_admin
    Write-Log ('[..] Application.Run "{0}"' -f $setupName)
    try {
        $excel.Run($setupName) | Out-Null
        Write-Log ('[OK] {0} 螳御ｺ�' -f $setupName)
    } catch {
        Write-Host ('[ERROR] {0} 螳溯｡後↓螟ｱ謨�: {1}' -f $setupName, $_.Exception.Message) -ForegroundColor Red
        throw
    }

    # ---- Step 11. 繧ｻ繝�繝医い繝�繝怜ｾ後�ｮ迥ｶ諷九ｒ菫晏ｭ� ----
    $wb.Save()
    Write-Log '[OK] xlsm 菫晏ｭ� 螳御ｺ�'
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

Write-Log '[DONE] install + setup 螳御ｺ�'
exit 0
```

---

## STEP 6: VBA モジュールフォルダを用意する

`installer\vba_modules\` フォルダの直下に `common\` / `register\` / `search\` / `admin\` の 4 つのサブフォルダを作り、各サブフォルダに該当する `.bas` / `.cls` ファイル一式を配置します。

```
C:\KnowledgeMgr\
  └─ installer\
       └─ vba_modules\
            ├─ common\    ← 全ブック共通モジュール
            ├─ register\  ← 登録修正.xlsm 専用モジュール
            ├─ search\    ← 検索.xlsm 専用モジュール
            └─ admin\     ← 管理.xlsm 専用モジュール
```

各フォルダの中身（モジュールごとのソースコード）は **[モジュール一覧](modules/index.md)** ページに 1 ファイル 1 ページで全文掲載しています。1 モジュールずつメモ帳に貼り付け、フォルダ名どおりのファイル名（例: `modCommon.bas`、`clsKnowledgeManager.cls` 等）で保存してください。

| 配置先サブフォルダ | ファイル数 | 一覧 |
|---|---|---|
| `vba_modules\admin\` | 10 | [管理.xlsm 用モジュール一覧](modules/index.md#xlsm-installervba_modulesadmin) |
| `vba_modules\register\` | 5 | [登録修正.xlsm 用モジュール一覧](modules/index.md#xlsm-installervba_modulesregister) |
| `vba_modules\search\` | 4 | [検索.xlsm 用モジュール一覧](modules/index.md#xlsm-installervba_modulessearch) |
| `vba_modules\common\` | 64 | [共通モジュール一覧](modules/index.md#installervba_modulescommon) |

合計 **83 ファイル** あります。すべて保存し終わったら STEP 7 に進んでください。

!!! warning "モジュールファイルの保存形式"
    `.bas` / `.cls` の文字コードは **ANSI（Shift-JIS）** で保存してください。本インストーラは UTF-8 (BOM 有無いずれも) でも自動変換しますが、CP932（ANSI）が確実です。改行コードは **CRLF** にしてください。

---

## STEP 7: bat を順番に実行する

`C:\KnowledgeMgr\installer\` を開き、次の順番で bat を **ダブルクリック** します。

1. `install_register.bat`
2. `install_search.bat`
3. `install_admin.bat`

各 bat は黒いコンソールウィンドウを開き、状況を表示しながら処理を進めます。最後に `[DONE] Install + setup succeeded.` と表示され、**「続行するには何かキーを押してください . . .」** で停止したら成功です。任意のキーを押してウィンドウを閉じてください。

!!! tip "実行中に Excel が一瞬開いて閉じることがあります"
    bat は内部で Excel を非表示モードで開き、VBA モジュールを取り込んでから保存・終了します。利用者の操作は不要です。

!!! warning "途中で Excel を起動しないでください"
    bat 実行中に手動で Excel を起動すると、インストーラが起動済み Excel を強制終了 (`taskkill /F`) して進行します。作業中のブックがあれば失われる可能性があります。

bat が `[ERROR]` 等で停止した場合は、[困ったとき](troubleshooting.md) を参照してください。

---

## STEP 8: 動作確認

3 つの bat がすべて成功したら、Excel で各ブックを開いて動作を確認します。

### 8.1 登録修正.xlsm

1. `C:\KnowledgeMgr\登録修正.xlsm` をダブルクリックで開く
2. 上部に「セキュリティの警告 マクロが無効にされました」と黄色いバーが出たら、**[コンテンツの有効化]** をクリック
3. 「ナレッジ登録」シートが自動で表示されればインストール成功
4. **[新規登録]** ボタン → フォーマット選択 → 各フィールド入力 → **[登録]** ボタンの順に押し、`data\` フォルダに `<FormatID>-0001.txt` の形でファイルが生成されることを確認

### 8.2 検索.xlsm

1. `C:\KnowledgeMgr\検索.xlsm` をダブルクリックで開く
2. 「コンテンツの有効化」をクリック
3. 「検索」シートが表示�