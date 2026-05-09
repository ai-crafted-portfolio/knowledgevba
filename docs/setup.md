---
title: セットアップガイド
description: ナレッジ管理ツール (VBA 製) の初回導入手順。bat/ps1 をコピペで作成し、Excel に紐づけて使えるようにする
tags:
  - excel
  - vba
  - setup
---

# セットアップガイド

knowledgevba をあなたの Excel ブックに紐づけて、「ナレッジ登録 / 検索」のボタンが並んだ画面が使えるようにするための手順です。

下の **STEP 1 〜 STEP 5** の順に進めれば、初めての方でも 10〜15 分ほどで完了します。

!!! info "このページの方針"
    インストーラの bat / ps1 はこのドキュメント内に **全文をコードブロックで掲載** しています。各コードブロック右上の **コピーボタン** を押し、メモ帳に貼り付けて保存してください。
    ダウンロードリンク・GitHub Releases・共有フォルダ等の **外部経由の配布は行いません** (ADR-0013)。

---

## 必要なもの

- Windows 11 か Windows 10 の PC（Windows 標準機能だけで動作します）
- Microsoft Excel が入っていること（バージョンは何でも OK、Windows 版推奨）
- 書き込み権限のあるローカルフォルダ（ネットワークドライブ直下は避ける）

GitHub アカウント・Python・PowerShell の知識はいりません。

---

## STEP 1. Excel の信頼設定（最初の 1 回だけ）

このインストーラは VBA（Excel の中で動くマクロ）を、あなたのブックに自動で取り込みます。Excel の初期設定だと「VBA の自動取り込み」がブロックされているので、**最初に 1 か所だけ** チェックを入れます。

1. Excel を起動
2. `[ファイル] → [オプション]` を開く
3. 左メニュー一番下の `[トラスト センター]` をクリック
4. 右側の `[トラスト センターの設定(T)…]` ボタンを押す
5. 開いたウィンドウの左メニューで `[マクロの設定]` を選択
6. `[VBA プロジェクト オブジェクト モデルへのアクセスを信頼する]` に **チェックを入れる**
7. 両方のウィンドウを `OK` で閉じる

!!! tip "場所が見つからないとき"
    `[オプション]` が画面左下に見えない場合、左メニューを一番下までスクロールしてください。`[アカウント]` の下に `[オプション]` があります。

この設定は Excel に保存されるので、**次回以降は STEP 1 をスキップ** できます。

---

## STEP 2. インストーラ用フォルダを作成

任意の場所に作業フォルダを作ります。例:

- `C:\Users\<あなた>\Desktop\knowledgevba_installer\`

以降の STEP 3 / STEP 4 では、このフォルダの中に **bat と ps1 の 2 ファイル** をコピペで作成します。

---

## STEP 3. `Install-KnowledgevbaModules.bat` を作成

以下のコードブロック右上の **コピーボタン** を押してクリップボードにコピーし、メモ帳に貼り付け、STEP 2 で作ったフォルダに `Install-KnowledgevbaModules.bat` という名前で **文字コード ANSI** を選んで保存してください。

!!! warning "保存時の注意"
    - ファイル名: `Install-KnowledgevbaModules.bat`（拡張子は必ず `.bat`）
    - 文字コード: メモ帳の場合 `[名前を付けて保存]` ダイアログ下部で **`ANSI`** を選択
    - ファイルの種類: 「すべてのファイル」にしてから上記ファイル名を入力

```bat
@echo off
setlocal

REM ============================================================
REM  Install-KnowledgevbaModules.bat
REM
REM  Drag & drop your target .xlsm onto this bat, OR run from cmd:
REM    Install-KnowledgevbaModules.bat "C:\path\to\book.xlsm" [/demo]
REM
REM  Prerequisite (one-time, on the target Excel):
REM    File -> Options -> Trust Center -> Trust Center Settings ->
REM      Macro Settings -> "Trust access to the VBA project object model" ON
REM ============================================================

if "%~1"=="" (
    echo [ERROR] Drop your target .xlsm onto this bat, or pass it as arg 1.
    echo.
    echo Usage:
    echo   Install-KnowledgevbaModules.bat "C:\path\to\book.xlsm"
    echo   Install-KnowledgevbaModules.bat "C:\path\to\book.xlsm" /demo
    pause
    exit /b 1
)

set "TARGET_XLSM=%~1"
set "DEMO_FLAG="
if /I "%~2"=="/demo" set "DEMO_FLAG=-Demo"

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%Install-KnowledgevbaModules.ps1"

if not exist "%PS1%" (
    echo [ERROR] PowerShell script not found: %PS1%
    pause
    exit /b 1
)

echo ============================================================
echo  knowledgevba VBA module installer
echo ============================================================
echo  Target xlsm: %TARGET_XLSM%
echo  Demo mode  : %DEMO_FLAG%
echo  PowerShell : %PS1%
echo ============================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -XlsmPath "%TARGET_XLSM%" %DEMO_FLAG%

echo.
echo ============================================================
echo  Done. Check console output above for any warnings.
echo ============================================================
pause
```

---

## STEP 4. `Install-KnowledgevbaModules.ps1` を作成

同じく以下のコードブロック右上のコピーボタンでコピーし、メモ帳に貼り付け、**同じフォルダ** に `Install-KnowledgevbaModules.ps1` という名前で **文字コード UTF-8 (BOM 付き)** を選んで保存してください。

!!! warning "保存時の注意"
    - ファイル名: `Install-KnowledgevbaModules.ps1`（拡張子は必ず `.ps1`）
    - 文字コード: メモ帳の場合 `[名前を付けて保存]` ダイアログ下部で **`UTF-8 (BOM 付き)`** を選択
    - 同じフォルダ（STEP 3 で bat を保存したフォルダ）に保存

```powershell
# ===================================================================
# Install-KnowledgevbaModules.ps1
#   Imports VBA modules into a target .xlsm via Excel COM,
#   then runs SetupSheetsAndButtons to generate sheets + buttons.
#
# Usage:
#   .\Install-KnowledgevbaModules.ps1 -XlsmPath "C:\path\to\book.xlsm" [-Demo]
#
# Prerequisite (one-time, on the target Excel):
#   File -> Options -> Trust Center -> Trust Center Settings ->
#     Macro Settings -> "Trust access to the VBA project object model" ON
#   (Without this, VBProject.VBComponents.Import will fail.)
# ===================================================================

param(
    [Parameter(Mandatory=$true, HelpMessage="Target .xlsm file to import modules into.")]
    [string]$XlsmPath,

    [switch]$Demo,

    [string]$BaseUrl = "https://raw.githubusercontent.com/ai-crafted-portfolio/knowledgevba/main/docs/source",
    [string]$WorkDir = "$env:TEMP\knowledgevba_modules",
    [switch]$KeepWorkDir
)

$ErrorActionPreference = "Stop"

# -------------------------------------------------------------------
# Module definitions
# -------------------------------------------------------------------
$modulesProduction = @(
    @{SitePath="entrypoint/modentrymain";       Name="modEntryMain";       Type="bas"},
    @{SitePath="entrypoint/modentrysearch";     Name="modEntrySearch";     Type="bas"},
    @{SitePath="entrypoint/modentryknowledge";  Name="modEntryKnowledge";  Type="bas"},
    @{SitePath="entrypoint/modentryformat";     Name="modEntryFormat";     Type="bas"},
    @{SitePath="entrypoint/modentrysettings";   Name="modEntrySettings";   Type="bas"},
    @{SitePath="entrypoint/modspecexamples";    Name="modSpecExamples";    Type="bas"},
    @{SitePath="business-logic/clssearchengine";    Name="clsSearchEngine";    Type="cls"},
    @{SitePath="business-logic/clsknowledgemanager";Name="clsKnowledgeManager";Type="cls"},
    @{SitePath="business-logic/clsformatmanager";   Name="clsFormatManager";   Type="cls"},
    @{SitePath="business-logic/clstaskcontroller";  Name="clsTaskController";  Type="cls"},
    @{SitePath="business-logic/clsstorageresolver"; Name="clsStorageResolver"; Type="cls"},
    @{SitePath="business-logic/clsfieldmigrator";   Name="clsFieldMigrator";   Type="cls"},
    @{SitePath="business-logic/clsformspec";        Name="clsFormSpec";        Type="cls"},
    @{SitePath="business-logic/clscontrolspec";     Name="clsControlSpec";     Type="cls"},
    @{SitePath="business-logic/modformbuilder";     Name="modFormBuilder";     Type="bas"},
    @{SitePath="utility/modcommon";       Name="modCommon";       Type="bas"},
    @{SitePath="utility/modfileio";       Name="modFileIO";       Type="bas"},
    @{SitePath="utility/moddateutil";     Name="modDateUtil";     Type="bas"},
    @{SitePath="utility/modstringutil";   Name="modStringUtil";   Type="bas"},
    @{SitePath="utility/modimagerender";  Name="modImageRender";  Type="bas"},
    @{SitePath="utility/modformatcolumns";Name="modFormatColumns";Type="bas"},
    @{SitePath="infrastructure/modsetup"; Name="modSetup"; Type="bas"},
    @{SitePath="logger/clslogger";  Name="clsLogger";  Type="cls"},
    @{SitePath="logger/clslogentry";Name="clsLogEntry";Type="cls"},
    @{SitePath="special/thisworkbook"; Name="ThisWorkbook"; Type="thisworkbook"}
)

$modulesDemo = $modulesProduction | Where-Object { $_.SitePath -ne "special/thisworkbook" }
$modulesDemo += @(
    @{SitePath="cowork-demo/thisworkbook"; Name="ThisWorkbook"; Type="thisworkbook"},
    @{SitePath="cowork-demo/modautoinit";  Name="modAutoInit";  Type="bas"},
    @{SitePath="cowork-demo/moddemoseeder";Name="modDemoSeeder";Type="bas"}
)

$modules = if ($Demo) { $modulesDemo } else { $modulesProduction }

# -------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------
function Get-CodeFromMarkdown {
    param([string]$Markdown)
    $pattern = '(?ms)```(?:vba|vbnet|VBA|VBNET|vb)\b[^\r\n]*\r?\n(.*?)\r?\n```'
    $m = [regex]::Match($Markdown, $pattern)
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

function Sanitize-ForCp932 {
    param([string]$Code)
    $sjis = [System.Text.Encoding]::GetEncoding("shift_jis")
    $Code = $Code -replace [char]0x2028, "`n"
    $Code = $Code -replace [char]0x2029, "`n"
    $Code = $Code -replace [char]0xFFFD, ""
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Code.ToCharArray()) {
        $bytes = $sjis.GetBytes($ch)
        $rt = $sjis.GetString($bytes)
        if ($rt -ne $ch) {
            [void]$sb.Append('?')
        } else {
            [void]$sb.Append($ch)
        }
    }
    $Code = $sb.ToString()
    $Code = $Code -replace "`r`n", "`n"
    $Code = $Code -replace "`r", "`n"
    $Code = $Code -replace "`n", "`r`n"
    return $Code
}

function Build-BasFile {
    param([string]$ModuleName, [string]$Code, [string]$OutPath)
    $code = (Sanitize-ForCp932 -Code $Code).Trim("`r","`n")
    $hasAttribute = $code -match '^\s*Attribute\s+VB_Name'
    if (-not $hasAttribute) {
        $code = "Attribute VB_Name = `"$ModuleName`"`r`n" + $code
    }
    $sjis = [System.Text.Encoding]::GetEncoding("shift_jis")
    [System.IO.File]::WriteAllText($OutPath, $code, $sjis)
}

function Build-ClsFile {
    param([string]$ModuleName, [string]$Code, [string]$OutPath)
    $code = (Sanitize-ForCp932 -Code $Code).Trim("`r","`n")
    $hasHeader = $code -match '^\s*VERSION\s+1\.0\s+CLASS'
    if (-not $hasHeader) {
        $header = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "$ModuleName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

"@
        $code = $header + $code
    }
    $sjis = [System.Text.Encoding]::GetEncoding("shift_jis")
    [System.IO.File]::WriteAllText($OutPath, $code, $sjis)
}

# -------------------------------------------------------------------
# Validate inputs
# -------------------------------------------------------------------
if (-not (Test-Path $XlsmPath)) {
    Write-Error "Xlsm not found: $XlsmPath"
    exit 1
}
$XlsmFull = (Resolve-Path $XlsmPath).Path

# -------------------------------------------------------------------
# Phase 1: Download + extract code
# -------------------------------------------------------------------
Write-Host "==== Phase 1: Fetch $($modules.Count) modules ===="
New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null

$basClsFiles = @()
$thisWorkbookCode = $null

foreach ($m in $modules) {
    $url = "$BaseUrl/$($m.SitePath).md"
    Write-Host ("  [{0,-22}] {1}" -f $m.Name, $url)
    try {
        $resp = Invoke-WebRequest -Uri $url -UseBasicParsing
    } catch {
        Write-Warning "    fetch failed: $_"
        continue
    }
    $code = Get-CodeFromMarkdown -Markdown $resp.Content
    if (-not $code) {
        Write-Warning "    no code block detected"
        continue
    }

    switch ($m.Type) {
        "bas" {
            $out = Join-Path $WorkDir "$($m.Name).bas"
            Build-BasFile -ModuleName $m.Name -Code $code -OutPath $out
            $basClsFiles += @{Name=$m.Name; Path=$out}
        }
        "cls" {
            $out = Join-Path $WorkDir "$($m.Name).cls"
            Build-ClsFile -ModuleName $m.Name -Code $code -OutPath $out
            $basClsFiles += @{Name=$m.Name; Path=$out}
        }
        "thisworkbook" {
            $thisWorkbookCode = Sanitize-ForCp932 -Code $code
            $sjis = [System.Text.Encoding]::GetEncoding("shift_jis")
            [System.IO.File]::WriteAllText((Join-Path $WorkDir "ThisWorkbook.code.txt"), $thisWorkbookCode, $sjis)
        }
    }
}

# -------------------------------------------------------------------
# Phase 2: Open Excel COM
# -------------------------------------------------------------------
Write-Host ""
Write-Host "==== Phase 2: Open Excel COM ===="
Write-Host "  XlsmPath: $XlsmFull"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

try {
    $wb = $excel.Workbooks.Open($XlsmFull)

    Write-Host ""
    Write-Host "==== Phase 3: Import $($basClsFiles.Count) modules ===="
    foreach ($f in $basClsFiles) {
        try {
            $existing = $wb.VBProject.VBComponents.Item($f.Name)
            if ($existing) {
                $wb.VBProject.VBComponents.Remove($existing) | Out-Null
                Write-Host "  - removed existing $($f.Name)"
            }
        } catch {
        }
        try {
            $wb.VBProject.VBComponents.Import($f.Path) | Out-Null
            Write-Host "  + imported $($f.Name)"
        } catch {
            Write-Warning "    import failed for $($f.Name): $_"
        }
    }

    if ($thisWorkbookCode) {
        Write-Host ""
        Write-Host "==== Phase 4: Replace ThisWorkbook code module ===="
        try {
            $tw = $wb.VBProject.VBComponents.Item("ThisWorkbook").CodeModule
            $linesBefore = $tw.CountOfLines
            if ($linesBefore -gt 0) { $tw.DeleteLines(1, $linesBefore) }
            $tw.AddFromString($thisWorkbookCode)
            Write-Host "  ThisWorkbook code replaced ($linesBefore lines -> $($tw.CountOfLines) lines)"
        } catch {
            Write-Warning "  ThisWorkbook replace failed: $_"
        }
    }

    Write-Host ""
    Write-Host "==== Phase 5: Save xlsm + close + reopen ===="
    $wb.Save()
    Write-Host "  Saved: $XlsmFull"
    $wb.Close($true)
    Start-Sleep -Milliseconds 500
    $excel.EnableEvents = $false
    $wb = $excel.Workbooks.Open($XlsmFull)
    Write-Host "  Reopened with events disabled (Workbook_Open suppressed)."

    Write-Host ""
    Write-Host "==== Phase 6: Run SetupSheetsAndButtons ===="
    try {
        $macroQualified = "'" + $wb.Name + "'!SetupSheetsAndButtons"
        $excel.Run($macroQualified, $true) | Out-Null
        Write-Host "  Setup completed (silent=True)."
    } catch {
        Write-Warning "  Setup failed (you can run it manually via Alt+F8 -> SetupSheetsAndButtons): $_"
    }

    $wb.Save()
    $wb.Close()
} catch {
    Write-Error "Workflow error: $_"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    if (-not $KeepWorkDir) {
        Remove-Item $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host "==== DONE ===="
Write-Host "Open $XlsmFull in Excel and verify."
```

---

## STEP 5. インストール対象の `.xlsm` を準備

インストール対象の Excel マクロ有効ブック (`.xlsm`) を用意します。

- 既に `.xlsm` を持っている場合: そのまま STEP 6 へ
- まだ無い場合:
    1. Excel で空白のブックを新規作成
    2. `[ファイル] → [名前を付けて保存]`
    3. ファイルの種類で **「Excel マクロ有効ブック (`*.xlsm`)」** を選択
    4. 任意のファイル名（例: `ナレッジ管理.xlsm`）で保存

!!! danger "拡張子に注意"
    `.xlsx` には VBA を入れられません。**必ず `.xlsm`** で保存してください。

---

## STEP 6. インストーラを実行

1. 対象 `.xlsm` を Excel で開いていれば **閉じる**（二重起動を避けるため）
2. エクスプローラで STEP 2 の作業フォルダを開く
3. 対象 `.xlsm` を **`Install-KnowledgevbaModules.bat` の上にドラッグ＆ドロップ**
4. 黒いコンソール窓が開き、自動で次の処理が走ります（30 秒〜1 分）:
    - VBA モジュールの取得
    - 対象 `.xlsm` への取り込み
    - 14 シート + 8〜29 ボタンの自動生成
5. 最後に `==== DONE ====` と出て `続行するには何かキーを押してください . . .` で待機すれば成功です。Enter で閉じます

!!! warning "コマンドプロンプトから実行する場合"
    ```cmd
    cd C:\Users\<あなた>\Desktop\knowledgevba_installer
    Install-KnowledgevbaModules.bat "C:\path\to\your_book.xlsm"
    ```

!!! tip "デモモード"
    `Workbook_Open` で自動的にシート生成 + デモナレッジを投入する版を試すなら:
    ```cmd
    Install-KnowledgevbaModules.bat "C:\path\to\demo_book.xlsm" /demo
    ```

---

## STEP 7. 完了した `.xlsm` を開いて動作確認

1. インストール完了した `.xlsm` を **ダブルクリックで開く**
2. 上部に黄色い `[セキュリティの警告 マクロが無効にされました]` バーが出たら、`[コンテンツの有効化]` をクリック
3. メインシートに `[初回セットアップ]` `[設定変更]` `[ナレッジ登録]` `[検索]` などのボタンが並んで表示されれば成功

---

## STEP 8. 「初回セットアップ」ボタンを押す

メインシートの一番上の **`[▶初回セットアップ]`** ボタンを **1 回だけ** 押してください。これで以下が実行されます:

- 必要シート 14 個（`メイン`, `検索`, `ログ` 等）の自動生成
- ログシートのヘッダー行書込
- 設定シート C6 にテストモードフラグ（`"FALSE"` = 本番モード）を設定
- メインシート B25:E26 にタスクボタン用ラベル書込
- フォームコントロールボタン 29 個の全シート配置 + マクロ割り当て
- メインシート以外を非表示
- 既定の空 Sheet1 削除

完了ダイアログ「セットアップ完了。」が出たら成功。`Ctrl + S` で保存してください。

!!! info "冪等性"
    `初回セットアップ` は何度押しても安全です（既存ボタンは削除→再配置、シートは存在チェック後スキップ）。

---

## トラブルシューティング

### Q1. 「access to the VBA Project is not trusted」エラー

STEP 1 のチェックボックスが入っていません。Excel に戻って STEP 1 をやり直してください。

### Q2. 黒い窓が一瞬出てすぐ消える

bat をダブルクリックしただけだと、エラーメッセージが見えずに閉じます。**`.xlsm` を bat にドラッグ＆ドロップ** で起動してください。bat 単体起動は使い方として想定していません。

### Q3. 「コンパイルエラー: ステートメントの最後」が VBE で出る

CRLF / CP932 サニタイズが効いていません。STEP 4 の ps1 を最新版（このページのコードブロック）でコピペし直してください。

### Q4. もう一度 STEP 6 を実行しても大丈夫？

何度実行しても OK（idempotent）。同じ `.xlsm` を再度ドラッグすれば最新の VBA に上書きされます。**シート上のデータ（あなたが登録したナレッジ）は消えません**。

### Q5. 「コンパイルエラー: メソッドまたはデータ メンバーが見つかりません」

どこかのモジュールがインポート漏れの可能性。`Alt + F11` で VBE を開き、左ペインで以下が全て揃っているか確認:

- 標準モジュール (`mod*`): 12 個
- クラスモジュール (`cls*`): 11 個
- Microsoft Excel Objects: `ThisWorkbook`（中身が貼られている）

不足していれば STEP 6 を再実行してください。

### Q6. 起動時に「インデックスが有効範囲にありません」

STEP 8 の `[▶初回セットアップ]` を実行していない可能性。再度押してください。

### Q7. ボタンが既存の表示と重なって見えにくい

新規 `.xlsm` の既定の列幅・行高さのため、ボタンが小さく見える場合があります。各シートの該当行（例: メイン 25-26 行目）の高さを増やすか、列幅を広げると見栄えが改善します。

---

## セットアップが触る範囲・触らない範囲

| 項目 | セットアップマクロが触るか |
|---|---|
| 必要シート 14 個の作成 | あり（既存スキップ、不在分のみ作成） |
| ログシート 1 行目ヘッダー | あり（上書き） |
| 設定シート C6（テストモード） | あり（`"FALSE"` 設定） |
| メインシート B24:E26 | あり（案内テキスト + ボタン用ラベル書込） |
| 各シートのフォームコントロールボタン | あり（同名は削除→再配置、29 個） |
| 各シートの初期可視性 | あり（メインのみ表示、他は非表示） |
| 空の既定 Sheet1 | あり（削除、データ・シェイプがあれば残す） |
| その他のセル内容 | なし |
| ユーザフォーム | なし |
| Excel シートの書式・色 | なし |

---

## 関連

- [操作マニュアル](operations.md) — 日常操作（ナレッジ登録・検索・編集 等）
- [仕様](spec.md) — フォーマット定義・検索仕様
- ADR-0008: VBA モジュール配布アーキテクチャ
- ADR-0013: 配布チャネルはコードブロック全文掲載 + コピペ保存に統一
