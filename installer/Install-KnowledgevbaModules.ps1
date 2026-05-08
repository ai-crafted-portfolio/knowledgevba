# ===================================================================
# Install-KnowledgevbaModules.ps1
#   Downloads VBA modules from knowledgevba site and imports them
#   into a target .xlsm via Excel COM. Then runs SetupSheetsAndButtons.
#
# Usage:
#   .\Install-KnowledgevbaModules.ps1 -XlsmPath "C:\path\to\book.xlsm" [-Demo]
#
# Prerequisite (one-time, on the target Excel):
#   File -> Options -> Trust Center -> Trust Center Settings ->
#     Macro Settings -> "Trust access to the VBA project object model" ON
#   (Without this, VBProject.VBComponents.Import will fail.)
#
# What it does:
#   1. Downloads .md files from raw.githubusercontent.com/.../knowledgevba/main/docs/source/
#   2. Extracts the ```vba|vbnet code block from each .md
#   3. Writes .bas / .cls files into a temp work dir, encoded in Shift_JIS
#   4. Imports each module into the target xlsm (replaces same-named module)
#   5. Replaces ThisWorkbook code (cannot be Imported, must use CodeModule)
#   6. Saves xlsm
#   7. Runs SetupSheetsAndButtons (so 14 sheets + 29 buttons get generated)
#   8. Saves and quits Excel
#
# Encoding: All output files are Shift_JIS (CP932) so VBE imports them
#   as Japanese-locale source. UTF-8 is NOT used because v13/v14 have
#   user input flowing through SJIS and we made the source CP932-safe
#   (ADR 0035 / repo cp932 commits 2026-05-08).
# ===================================================================

param(
    [Parameter(Mandatory=$true, HelpMessage="Target .xlsm file to import modules into.")]
    [string]$XlsmPath,

    [switch]$Demo,   # If set, includes Cowork-demo modules (modAutoInit, modDemoSeeder, cowork-demo/thisworkbook)

    [string]$BaseUrl = "https://raw.githubusercontent.com/ai-crafted-portfolio/knowledgevba/main/docs/source",
    [string]$WorkDir = "$env:TEMP\knowledgevba_modules",
    [switch]$KeepWorkDir
)

$ErrorActionPreference = "Stop"

# -------------------------------------------------------------------
# Module definitions
# -------------------------------------------------------------------
# Each entry: SitePath (relative under /docs/source/), ModuleName, Type
# Type is one of: "bas", "cls", "thisworkbook"
$modulesProduction = @(
    # entrypoint layer
    @{SitePath="entrypoint/modentrymain";       Name="modEntryMain";       Type="bas"},
    @{SitePath="entrypoint/modentrysearch";     Name="modEntrySearch";     Type="bas"},
    @{SitePath="entrypoint/modentryknowledge";  Name="modEntryKnowledge";  Type="bas"},
    @{SitePath="entrypoint/modentryformat";     Name="modEntryFormat";     Type="bas"},
    @{SitePath="entrypoint/modentrysettings";   Name="modEntrySettings";   Type="bas"},
    @{SitePath="entrypoint/modspecexamples";    Name="modSpecExamples";    Type="bas"},
    # business-logic layer
    @{SitePath="business-logic/clssearchengine";    Name="clsSearchEngine";    Type="cls"},
    @{SitePath="business-logic/clsknowledgemanager";Name="clsKnowledgeManager";Type="cls"},
    @{SitePath="business-logic/clsformatmanager";   Name="clsFormatManager";   Type="cls"},
    @{SitePath="business-logic/clstaskcontroller";  Name="clsTaskController";  Type="cls"},
    @{SitePath="business-logic/clsstorageresolver"; Name="clsStorageResolver"; Type="cls"},
    @{SitePath="business-logic/clsfieldmigrator";   Name="clsFieldMigrator";   Type="cls"},
    @{SitePath="business-logic/clsformspec";        Name="clsFormSpec";        Type="cls"},
    @{SitePath="business-logic/clscontrolspec";     Name="clsControlSpec";     Type="cls"},
    @{SitePath="business-logic/modformbuilder";     Name="modFormBuilder";     Type="bas"},
    # utility layer
    @{SitePath="utility/modcommon";       Name="modCommon";       Type="bas"},
    @{SitePath="utility/modfileio";       Name="modFileIO";       Type="bas"},
    @{SitePath="utility/moddateutil";     Name="modDateUtil";     Type="bas"},
    @{SitePath="utility/modstringutil";   Name="modStringUtil";   Type="bas"},
    @{SitePath="utility/modimagerender";  Name="modImageRender";  Type="bas"},
    @{SitePath="utility/modformatcolumns";Name="modFormatColumns";Type="bas"},
    # infrastructure layer
    @{SitePath="infrastructure/modsetup"; Name="modSetup"; Type="bas"},
    # logger layer
    @{SitePath="logger/clslogger";  Name="clsLogger";  Type="cls"},
    @{SitePath="logger/clslogentry";Name="clsLogEntry";Type="cls"},
    # ThisWorkbook (production: special, no auto-init)
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
    # match ```vba ... ```  or  ```vbnet ... ```
    $pattern = '(?ms)```(?:vba|vbnet|VBA|VBNET|vb)\b[^\r\n]*\r?\n(.*?)\r?\n```'
    $m = [regex]::Match($Markdown, $pattern)
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

function Sanitize-ForCp932 {
    # Make $code safe to write in Shift_JIS *AND* parseable by VBE Import:
    #  1. Convert U+2028 (line sep) and U+2029 (paragraph sep) to CRLF.
    #  2. Drop U+FFFD (Unicode replacement) - lossy malformed-UTF-8 byte.
    #  3. Per-character: any char that cannot CP932-encode -> "?" (safe
    #     in comments/strings; bad chars are almost always in comments).
    #  4. CRITICAL: normalize ALL line endings to CRLF. VBA .cls files
    #     with LF-only newlines parse incorrectly - the VERSION/BEGIN/END
    #     header gets "swallowed" and Attribute VB_Name lands at line 1,
    #     which yields a "ステートメントの最後" compile error on every
    #     `Public Sub`. Source from curl uses LF; we must rewrite to CRLF.
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
    # CRLF normalization: collapse all CRLF/CR/LF to LF, then expand all
    # LF to CRLF. Idempotent regardless of input newline style.
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
Write-Host "==== Phase 1: Download $($modules.Count) modules ===="
New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null

$basClsFiles = @()
$thisWorkbookCode = $null

foreach ($m in $modules) {
    $url = "$BaseUrl/$($m.SitePath).md"
    Write-Host ("  [{0,-22}] {1}" -f $m.Name, $url)
    try {
        $resp = Invoke-WebRequest -Uri $url -UseBasicParsing
    } catch {
        Write-Warning "    download failed: $_"
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
            # ThisWorkbook is added via CodeModule.AddFromString (not file
            # import), so we only need to sanitize Unicode -> CP932 once
            # before passing it to CodeModule.
            $thisWorkbookCode = Sanitize-ForCp932 -Code $code
            # also save for inspection
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
$excel.AutomationSecurity = 1   # msoAutomationSecurityLow

try {
    $wb = $excel.Workbooks.Open($XlsmFull)

    # -------------------------------------------------------------------
    # Phase 3: Import .bas / .cls (replace existing same-name)
    # -------------------------------------------------------------------
    Write-Host ""
    Write-Host "==== Phase 3: Import $($basClsFiles.Count) modules ===="
    foreach ($f in $basClsFiles) {
        # Remove existing module of same name first
        try {
            $existing = $wb.VBProject.VBComponents.Item($f.Name)
            if ($existing) {
                $wb.VBProject.VBComponents.Remove($existing) | Out-Null
                Write-Host "  - removed existing $($f.Name)"
            }
        } catch {
            # not present, fine
        }
        try {
            $wb.VBProject.VBComponents.Import($f.Path) | Out-Null
            Write-Host "  + imported $($f.Name)"
        } catch {
            Write-Warning "    import failed for $($f.Name): $_"
        }
    }

    # -------------------------------------------------------------------
    # Phase 4: Replace ThisWorkbook code
    # -------------------------------------------------------------------
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

    # -------------------------------------------------------------------
    # Phase 5: Save xlsm + close + reopen
    #   (Reopen is required so Excel re-indexes the just-imported VBA
    #    modules; otherwise Application.Run("SetupSheetsAndButtons")
    #    fails with "macro not available in this workbook".)
    # -------------------------------------------------------------------
    Write-Host ""
    Write-Host "==== Phase 5: Save xlsm + close + reopen ===="
    $wb.Save()
    Write-Host "  Saved: $XlsmFull"
    $wb.Close($true)
    Start-Sleep -Milliseconds 500
    # IMPORTANT: disable events before reopen, otherwise Workbook_Open
    # fires inside Excel COM. If Workbook_Open hits any error path it
    # calls MsgBox(...), which is NOT suppressed by DisplayAlerts=False
    # and blocks PowerShell forever (Excel waiting for user OK on a
    # hidden dialog).
    $excel.EnableEvents = $false
    $wb = $excel.Workbooks.Open($XlsmFull)
    Write-Host "  Reopened with events disabled (Workbook_Open suppressed)."

    # -------------------------------------------------------------------
    # Phase 6: Run SetupSheetsAndButtons (silent=True, no MsgBox)
    # -------------------------------------------------------------------
    Write-Host ""
    Write-Host "==== Phase 6: Run SetupSheetsAndButtons ===="
    try {
        # Qualify the macro name with workbook so Excel finds it for sure
        $macroQualified = "'" + $wb.Name + "'!SetupSheetsAndButtons"
        $excel.Run($macroQualified, $true) | Out-Null
        Write-Host "  Setup completed (silent=True)."
    } catch {
        Write-Warning "  Setup failed (you can run it manually via Alt+F8 -> SetupSheetsAndButtons): $_"
    }

    # -------------------------------------------------------------------
    # Phase 7: Save again + close
    # -------------------------------------------------------------------
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
