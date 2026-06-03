$ErrorActionPreference = "Continue"
$log = "C:\kvba\push\_merge_mkdocs.log"
"" | Out-File $log -Encoding utf8
function L($m) { Add-Content -Path $log -Value $m -Encoding utf8 }

Set-Location "C:\kvba\push"

L "=== BEGIN $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

# Step 1: backup files BEFORE we touch git
$tmp = "C:\kvba\push\_bk"
if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
New-Item -ItemType Directory -Path "$tmp\docs" -Force | Out-Null

Copy-Item -Path "mkdocs.yml" -Destination "$tmp\mkdocs_local.yml" -Force
Copy-Item -Path "docs\install.md" -Destination "$tmp\docs\install.md" -Force
Copy-Item -Path "docs\usage.md" -Destination "$tmp\docs\usage.md" -Force
Copy-Item -Path "docs\troubleshooting.md" -Destination "$tmp\docs\troubleshooting.md" -Force
Copy-Item -Path "docs\modules" -Destination "$tmp\docs\modules" -Recurse -Force
L "backup done"

# Step 2: extract remote mkdocs.yml (raw bytes, no encoding conversion)
& git show origin/main:mkdocs.yml | Out-File "$tmp\mkdocs_remote.yml" -Encoding utf8
L "remote mkdocs.yml saved to $tmp\mkdocs_remote.yml"

# Step 3: extract modules nav block from local mkdocs.yml
# Read local file as bytes to preserve encoding (it should be UTF-8 since I wrote it via Write tool)
$localBytes = [System.IO.File]::ReadAllBytes("$tmp\mkdocs_local.yml")
L ("local mkdocs bytes=" + $localBytes.Length)
$remoteBytes = [System.IO.File]::ReadAllBytes("$tmp\mkdocs_remote.yml")
L ("remote mkdocs bytes=" + $remoteBytes.Length)

# detect BOM / encoding
$localStr = [System.IO.File]::ReadAllText("$tmp\mkdocs_local.yml", [System.Text.Encoding]::UTF8)
$remoteStr = [System.IO.File]::ReadAllText("$tmp\mkdocs_remote.yml", [System.Text.Encoding]::UTF8)
L ("local string length=" + $localStr.Length)
L ("remote string length=" + $remoteStr.Length)

L ""
L "=== first 20 lines of remote ==="
($remoteStr -split "`n") | Select-Object -First 20 | ForEach-Object { L $_ }

L ""
L "=== nav section of local (search for 'モジュール一覧') ==="
$lines = $localStr -split "`n"
$start = -1; $end = -1
for ($i = 0; $i -lt $lines.Length; $i++) {
    if ($lines[$i] -match 'モジュール一覧') { $start = $i; break }
}
L "modules nav start line=$start"
if ($start -ge 0) {
    # find end (next top-level nav item at same indent, or EOF)
    for ($j = $start + 1; $j -lt $lines.Length; $j++) {
        $line = $lines[$j]
        if ($line -match '^\s*-\s*\S+:' -and $line -notmatch '^      ' -and $line -notmatch '^          ') {
            # this is at top-level nav (2 spaces indent: "  - X: Y")
            if ($line.StartsWith('  - ')) { $end = $j - 1; break }
        }
    }
    if ($end -lt 0) { $end = $lines.Length - 1 }
    L "modules nav end line=$end"
    L "--- nav block ---"
    for ($k = $start; $k -le $end; $k++) { L $lines[$k] }
}

L ""
L "=== END $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
