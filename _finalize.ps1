$ErrorActionPreference = "Continue"
$log = "C:\kvba\push\_finalize.log"
"" | Out-File $log -Encoding utf8
function L($m) { Add-Content -Path $log -Value $m -Encoding utf8 }

Set-Location "C:\kvba\push"

L "=== BEGIN $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

# Phase A: reset hard to origin/main (drops our 7757b37 commit, but we have backup in _bk\)
L ""
L "=== Phase A: git reset --hard origin/main ==="
$r = (& git reset --hard origin/main 2>&1) | Out-String
L $r
L "reset exit=$LASTEXITCODE"
L ("HEAD=" + (& git rev-parse HEAD 2>&1))

# Phase B: restore our specific files
L ""
L "=== Phase B: restore install.md / usage.md / troubleshooting.md ==="
Copy-Item -Path "_bk\docs\install.md" -Destination "docs\install.md" -Force
Copy-Item -Path "_bk\docs\usage.md" -Destination "docs\usage.md" -Force
Copy-Item -Path "_bk\docs\troubleshooting.md" -Destination "docs\troubleshooting.md" -Force
L "copied install/usage/troubleshooting"

L ""
L "=== Phase C: restore docs\modules\ tree ==="
if (Test-Path "docs\modules") { Remove-Item "docs\modules" -Recurse -Force }
New-Item -ItemType Directory -Path "docs\modules" -Force | Out-Null
Copy-Item -Path "_bk\docs\modules\*" -Destination "docs\modules\" -Recurse -Force
$cnt = (Get-ChildItem "docs\modules" -Recurse -File | Measure-Object).Count
L "modules tree restored, file count=$cnt"

L ""
L "=== Phase D: replace mkdocs.yml with merged version ==="
Copy-Item -Path "_bk\mkdocs_merged.yml" -Destination "mkdocs.yml" -Force
L "mkdocs.yml replaced"

L ""
L "=== Phase E: clean stray writetest tmp ==="
if (Test-Path "docs\_writetest.tmp") { Remove-Item "docs\_writetest.tmp" -Force; L "removed _writetest.tmp" }

L ""
L "=== Phase F: git status ==="
(& git status --short 2>&1) | Select-Object -First 100 | ForEach-Object { L $_ }

L ""
L "=== Phase G: git add ==="
$g = (& git add docs/install.md docs/usage.md docs/troubleshooting.md docs/modules/ mkdocs.yml 2>&1) | Out-String
L $g
L "add exit=$LASTEXITCODE"

L ""
L "=== Phase H: git status (post-add) ==="
(& git status --short 2>&1) | Select-Object -First 100 | ForEach-Object { L $_ }

L ""
L "=== Phase I: git diff --cached --stat ==="
$ds = (& git diff --cached --stat 2>&1)
L ("staged file count=" + ($ds.Count))
$ds | Select-Object -Last 3 | ForEach-Object { L $_ }

L ""
L "=== Phase J: git commit ==="
$msg = "knowledgevba v2.3 site: 3 bat + ps1 fulltext + 68 module pages"
$c = (& git commit -m $msg 2>&1) | Out-String
L $c
L "commit exit=$LASTEXITCODE"
L ("HEAD=" + (& git rev-parse HEAD 2>&1))

L ""
L "=== Phase K: git push origin main ==="
$p = (& git push origin main 2>&1) | Out-String
L $p
L "push exit=$LASTEXITCODE"

L ""
L "=== Phase L: post-push ls-remote ==="
L ("ls-remote=" + (& git ls-remote origin main 2>&1))

L ""
L "=== END $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
