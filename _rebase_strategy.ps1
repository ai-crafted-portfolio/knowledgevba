$ErrorActionPreference = "Continue"
$log = "C:\kvba\push\_rebase_strategy.log"
"" | Out-File $log -Encoding utf8
function L($m) { Add-Content -Path $log -Value $m -Encoding utf8 }

Set-Location "C:\kvba\push"

L "=== BEGIN $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

L ""
L "=== Step 1: backup local 7757b37 mkdocs.yml and install.md to temp ==="
$tmp = "C:\kvba\push\_backup_for_remote_rebase"
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
New-Item -ItemType Directory -Path "$tmp\docs" -Force | Out-Null
New-Item -ItemType Directory -Path "$tmp\docs\modules" -Force | Out-Null

# copy files from current working tree (which is at 7757b37)
Copy-Item -Path "mkdocs.yml" -Destination "$tmp\mkdocs.yml" -Force
Copy-Item -Path "docs\install.md" -Destination "$tmp\docs\install.md" -Force
Copy-Item -Path "docs\usage.md" -Destination "$tmp\docs\usage.md" -Force
Copy-Item -Path "docs\troubleshooting.md" -Destination "$tmp\docs\troubleshooting.md" -Force
Copy-Item -Path "docs\modules" -Destination "$tmp\docs\modules" -Recurse -Force
L "backup done. files:"
(Get-ChildItem $tmp -Recurse -File | Select-Object -ExpandProperty FullName) | ForEach-Object { L $_ }

L ""
L "=== Step 2: git reset --hard origin/main ==="
$r = (& git reset --hard origin/main 2>&1) | Out-String
L $r
L "reset exit=$LASTEXITCODE"

L ""
L "=== Step 3: confirm HEAD ==="
L ("HEAD=" + (& git rev-parse HEAD 2>&1))

L ""
L "=== Step 4: restore files from backup ==="
# install.md (replace remote-deleted version with our new one)
Copy-Item -Path "$tmp\docs\install.md" -Destination "docs\install.md" -Force
Copy-Item -Path "$tmp\docs\usage.md" -Destination "docs\usage.md" -Force
Copy-Item -Path "$tmp\docs\troubleshooting.md" -Destination "docs\troubleshooting.md" -Force
# modules/ entire dir
New-Item -ItemType Directory -Path "docs\modules" -Force | Out-Null
Copy-Item -Path "$tmp\docs\modules\*" -Destination "docs\modules\" -Recurse -Force
L "restored: install.md, usage.md, troubleshooting.md, docs\modules\"

L ""
L "=== Step 5: mkdocs.yml strategy ==="
# Use our local version (7757b37 modified version) since it has the new modules nav
# but we need to verify it doesn't reference now-deleted files (customize.md, operations.md, etc)
Copy-Item -Path "$tmp\mkdocs.yml" -Destination "mkdocs.yml" -Force
L "mkdocs.yml: using our local 7757b37 version"

L ""
L "=== Step 6: git status ==="
(& git status --short 2>&1) | Select-Object -First 80 | ForEach-Object { L $_ }

L ""
L "=== END $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
