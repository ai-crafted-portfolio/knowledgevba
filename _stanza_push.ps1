$ErrorActionPreference = "Continue"
$log = "C:\kvba\push\_stanza_push.log"
"" | Out-File $log -Encoding utf8
function L($m) { Add-Content -Path $log -Value $m -Encoding utf8 }

Set-Location "C:\kvba\push"

L "=== unlock ==="
if (Test-Path "C:\kvba\push\.git\index.lock") {
    Remove-Item "C:\kvba\push\.git\index.lock" -Force -ErrorAction SilentlyContinue
    L "removed .git/index.lock"
} else {
    L "no lock"
}

L ""
L "=== git status before ==="
(& git status --short 2>&1) | Select-Object -First 5 | ForEach-Object { L $_ }
L "(omitted, focusing on stanza scope)"

L ""
L "=== git add (stanza scope only) ==="
$paths = @(
    "docs/stanza/",
    "mkdocs.yml",
    ".gitattributes"
)
foreach ($p in $paths) {
    L "add: $p"
    (& git add $p 2>&1) | ForEach-Object { L $_ }
}

L ""
L "=== git diff --cached --stat ==="
(& git diff --cached --stat 2>&1) | ForEach-Object { L $_ }

L ""
L "=== git diff --cached --name-only ==="
(& git diff --cached --name-only 2>&1) | ForEach-Object { L $_ }

L ""
L "=== git commit ==="
$msg = "add: ui_seed stanza setup files (10 screens) for site"
(& git commit -m $msg 2>&1) | ForEach-Object { L $_ }

L ""
L "=== git log -1 ==="
(& git log --oneline -1 2>&1) | ForEach-Object { L $_ }

L ""
L "=== git push origin main ==="
(& git push origin main 2>&1) | ForEach-Object { L $_ }

L ""
L "=== done ==="
