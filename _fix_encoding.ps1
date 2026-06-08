$ErrorActionPreference = "Continue"
$log = "C:\kvba\push\_fix_encoding.log"
"" | Out-File $log -Encoding utf8
function L($m) { Add-Content -Path $log -Value $m -Encoding utf8 }

Set-Location "C:\kvba\push"

L "=== BEGIN $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

L ""
L "=== Phase A: pre HEAD ==="
$preHead = (& git rev-parse HEAD 2>&1)
L "preHead=$preHead"

L ""
L "=== Phase B: git status (untracked + modified) ==="
$st = (& git status --porcelain 2>&1) | Out-String
L $st

L ""
L "=== Phase C: git add docs/modules + tools ==="
& git add docs/modules/ 2>&1 | Out-String | ForEach-Object { L $_ }
& git add tools/regen_modules_pages.py 2>&1 | Out-String | ForEach-Object { L $_ }

L ""
L "=== Phase D: git status after add ==="
$st2 = (& git status --porcelain 2>&1) | Out-String
L $st2

L ""
L "=== Phase E: git commit ==="
$msg = "fix(modules): regenerate VBA code block pages from CP932 source (mojibake fix)"
$out = (& git commit -m $msg 2>&1) | Out-String
L $out
L "commit exit=$LASTEXITCODE"

L ""
L "=== Phase F: post HEAD ==="
$postHead = (& git rev-parse HEAD 2>&1)
L "postHead=$postHead"
$logone = (& git log -1 --oneline 2>&1)
L "log-1=$logone"

L ""
L "=== Phase G: git push origin main ==="
$pushOut = (& git push origin main 2>&1) | Out-String
L $pushOut
L "push exit=$LASTEXITCODE"

L ""
L "=== Phase H: post-push ls-remote ==="
$lsr = (& git ls-remote origin main 2>&1)
L "ls-remote=$lsr"

L ""
L "=== END $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
