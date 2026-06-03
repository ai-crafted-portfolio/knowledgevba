$ErrorActionPreference = "Continue"
$log = "C:\kvba\push\_commit_push.log"
"" | Out-File $log -Encoding utf8
function L($m) { Add-Content -Path $log -Value $m -Encoding utf8 }

Set-Location "C:\kvba\push"

L "=== BEGIN $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
L ""

L "=== Phase 3a: pre-commit HEAD ==="
$preHead = (& git rev-parse HEAD 2>&1)
L "preHead=$preHead"

L ""
L "=== Phase 3b: git config ==="
L ("email=" + (& git config user.email 2>&1))
L ("name=" + (& git config user.name 2>&1))
L ("origin=" + (& git config remote.origin.url 2>&1))

L ""
L "=== Phase 3c: git commit ==="
$msg = "knowledgevba v2.3 site: 3 bat + ps1 fulltext + 68 module pages"
$out = (& git commit -m $msg 2>&1) | Out-String
L $out
L "commit exit=$LASTEXITCODE"

L ""
L "=== Phase 3d: post-commit HEAD ==="
$postHead = (& git rev-parse HEAD 2>&1)
L "postHead=$postHead"
$logone = (& git log -1 --oneline 2>&1)
L "log-1=$logone"

L ""
L "=== Phase 3e: git push origin main ==="
$pushOut = (& git push origin main 2>&1) | Out-String
L $pushOut
L "push exit=$LASTEXITCODE"

L ""
L "=== Phase 3f: post-push remote ls ==="
$lsr = (& git ls-remote origin main 2>&1)
L "ls-remote=$lsr"

L ""
L "=== END $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
