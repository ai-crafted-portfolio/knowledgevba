$ErrorActionPreference = "Continue"
$log = "C:\kvba\push\_pull_push.log"
"" | Out-File $log -Encoding utf8
function L($m) { Add-Content -Path $log -Value $m -Encoding utf8 }

Set-Location "C:\kvba\push"

L "=== BEGIN $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
L ""

L "=== pre HEAD ==="
L ("local=" + (& git rev-parse HEAD 2>&1))

L ""
L "=== git fetch ==="
$f = (& git fetch origin 2>&1) | Out-String
L $f
L "fetch exit=$LASTEXITCODE"

L ""
L "=== branch state ==="
L ("local=" + (& git rev-parse HEAD 2>&1))
L ("origin/main=" + (& git rev-parse origin/main 2>&1))
L ""
L "=== log local vs origin (last 5) ==="
(& git log --oneline -5 HEAD 2>&1) | ForEach-Object { L ("local: " + $_) }
(& git log --oneline -5 origin/main 2>&1) | ForEach-Object { L ("origin: " + $_) }

L ""
L "=== git pull --rebase ==="
$p = (& git pull --rebase origin main 2>&1) | Out-String
L $p
L "pull exit=$LASTEXITCODE"

L ""
L "=== post-pull HEAD ==="
L ("local=" + (& git rev-parse HEAD 2>&1))
(& git log --oneline -5 2>&1) | ForEach-Object { L $_ }

L ""
L "=== git push origin main ==="
$pu = (& git push origin main 2>&1) | Out-String
L $pu
L "push exit=$LASTEXITCODE"

L ""
L "=== post-push remote ls ==="
L ("ls-remote=" + (& git ls-remote origin main 2>&1))

L ""
L "=== END $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
