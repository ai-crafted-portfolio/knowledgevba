$ErrorActionPreference = "Continue"
$log = "C:\kvba\push\_chk.log"
"" | Out-File $log -Encoding utf8
function L($m) { Add-Content -Path $log -Value $m -Encoding utf8 }

Set-Location "C:\kvba\push"

L "=== git log -3 ==="
(& git log --oneline -3 2>&1) | ForEach-Object { L $_ }

L ""
L "=== git status --short (count) ==="
$st = (& git status --short 2>&1)
L "lines=$($st.Count)"
$st | Select-Object -First 30 | ForEach-Object { L $_ }

L ""
L "=== git diff --cached --stat (head) ==="
$ds = (& git diff --cached --stat 2>&1)
L "diff lines=$($ds.Count)"
$ds | Select-Object -First 5 | ForEach-Object { L $_ }
$ds | Select-Object -Last 3 | ForEach-Object { L $_ }

L ""
L "=== HEAD ==="
$head = (& git rev-parse HEAD 2>&1)
L "HEAD=$head"

L ""
L "=== remote main ==="
$ls = (& git ls-remote origin main 2>&1)
L "$ls"

L ""
L "=== done ==="
