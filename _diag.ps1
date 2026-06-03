$ErrorActionPreference = "Continue"
$log = "C:\kvba\push\_diag.log"
"" | Out-File $log -Encoding utf8
function L($m) { Add-Content -Path $log -Value $m -Encoding utf8 }

Set-Location "C:\kvba\push"

L "=== current state ==="
L ("HEAD=" + (& git rev-parse HEAD 2>&1))
L ("status:")
(& git status --short 2>&1) | Select-Object -First 100 | ForEach-Object { L $_ }

L ""
L "=== rebase abort ==="
$a = (& git rebase --abort 2>&1) | Out-String
L $a
L "abort exit=$LASTEXITCODE"

L ""
L "=== post-abort state ==="
L ("HEAD=" + (& git rev-parse HEAD 2>&1))
L ("status:")
(& git status --short 2>&1) | Select-Object -First 100 | ForEach-Object { L $_ }

L ""
L "=== remote 0a236b2 file listing under docs/ ==="
(& git ls-tree -r --name-only origin/main -- docs/ 2>&1) | ForEach-Object { L $_ }

L ""
L "=== local 7757b37 file listing under docs/ ==="
(& git ls-tree -r --name-only 7757b37 -- docs/ 2>&1) | ForEach-Object { L $_ }

L ""
L "=== files in remote NOT in local 7757b37 ==="
$rem = (& git ls-tree -r --name-only origin/main -- docs/ 2>&1)
$loc = (& git ls-tree -r --name-only 7757b37 -- docs/ 2>&1)
$diff = Compare-Object $rem $loc -PassThru | Where-Object { $_.SideIndicator -eq '<=' }
$diff | ForEach-Object { L $_ }

L ""
L "=== files in local 7757b37 NOT in remote ==="
$diff2 = Compare-Object $rem $loc -PassThru | Where-Object { $_.SideIndicator -eq '=>' }
$diff2 | ForEach-Object { L $_ }
