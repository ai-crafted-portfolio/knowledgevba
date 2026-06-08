$ErrorActionPreference = "Continue"
$logPath = "C:\kvba\push\_session_push.log"
Start-Transcript -Path $logPath -Force | Out-Null

Write-Host "=== Phase 1: Excel zombie check ==="
$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excel) {
    Write-Host "EXCEL processes found: $($excel.Count) - terminating"
    $excel | ForEach-Object { try { Stop-Process -Id $_.Id -Force -ErrorAction Stop } catch {} }
    Start-Sleep -Seconds 2
} else {
    Write-Host "EXCEL: 0 process"
}

Write-Host ""
Write-Host "=== Phase 1: cd + remove index.lock ==="
Set-Location "C:\kvba\push"
$lock = "C:\kvba\push\.git\index.lock"
if (Test-Path $lock) {
    Remove-Item $lock -Force
    Write-Host "Removed: $lock"
} else {
    Write-Host "No index.lock"
}

Write-Host ""
Write-Host "=== Phase 2: git add ==="
& git add docs/install.md docs/modules/ mkdocs.yml docs/troubleshooting.md docs/usage.md docs/customize.md docs/operations.md 2>&1
$addExit = $LASTEXITCODE
Write-Host "git add exit: $addExit"

Write-Host ""
Write-Host "=== Phase 2: git status ==="
& git status --short 2>&1
Write-Host "git status exit: $LASTEXITCODE"

Write-Host ""
Write-Host "=== Phase 2: git diff --cached --stat ==="
& git diff --cached --stat 2>&1
Write-Host "diff stat exit: $LASTEXITCODE"

Write-Host ""
Write-Host "=== Phase 3: git config check ==="
& git config user.email 2>&1
& git config user.name 2>&1
& git config remote.origin.url 2>&1

Write-Host ""
Write-Host "=== Phase 3: git commit ==="
$msg = "knowledgevba v2.3 site: 3 bat + ps1 fulltext + 68 module pages"
& git commit -m $msg 2>&1
$commitExit = $LASTEXITCODE
Write-Host "git commit exit: $commitExit"

Write-Host ""
Write-Host "=== Phase 3: latest commit hash ==="
& git log -1 --oneline 2>&1
$commitHash = & git rev-parse HEAD 2>&1
Write-Host "HEAD: $commitHash"

Write-Host ""
Write-Host "=== Phase 3: git push origin main ==="
& git push origin main 2>&1
$pushExit = $LASTEXITCODE
Write-Host "git push exit: $pushExit"

Write-Host ""
Write-Host "=== Final state ==="
Write-Host "addExit=$addExit commitExit=$commitExit pushExit=$pushExit"
Write-Host "COMMIT_HASH=$commitHash"

Stop-Transcript | Out-Null
