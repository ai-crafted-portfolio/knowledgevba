$ErrorActionPreference = "Continue"
$log = "C:\kvba\push\_show_remote.log"
"" | Out-File $log -Encoding utf8
function L($m) { Add-Content -Path $log -Value $m -Encoding utf8 }

Set-Location "C:\kvba\push"

L "=== remote mkdocs.yml ==="
(& git show origin/main:mkdocs.yml 2>&1) | ForEach-Object { L $_ }

L ""
L "=== remote docs/index.md (first 50) ==="
$idx = (& git show origin/main:docs/index.md 2>&1)
$idx | Select-Object -First 50 | ForEach-Object { L $_ }

L ""
L "=== local index.md (first 50) ==="
Get-Content "docs\index.md" -ErrorAction SilentlyContinue | Select-Object -First 50 | ForEach-Object { L $_ }
