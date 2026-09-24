$ErrorActionPreference = 'Continue'
$root = 'C:\Users\SAM\BigFix-FailureDashboard'
$script = Join-Path $root 'BigFix-FailureDashboard.ps1'
$stamp = Get-Date -Format 'yyyyMMdd'
$reports = Join-Path $root 'reports'
$history = Join-Path $reports 'history'
if (-not (Test-Path -LiteralPath $history)) { New-Item -ItemType Directory -Path $history -Force | Out-Null }
$html = Join-Path $reports ("bf-report-" + $stamp + ".html")
$csv  = Join-Path $history ("bf-report-" + $stamp + ".csv")
$json = Join-Path $history ("bf-report-" + $stamp + ".json")
$log  = Join-Path $reports 'last-run.log'

$cache = Join-Path $root 'cache\last-pull.json'
$sample = Join-Path $root 'sample-data\sample-failures.csv'

$args = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script, '-ExportHtml', $html, '-ExportCsv', $csv, '-ExportJson', $json)
if (Test-Path -LiteralPath $cache) {
    $args += '-UseCache'
} else {
    $args += @('-ImportPath', $sample)
}

$lines = @('run=' + (Get-Date -Format o), 'args=' + ($args -join ' '))
try {
    $out = & powershell.exe @args 2>&1
    $code = $LASTEXITCODE
    $lines += ('exit=' + $code)
    $lines += ($out | ForEach-Object { $_.ToString() })
} catch {
    $lines += ('error=' + $_.Exception.Message)
    $code = 1
}
$lines | Set-Content -LiteralPath $log -Encoding UTF8
exit $code
