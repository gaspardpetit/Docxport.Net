# Run after building DocxportNet.Tests. Example:
# pwsh -File tools/doc-docx/Run-WordRenderSweep.ps1 -Workers 4 -ResultsDirectory D:\docxport-render-results
param(
    [ValidateRange(1, 16)][int]$Workers = 4,
    [string[]]$FixtureNames,
    [string]$ResultsDirectory = (Join-Path ([IO.Path]::GetPathRoot($PSScriptRoot)) ('docxport-word-render-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))),
    [string]$ScratchRoot
)

$ErrorActionPreference = 'Stop'
$repository = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$source = Join-Path $repository 'DocxportNet.Tests/DocStructureWalkerTests.cs'
$project = Join-Path $repository 'DocxportNet.Tests/DocxportNet.Tests.csproj'
$text = [IO.File]::ReadAllText($source)
$method = $text.IndexOf('public void WordAuthoredStoriesRenderCloseToReference(')
if ($method -lt 0) { throw 'Word render theory was not found.' }
$theory = $text.LastIndexOf('[Theory]', $method, [StringComparison]::Ordinal)
if ($theory -lt 0) { throw 'Word render theory data was not found.' }
$all = @([regex]::Matches($text.Substring($theory, $method - $theory),
    '\[InlineData\("([^"]+\.docx)"') | ForEach-Object { $_.Groups[1].Value })
if ($all.Count -lt 200 -or @($all | Sort-Object -Unique).Count -ne $all.Count) {
    throw "Unexpected or duplicate render fixture list ($($all.Count) cases)."
}
# xUnit abbreviates long theory arguments in DisplayName. A unique short prefix
# survives that abbreviation, while the full .docx name may never be searchable.
$tokens = @($all | ForEach-Object { $_.Substring(0, [Math]::Min(32, $_.Length)) })
if (@($tokens | Sort-Object -Unique).Count -ne $all.Count) {
    throw 'Render fixture prefixes are no longer unique; adjust filter tokens.'
}
if ($FixtureNames.Count -gt 0) {
    foreach ($name in $FixtureNames) {
        if ($name -notin $all) { throw "Unknown render fixture: $name" }
    }
    $selected = @($all | Where-Object { $_ -in $FixtureNames })
} else {
    $selected = $all
}
$Workers = [Math]::Min($Workers, $selected.Count)
New-Item -ItemType Directory -Path $ResultsDirectory -Force | Out-Null
if (-not $ScratchRoot) { $ScratchRoot = Join-Path $ResultsDirectory "scratch" }
New-Item -ItemType Directory -Path $ScratchRoot -Force | Out-Null
$jobs = @()
for ($i = 0; $i -lt $Workers; $i++) {
    $shard = @($selected | Where-Object { [array]::IndexOf($selected, $_) % $Workers -eq $i })
    $filter = 'FullyQualifiedName~WordAuthoredStoriesRenderCloseToReference&(' +
        (($shard | ForEach-Object { 'DisplayName~' + $_.Substring(0, [Math]::Min(32, $_.Length)) }) -join '|') + ')'
    $log = Join-Path $ResultsDirectory "worker-$i.log"
    $trx = "worker-$i.trx"
    $workerTemp = Join-Path $ScratchRoot "worker-$i"
    New-Item -ItemType Directory -Path $workerTemp -Force | Out-Null
    $jobs += Start-Job -ArgumentList $project, $filter, $log, $trx, $ResultsDirectory, $workerTemp -ScriptBlock {
        param($project, $filter, $log, $trx, $directory, $workerTemp)
        $env:TEMP = $workerTemp
        $env:TMP = $workerTemp
        $env:DOCXPORT_VERIFY_WORD_RENDER = '1'
        & dotnet test $project --no-build --no-restore --filter $filter `
            --results-directory $directory --logger "trx;LogFileName=$trx" `
            --verbosity quiet > $log 2>&1
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Log = $log; Trx = (Join-Path $directory $trx) }
    }
    Write-Host "Worker $i : $($shard.Count) cases"
}
$results = @($jobs | Wait-Job | Receive-Job)
$passed = 0
$failed = $false
foreach ($result in $results) {
    if ($result.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $result.Trx)) {
        $failed = $true
        Write-Host "FAILED: $($result.Log)"
        Get-Content -LiteralPath $result.Log -Tail 20
        continue
    }
    [xml]$report = Get-Content -LiteralPath $result.Trx -Raw
    $counters = $report.TestRun.ResultSummary.Counters
    $passed += [int]$counters.passed
    if ([int]$counters.failed -ne 0 -or [int]$counters.total -ne [int]$counters.passed) {
        $failed = $true
        Write-Host "Unexpected TRX counts: $($result.Trx)"
    }
}
$jobs | Remove-Job
Write-Host "Word render sweep: $passed/$($selected.Count) passed. Logs: $ResultsDirectory"
if ($failed -or $passed -ne $selected.Count) { exit 1 }
