# Diagnostic for the three strict source-appearance cases with Word as oracle.
# These are accepted first-milestone exceptions under Run-NoWrapExceptionGate.ps1.
# This diagnostic still reports failures until the DOC auto-fit layout is fixed.
param(
    [string]$ResultsDirectory = (Join-Path ([IO.Path]::GetPathRoot($PSScriptRoot)) ('docxport-strict-nowrap-' + (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
$project = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path `
    'DocxportNet.Tests/DocxportNet.Tests.csproj'
New-Item -ItemType Directory -Path $ResultsDirectory -Force | Out-Null
$prior = $env:DOCXPORT_VERIFY_STRICT_NOWRAP
try {
    $env:DOCXPORT_VERIFY_STRICT_NOWRAP = '1'
    & dotnet test $project --no-restore `
        --filter 'FullyQualifiedName~StrictAutoWidthNoWrapSourceAppearanceGate' `
        --results-directory $ResultsDirectory `
        --logger 'trx;LogFileName=strict-nowrap.trx' `
        --verbosity quiet -p:WarningLevel=0
    $exitCode = $LASTEXITCODE
    $reportPath = Join-Path $ResultsDirectory 'strict-nowrap.trx'
    [xml]$report = Get-Content -LiteralPath $reportPath -Raw
    $counts = $report.TestRun.ResultSummary.Counters
    if ([int]$counts.total -ne 3) {
        throw "Expected all three strict no-wrap cases; got $($counts.total). Report: $reportPath"
    }
    Write-Host "Strict no-wrap Word gate: $($counts.passed)/3 passed, $($counts.failed) failed. Report: $reportPath"
    if ($exitCode -ne 0 -or [int]$counts.passed -ne 3) {
        throw 'The strict 0.04 no-wrap source-appearance gate is still open.'
    }
}
finally {
    $env:DOCXPORT_VERIFY_STRICT_NOWRAP = $prior
}
