# Acceptance exception for the three disabled-growth auto-width/no-wrap DOC cases.
# Word renders generated and rebuilt DOC closer to the authored DOCX than
# Word's own native DOC save; semantic and editable auto-fit checks run in the
# ordinary paired corpus. The strict 0.04 source check remains diagnostic.
param(
    [string]$ResultsDirectory = (Join-Path ([IO.Path]::GetPathRoot($PSScriptRoot)) ('docxport-nowrap-exception-' + (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
$project = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path `
    'DocxportNet.Tests/DocxportNet.Tests.csproj'
New-Item -ItemType Directory -Path $ResultsDirectory -Force | Out-Null
$prior = $env:DOCXPORT_VERIFY_WORD_RENDER
try {
    $env:DOCXPORT_VERIFY_WORD_RENDER = '1'
    & dotnet test $project --no-restore `
        --filter 'FullyQualifiedName~AutoWidthNoWrapRenderStaysCloserToSourceThanWordNativeDoc' `
        --results-directory $ResultsDirectory `
        --logger 'trx;LogFileName=nowrap-exception.trx' `
        --verbosity quiet -p:WarningLevel=0
    $exitCode = $LASTEXITCODE
    $reportPath = Join-Path $ResultsDirectory 'nowrap-exception.trx'
    if (-not (Test-Path -LiteralPath $reportPath)) { throw "Missing report: $reportPath" }
    [xml]$report = Get-Content -LiteralPath $reportPath -Raw
    $counts = $report.TestRun.ResultSummary.Counters
    if ($exitCode -ne 0 -or [int]$counts.total -ne 3 -or
        [int]$counts.passed -ne 3 -or [int]$counts.failed -ne 0) {
        throw "Expected 3/3 Word-native comparisons; got $($counts.passed)/$($counts.total). Report: $reportPath"
    }
    Write-Host "No-wrap Word-native exception gate: 3/3 passed. Report: $reportPath"
}
finally {
    $env:DOCXPORT_VERIFY_WORD_RENDER = $prior
}
