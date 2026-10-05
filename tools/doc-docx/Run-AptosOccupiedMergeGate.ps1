# Run the strict occupied-merge source-appearance gate with the official
# Aptos font extracted outside this repository. Word is the render oracle.
param(
    [Parameter(Mandatory)][string]$FontDirectory,
    [string]$ResultsDirectory = (Join-Path ([IO.Path]::GetPathRoot($PSScriptRoot)) ('docxport-aptos-merge-' + (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
$fontRoot = (Resolve-Path -LiteralPath $FontDirectory).Path
$fontFile = Join-Path $fontRoot 'Aptos.ttf'
if (-not (Test-Path -LiteralPath $fontFile -PathType Leaf)) {
    throw "Aptos.ttf is missing from $fontRoot"
}
$project = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path `
    'DocxportNet.Tests/DocxportNet.Tests.csproj'
New-Item -ItemType Directory -Path $ResultsDirectory -Force | Out-Null
$priorFont = $env:DOCXPORT_FONT_DIRECTORY
$priorRequire = $env:DOCXPORT_REQUIRE_APTOS_FONT
$priorWord = $env:DOCXPORT_VERIFY_WORD_RENDER
try {
    $env:DOCXPORT_FONT_DIRECTORY = $fontRoot
    $env:DOCXPORT_REQUIRE_APTOS_FONT = '1'
    $env:DOCXPORT_VERIFY_WORD_RENDER = '1'
    & dotnet test $project --no-restore `
        --filter 'FullyQualifiedName~AvailableAptosMetricsFitOccupiedMergesInAllStories' `
        --results-directory $ResultsDirectory `
        --logger 'trx;LogFileName=aptos-occupied-merge.trx' `
        --verbosity quiet -p:WarningLevel=0
    if ($LASTEXITCODE -ne 0) { throw 'The strict Aptos occupied-merge gate failed.' }
    $reportPath = Join-Path $ResultsDirectory 'aptos-occupied-merge.trx'
    [xml]$report = Get-Content -LiteralPath $reportPath -Raw
    $counts = $report.TestRun.ResultSummary.Counters
    if ([int]$counts.total -ne 3 -or [int]$counts.passed -ne 3 -or
        [int]$counts.failed -ne 0) {
        throw "Expected 3/3 executed Aptos cases; got $($counts.passed)/$($counts.total)."
    }
    Write-Host "Strict Aptos occupied-merge gate: 3/3 passed. Report: $reportPath"
}
finally {
    $env:DOCXPORT_FONT_DIRECTORY = $priorFont
    $env:DOCXPORT_REQUIRE_APTOS_FONT = $priorRequire
    $env:DOCXPORT_VERIFY_WORD_RENDER = $priorWord
}