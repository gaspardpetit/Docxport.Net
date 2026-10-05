# Measure the three strict no-wrap cases in Word without a PDF render sweep.
# Requires a built DocxportNet.Tests net10.0 output and desktop Word.
param(
    [string]$ResultsDirectory = (Join-Path ([IO.Path]::GetPathRoot($PSScriptRoot)) ('docxport-nowrap-geometry-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))),
    [switch]$CharacterLayout,
    [switch]$Mode11Control
)

$ErrorActionPreference = 'Stop'
$repository = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$fixtures = Join-Path $repository 'DocxportNet.Tests/Fixtures'
$project = Join-Path $repository 'DocxportNet.Tests/DocxportNet.Tests.csproj'
$binary = Join-Path $repository 'DocxportNet.Tests/bin/Debug/net10.0'
& dotnet build $project --no-restore --verbosity quiet -p:WarningLevel=0
if ($LASTEXITCODE -ne 0) { throw 'The test project build failed.' }
Add-Type -Path (Join-Path $binary 'Microsoft.Extensions.Logging.Abstractions.dll')
Add-Type -Path (Join-Path $binary 'DocxportNet.dll')
New-Item -ItemType Directory -Path $ResultsDirectory -Force | Out-Null
$names = @(
    'WordVisibleNoWrapOffAllStories'
    'WordVisibleNoWrapAllStories'
    'WordVisibleNoWrapArialAllStories'
)
# Keep the source table markup and change only the DOCX compatibility mode.
function New-Mode11Control([string]$SourcePath, [string]$TargetPath) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.File]::Copy($SourcePath, $TargetPath, $true)
    $archive = [IO.Compression.ZipFile]::Open($TargetPath,
        [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $archive.GetEntry('word/settings.xml')
        if ($null -eq $entry) { throw 'Source DOCX has no settings part.' }
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $pattern = '(<w:compatSetting\b[^>]*w:name="compatibilityMode"[^>]*w:val=")15("[^>]*/>)'
        $updated = [regex]::Replace($xml, $pattern, '${1}11${2}')
        if ($updated -eq $xml) { throw 'Expected mode-15 setting was not found.' }
        $entry.Delete()
        $replacement = $archive.CreateEntry('word/settings.xml',
            [IO.Compression.CompressionLevel]::Optimal)
        $writer = [IO.StreamWriter]::new($replacement.Open(),
            [Text.UTF8Encoding]::new($false))
        try { $writer.Write($updated) } finally { $writer.Dispose() }
    } finally { $archive.Dispose() }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$rows = @()
$characters = @()
try {
    foreach ($name in $names) {
        $folder = if ($name -like '*Arial*') { 'DocKnownGaps' } else { 'Doc' }
        $source = Join-Path (Join-Path $fixtures $folder) ($name + '.docx')
        $generated = Join-Path $ResultsDirectory ($name + '-generated.doc')
        [IO.File]::WriteAllBytes($generated,
            [DocxportNet.DxpDocExport]::Export([IO.File]::ReadAllBytes($source)))
        $native = [IO.Path]::ChangeExtension($source, '.doc')
        $routes = @('source', 'native', 'generated')
        if ($Mode11Control) {
            $control = Join-Path $ResultsDirectory ($name + '-mode11.docx')
            New-Mode11Control $source $control
            $routes += 'mode11'
        }
        foreach ($route in $routes) {
            $path = if ($route -eq 'source') { $source }
                elseif ($route -eq 'native') { $native }
                elseif ($route -eq 'mode11') { $control }
                else { $generated }
            $document = $word.Documents.Open($path, $false, $true)
            try {
                foreach ($story in @('body', 'header', 'footer')) {
                    $table = if ($story -eq 'body') {
                        $document.Tables.Item(1)
                    } elseif ($story -eq 'header') {
                        $document.Sections.Item(1).Headers.Item(1).Range.Tables.Item(1)
                    } else {
                        $document.Sections.Item(1).Footers.Item(1).Range.Tables.Item(1)
                    }
                    $firstCell = $table.Cell(1, 1)
                    $secondCell = $table.Cell(1, 2)
                    $first = [double]$firstCell.Width
                    $second = [double]$secondCell.Width
                    # WdInformation.wdHorizontalPositionRelativeToPage = 5.
                    $firstTextX = [double]$firstCell.Range.Characters.Item(1).Information(5)
                    $secondTextX = [double]$secondCell.Range.Characters.Item(1).Information(5)
                    # The cell range ends with a paragraph/cell marker; measure
                    # the final text character's vertical position.
                    $firstText = $firstCell.Range.Text.TrimEnd([char]13, [char]7)
                    $firstLast = $firstCell.Range.Characters.Item($firstText.Length)
                    $firstLastY = [double]$firstLast.Information(6)
                    $page = $document.Sections.Item(1).PageSetup
                    $paragraph = $firstCell.Range.ParagraphFormat
                    $firstFont = $firstCell.Range.Characters.Item(1).Font
                    if ($CharacterLayout) {
                        for ($characterIndex = 1; $characterIndex -le $firstText.Length;
                            $characterIndex++) {
                            $character = $firstCell.Range.Characters.Item($characterIndex)
                            $characters += [pscustomobject]@{
                                Fixture = $name; Route = $route; Story = $story
                                CharacterIndex = $characterIndex
                                Character = [string]$firstText[$characterIndex - 1]
                                XPt = [Math]::Round([double]$character.Information(5), 3)
                                YPt = [Math]::Round([double]$character.Information(6), 3)
                            }
                        }
                    }
                    $rows += [pscustomobject]@{
                        Fixture = $name; Route = $route; Story = $story
                        CompatibilityMode = [int]$document.CompatibilityMode
                        TextContainerPt = [Math]::Round(
                            [double]$page.PageWidth - [double]$page.LeftMargin -
                            [double]$page.RightMargin - [double]$page.Gutter, 3)
                        ParagraphLeftPt = [Math]::Round([double]$paragraph.LeftIndent, 3)
                        ParagraphRightPt = [Math]::Round([double]$paragraph.RightIndent, 3)
                        ParagraphAfterPt = [Math]::Round([double]$paragraph.SpaceAfter, 3)
                        ParagraphLinePt = [Math]::Round([double]$paragraph.LineSpacing, 3)
                        FirstFontName = [string]$firstFont.Name
                        FirstFontSizePt = [double]$firstFont.Size
                        FirstFontKerningPt = [double]$firstFont.Kerning
                        FirstFontScalePercent = [double]$firstFont.Scaling
                        FirstCellPt = [Math]::Round($first, 3)
                        SecondCellPt = [Math]::Round($second, 3)
                        TablePt = [Math]::Round($first + $second, 3)
                        RowIndentPt = [Math]::Round(
                            [double]$table.Rows.Item(1).LeftIndent, 3)
                        FirstTextXPt = [Math]::Round($firstTextX, 3)
                        SecondTextXPt = [Math]::Round($secondTextX, 3)
                        FirstTextLastYPt = [Math]::Round($firstLastY, 3)
                        RowAlignment = [int]$table.Rows.Item(1).Alignment
                        Direction = [int]$table.TableDirection
                        AutoFit = [bool]$table.AllowAutoFit
                        FirstCellWrap = [bool]$table.Cell(1, 1).WordWrap
                        LeftPaddingPt = [Math]::Round([double]$table.LeftPadding, 3)
                        RightPaddingPt = [Math]::Round([double]$table.RightPadding, 3)
                    }
                }
            } finally { $document.Close($false) }
        }
    }
} finally { $word.Quit() }
$path = Join-Path $ResultsDirectory 'word-geometry.csv'
$rows | Export-Csv -LiteralPath $path -NoTypeInformation
if ($CharacterLayout) {
    $characterPath = Join-Path $ResultsDirectory 'word-character-geometry.csv'
    $characters | Export-Csv -LiteralPath $characterPath -NoTypeInformation
    $lineSummary = foreach ($group in ($characters | Group-Object Fixture, Route, Story)) {
        $positions = @($group.Group)
        $firstLineY = $positions[0].YPt
        [pscustomobject]@{
            Fixture = $positions[0].Fixture; Route = $positions[0].Route
            Story = $positions[0].Story
            FirstLineCharacters = @($positions | Where-Object YPt -eq $firstLineY).Count
            LineCount = @($positions | Select-Object -ExpandProperty YPt -Unique).Count
        }
    }
    $linePath = Join-Path $ResultsDirectory 'word-line-summary.csv'
    $lineSummary | Export-Csv -LiteralPath $linePath -NoTypeInformation
    Write-Host "Character positions: $characterPath"
    Write-Host "Line summary: $linePath"
}
$sourceByCase = @{}
foreach ($row in $rows | Where-Object Route -eq 'source') {
    $sourceByCase[$row.Fixture + '/' + $row.Story] = $row
}
$deltas = foreach ($row in $rows | Where-Object Route -ne 'source') {
    $source = $sourceByCase[$row.Fixture + '/' + $row.Story]
    [pscustomobject]@{
        Fixture = $row.Fixture; Route = $row.Route; Story = $row.Story
        TableDeltaPt = [Math]::Round($row.TablePt - $source.TablePt, 3)
        FirstCellDeltaPt = [Math]::Round($row.FirstCellPt - $source.FirstCellPt, 3)
        SecondCellDeltaPt = [Math]::Round($row.SecondCellPt - $source.SecondCellPt, 3)
        IndentDeltaPt = [Math]::Round($row.RowIndentPt - $source.RowIndentPt, 3)
        FirstTextXDeltaPt = [Math]::Round($row.FirstTextXPt - $source.FirstTextXPt, 3)
        SecondTextXDeltaPt = [Math]::Round($row.SecondTextXPt - $source.SecondTextXPt, 3)
        FirstTextLastYDeltaPt = [Math]::Round(
            $row.FirstTextLastYPt - $source.FirstTextLastYPt, 3)
        WrapMatches = $row.FirstCellWrap -eq $source.FirstCellWrap
        PaddingMatches = $row.LeftPaddingPt -eq $source.LeftPaddingPt -and
            $row.RightPaddingPt -eq $source.RightPaddingPt
    }
}
$deltaPath = Join-Path $ResultsDirectory 'word-geometry-deltas.csv'
$deltas | Export-Csv -LiteralPath $deltaPath -NoTypeInformation
$deltas | Format-Table -AutoSize
Write-Host "Word geometry: $path"
Write-Host "Source-relative deltas: $deltaPath"
