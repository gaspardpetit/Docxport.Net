# Compare decoded visible image pixels across the paired DOC corpus.
# ImageMagick is the decoder; Word remains the final page-appearance oracle.
param(
    [string]$ResultsDirectory = (Join-Path ([IO.Path]::GetPathRoot($PSScriptRoot)) ('docxport-image-pixels-' + (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Command magick -ErrorAction SilentlyContinue)) {
    throw 'ImageMagick (magick) is required.'
}
$repository = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$project = Join-Path $repository 'DocxportNet.Tests/DocxportNet.Tests.csproj'
$fixtures = Join-Path $repository 'DocxportNet.Tests/Fixtures/Doc'
$binary = Join-Path $repository 'DocxportNet.Tests/bin/Debug/net10.0'
& dotnet build $project --no-restore --verbosity quiet -p:WarningLevel=0
if ($LASTEXITCODE -ne 0) { throw 'The test project build failed.' }
Add-Type -AssemblyName System.IO.Compression
Add-Type -Path (Join-Path $binary 'Microsoft.Extensions.Logging.Abstractions.dll')
Add-Type -Path (Join-Path $binary 'DocxportNet.dll')
New-Item -ItemType Directory -Path $ResultsDirectory -Force | Out-Null
$mediaDirectory = Join-Path $ResultsDirectory 'media'
New-Item -ItemType Directory -Path $mediaDirectory -Force | Out-Null

function Get-PixelHashes([byte[]]$bytes, [string]$prefix) {
    $zip = [IO.Compression.ZipArchive]::new([IO.MemoryStream]::new($bytes),
        [IO.Compression.ZipArchiveMode]::Read)
    try {
        $hashes = @()
        $entries = @($zip.Entries | Where-Object {
            $_.FullName -match '(^|/)media/.*\.(bmp|png|jpe?g|tiff?|gif)$'
        })
        for ($i = 0; $i -lt $entries.Count; $i++) {
            $entry = $entries[$i]
            $path = Join-Path $mediaDirectory ($prefix + '-' + $i +
                [IO.Path]::GetExtension($entry.FullName))
            $source = $entry.Open()
            $destination = [IO.File]::Create($path)
            try { $source.CopyTo($destination) }
            finally { $destination.Dispose(); $source.Dispose() }
            $rgb = Join-Path $mediaDirectory ($prefix + '-' + $i + '.rgb')
            & magick $path -background white -alpha remove -alpha off `
                -colorspace sRGB -depth 8 ('rgb:' + $rgb)
            if ($LASTEXITCODE -ne 0) { throw "Image decode failed: $path" }
            $size = & magick identify -format '%wx%h' $path
            if ($LASTEXITCODE -ne 0) { throw "Image dimensions failed: $path" }
            $hashes += $size + ':' + [Convert]::ToHexString(
                [Security.Cryptography.SHA256]::HashData(
                    [IO.File]::ReadAllBytes($rgb)))
        }
        return @($hashes | Sort-Object -Unique)
    }
    finally { $zip.Dispose() }
}

$rows = @()
$checked = 0
foreach ($fixture in Get-ChildItem -LiteralPath $fixtures -Filter '*.docx') {
    $nativePath = [IO.Path]::ChangeExtension($fixture.FullName, '.doc')
    if (-not (Test-Path -LiteralPath $nativePath)) { continue }
    [byte[]]$source = [IO.File]::ReadAllBytes($fixture.FullName)
    $expected = @(Get-PixelHashes $source ($fixture.BaseName + '-source'))
    if ($expected.Count -eq 0) { continue }
    $checked++
    [byte[]]$native = [IO.File]::ReadAllBytes($nativePath)
    [byte[]]$generated = [DocxportNet.DxpDocExport]::Export($source)
    foreach ($route in @('native', 'generated')) {
        [byte[]]$doc = if ($route -eq 'native') { $native } else { $generated }
        [byte[]]$projection = [DocxportNet.DxpDocToDocx]::Project($doc).DocxBytes
        foreach ($hop in @(1, 2)) {
            if ($hop -eq 2) {
                $projection = [DocxportNet.DxpDocToDocx]::Project(
                    [DocxportNet.DxpDocExport]::Export($projection)).DocxBytes
            }
            $observed = @(Get-PixelHashes $projection ($fixture.BaseName +
                '-' + $route + '-hop' + $hop))
            $rows += [pscustomobject]@{
                Fixture = $fixture.BaseName; Route = $route; Hop = $hop
                SourceImages = $expected.Count; OutputImages = $observed.Count
                PixelEqual = (@(Compare-Object $expected $observed).Count -eq 0)
            }
        }
    }
}
$report = Join-Path $ResultsDirectory 'pixel-results.csv'
$rows | Export-Csv -LiteralPath $report -NoTypeInformation
$failures = @($rows | Where-Object { -not $_.PixelEqual })
Write-Host "Image pixel gate: $checked paired fixtures; $($rows.Count) routes; $($failures.Count) differences. Report: $report"
$failures | Format-Table Fixture, Route, Hop, SourceImages, OutputImages -AutoSize
if ($checked -lt 56) { throw "Only $checked paired image fixtures were checked." }
$wordNativeDifferences = @(
    'WordBmpRle4DeltaAllStories'
    'WordFloatingTiffAllStories'
    'WordTiffAllStories'
)
$unexplained = @($failures | Where-Object {
    $_.Route -ne 'native' -or $_.Fixture -notin $wordNativeDifferences
})
if ($unexplained.Count -gt 0) {
    throw "$($unexplained.Count) image pixel routes differ unexpectedly. Report: $report"
}
if ($failures.Count -gt 0) {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {
        foreach ($name in @($failures.Fixture | Sort-Object -Unique)) {
            $nativePath = Join-Path $fixtures ($name + '.doc')
            $wordDocx = Join-Path $ResultsDirectory ($name + '-word-reopened.docx')
            $document = $word.Documents.Open($nativePath, $false, $true)
            try { $document.SaveAs2($wordDocx, 16) }
            finally { $document.Close($false) }
            $wordPixels = @(Get-PixelHashes ([IO.File]::ReadAllBytes($wordDocx)) ($name + '-word-reopened'))
            [byte[]]$native = [IO.File]::ReadAllBytes($nativePath)
            [byte[]]$projected = [DocxportNet.DxpDocToDocx]::Project($native).DocxBytes
            foreach ($hop in @(1, 2)) {
                if ($hop -eq 2) {
                    $projected = [DocxportNet.DxpDocToDocx]::Project(
                        [DocxportNet.DxpDocExport]::Export($projected)).DocxBytes
                }
                $ours = @(Get-PixelHashes $projected ($name + '-native-word-check-' + $hop))
                if (@(Compare-Object $wordPixels $ours).Count -ne 0) {
                    throw "$name native hop $hop differs from Word's reopened DOC image."
                }
            }
        }
    }
    finally { $word.Quit() }
}
Write-Host "Generated DOC pixel gate: $checked/$checked fixtures passed through both hops."
Write-Host "$($failures.Count) native differences match Word's reopened DOC pixels."
