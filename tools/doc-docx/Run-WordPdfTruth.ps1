param(
    [string]$FixtureDirectory = (Join-Path $PSScriptRoot '../../fixtures/docx'),
    [string]$TruthDirectory = (Join-Path $PSScriptRoot '../../truth/pdf-word'),
    [string]$NameFilter = '*',
    [switch]$MissingOnly
)

$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'Word PDF truth generation requires Microsoft Word on Windows.' }

$fixtureRoot = (Resolve-Path -LiteralPath $FixtureDirectory).Path
$truthRoot = [IO.Path]::GetFullPath($TruthDirectory)
$sources = @(Get-ChildItem -LiteralPath $fixtureRoot -Filter '*.docx' -File -Recurse |
    Where-Object { $_.Name -notlike '~$*' -and $_.BaseName -like $NameFilter } |
    Sort-Object FullName)
if ($sources.Count -eq 0) { throw "No DOCX fixtures matched '$NameFilter'." }

$word = $null
$exported = 0
$skipped = 0
$failures = [System.Collections.Generic.List[string]]::new()
try {
    foreach ($source in $sources) {
        $relativePath = [IO.Path]::GetRelativePath($fixtureRoot, $source.FullName)
        $target = Join-Path $truthRoot ([IO.Path]::ChangeExtension($relativePath, '.pdf'))
        if ($MissingOnly -and (Test-Path -LiteralPath $target -PathType Leaf)) {
            $skipped++
            continue
        }

        $document = $null
        try {
            if (-not $word) {
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $word.DisplayAlerts = 0
            }
            $targetDirectory = Split-Path -Parent $target
            New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
            $document = $word.Documents.Open($source.FullName, $false, $true)
            $document.PrintRevisions = $false
            $document.ExportAsFixedFormat($target, 17)
            $exported++
            Write-Output "Word: $relativePath -> $target"
        }
        catch {
            $failures.Add($relativePath)
            Write-Warning "Word export failed for $relativePath`: $($_.Exception.Message)"
        }
        finally {
            if ($document) {
                try { $document.Close($false) } catch { Write-Warning "Could not close $relativePath`: $($_.Exception.Message)" }
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
            }
        }
    }
}
finally {
    if ($word) {
        try { $word.Quit() } catch { Write-Warning "Could not quit Word: $($_.Exception.Message)" }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}
Write-Output "Word PDF truth: $exported exported, $skipped already present, $($failures.Count) failed (of $($sources.Count) selected)."
if ($failures.Count -gt 0) { throw "Word export failed for: $($failures -join ', ')" }
