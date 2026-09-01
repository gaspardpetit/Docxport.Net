[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$toolDirectory = $PSScriptRoot
$repositoryRoot = Resolve-Path (Join-Path $toolDirectory '..\..')
$inputDirectory = Join-Path $repositoryRoot 'DocxportNet.Tests\Fixtures\Omml\Upstream\Plurimath'
$outputDirectory = Join-Path $repositoryRoot 'DocxportNet.Tests\Fixtures\Omml\OracleGenerated'

$additionalPaths = @()
if ($IsWindows -and $env:USERPROFILE) {
    $scoopRoot = Join-Path $env:USERPROFILE 'scoop\apps'
    $additionalPaths += @(
        (Join-Path $scoopRoot 'ruby\current\bin'),
        (Join-Path $scoopRoot 'msys2\current\ucrt64\bin'),
        (Join-Path $scoopRoot 'msys2\current\usr\bin')
    ) | Where-Object { Test-Path -LiteralPath $_ }
}

if ($additionalPaths.Count -gt 0) {
    $env:PATH = ($additionalPaths + $env:PATH) -join [IO.Path]::PathSeparator
}

if (-not (Get-Command ruby -ErrorAction SilentlyContinue) -or
    -not (Get-Command bundle -ErrorAction SilentlyContinue)) {
    throw 'Bundler is required. Install Ruby 3.0+ and run `gem install bundler`.'
}

Push-Location $toolDirectory
try {
    & bundle install
    if ($LASTEXITCODE -ne 0) {
        throw "bundle install failed with exit code $LASTEXITCODE."
    }

    & bundle exec ruby generate.rb $inputDirectory $outputDirectory
    if ($LASTEXITCODE -ne 0) {
        throw "Oracle generation failed with exit code $LASTEXITCODE."
    }
}
finally {
    Pop-Location
}
