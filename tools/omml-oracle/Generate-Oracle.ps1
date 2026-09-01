[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$toolDirectory = $PSScriptRoot
$repositoryRoot = Resolve-Path (Join-Path $toolDirectory '..\..')
$inputDirectory = Join-Path $repositoryRoot 'DocxportNet.Tests\Fixtures\Omml\Upstream\Plurimath'
$outputDirectory = Join-Path $repositoryRoot 'DocxportNet.Tests\Fixtures\Omml\OracleGenerated'

if (-not (Get-Command bundle -ErrorAction SilentlyContinue)) {
    throw 'Bundler is required. Install Ruby 3.0+ and run `gem install bundler`.'
}

Push-Location $toolDirectory
try {
    bundle install
    bundle exec ruby generate.rb $inputDirectory $outputDirectory
}
finally {
    Pop-Location
}
