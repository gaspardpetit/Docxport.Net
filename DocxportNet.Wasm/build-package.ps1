$ErrorActionPreference = 'Stop'

$projectDirectory = $PSScriptRoot
$projectPath = Join-Path $projectDirectory 'DocxportNet.Wasm.csproj'
$packageDirectory = Join-Path $projectDirectory 'bin\Release\net10.0\publish\wwwroot'
$resolvedProjectDirectory = [System.IO.Path]::GetFullPath($projectDirectory)
$resolvedPackageDirectory = [System.IO.Path]::GetFullPath($packageDirectory)

if (-not $resolvedPackageDirectory.StartsWith($resolvedProjectDirectory, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'Refusing to clean a package directory outside DocxportNet.Wasm.'
}

if (Test-Path -LiteralPath $resolvedPackageDirectory) {
    Remove-Item -LiteralPath $resolvedPackageDirectory -Recurse -Force
}

dotnet publish $projectPath -c Release
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed with exit code $LASTEXITCODE." }

Write-Host "npm package ready at $resolvedPackageDirectory"
