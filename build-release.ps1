[CmdletBinding()]
param(
    [string]$OutputDirectory = 'dist',
    [string]$ZipName = 'WindowLayout.zip'
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$stagingRoot = Join-Path $repoRoot 'build\WindowLayout'
$outputRoot = Join-Path $repoRoot $OutputDirectory
$zipPath = Join-Path $outputRoot $ZipName

if (Test-Path -LiteralPath $stagingRoot) {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force
}

if (-not (Test-Path -LiteralPath $outputRoot)) {
    New-Item -ItemType Directory -Path $outputRoot | Out-Null
}

New-Item -ItemType Directory -Path $stagingRoot | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stagingRoot 'WindowLayout') | Out-Null

Copy-Item -LiteralPath (Join-Path $repoRoot 'WindowLayout.cmd') -Destination $stagingRoot
Copy-Item -LiteralPath (Join-Path $repoRoot 'WindowLayout\WindowLayout.ps1') -Destination (Join-Path $stagingRoot 'WindowLayout')
Copy-Item -LiteralPath (Join-Path $repoRoot 'WindowLayout\readme.txt') -Destination (Join-Path $stagingRoot 'WindowLayout')

if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -Path (Join-Path $stagingRoot '*') -DestinationPath $zipPath -CompressionLevel Optimal
Write-Host "Created $zipPath"
Pause