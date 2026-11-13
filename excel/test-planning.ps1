param([Parameter(Mandatory=$true)][string]$WorkbookPath)
& (Join-Path $PSScriptRoot "test-coverage.ps1") -WorkbookPath $WorkbookPath
