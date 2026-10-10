#Requires -Version 5.1
$ErrorActionPreference = "Stop"
Set-Location -LiteralPath (Join-Path $PSScriptRoot "..")
python scripts\build_exe.py
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Invoke-Item dist
