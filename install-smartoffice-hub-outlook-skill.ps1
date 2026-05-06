$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $PSCommandPath
$installer = Join-Path $scriptDir "docs\ai_plugin\skills\smartoffice-hub-outlook\scripts\install.ps1"

& $installer @args
exit $LASTEXITCODE
