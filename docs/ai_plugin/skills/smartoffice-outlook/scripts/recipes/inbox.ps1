param(
    [string]$BaseUrl = $env:SMARTOFFICE_OUTLOOK_BASE_URL,
    [int]$Take = 100
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $PSCommandPath
$apiScript = Join-Path (Split-Path -Parent $scriptDir) "outlook-api.ps1"

$body = @{
    name = ""
    folderPath = ""
    folderType = "Inbox"
    storeId = ""
    includeHidden = $false
    maxResults = 20
} | ConvertTo-Json -Depth 5 -Compress

$argsList = @()
if (-not [string]::IsNullOrWhiteSpace($BaseUrl)) {
    $argsList += @("-BaseUrl", $BaseUrl)
}
$argsList += @("request-fetch", "/api/outlook/request-find-folder", $body, "-Take", $Take)

$result = & pwsh $apiScript @argsList | ConvertFrom-Json
$folders = @($result.fetchResult.pages | ForEach-Object { $_.data.folders } | Where-Object { $_ })
$matchCount = @($result.fetchResult.pages | ForEach-Object { $_.data.matchCount } | Measure-Object -Sum).Sum
$isAmbiguous = @($result.fetchResult.pages | ForEach-Object { $_.data.isAmbiguous } | Where-Object { $_ }).Count -gt 0
if ($matchCount -ne 1 -or $isAmbiguous -or $folders.Count -lt 1) {
    $result | ConvertTo-Json -Depth 80
    throw "Inbox could not be uniquely located."
}

[pscustomobject]@{
    folder = $folders[0]
} | ConvertTo-Json -Depth 50
