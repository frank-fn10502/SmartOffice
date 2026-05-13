param(
    [string]$BaseUrl = $env:SMARTOFFICE_OUTLOOK_BASE_URL,
    [int]$LookbackHours = 168,
    [int]$MaxCount = 30,
    [int]$Take = 100
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $PSCommandPath
$apiScript = Join-Path (Split-Path -Parent $scriptDir) "outlook-api.ps1"
$inboxScript = Join-Path $scriptDir "inbox.ps1"

$commonArgs = @()
if (-not [string]::IsNullOrWhiteSpace($BaseUrl)) {
    $commonArgs += @("-BaseUrl", $BaseUrl)
}

$inbox = & pwsh $inboxScript @commonArgs -Take $Take | ConvertFrom-Json
$folderPath = [string]$inbox.folder.folderPath
$bodyObject = @{
    folderPath = $folderPath
    lookbackHours = $LookbackHours
    maxCount = $MaxCount
}
$body = $bodyObject | ConvertTo-Json -Depth 5 -Compress

$result = & pwsh $apiScript @commonArgs request-fetch /api/outlook/request-mails $body -Take $Take | ConvertFrom-Json
$mails = @($result.fetchResult.pages | ForEach-Object { $_.data.mails } | Where-Object { $_ })

[pscustomobject]@{
    folderPath = $folderPath
    request = $bodyObject
    mails = $mails
} | ConvertTo-Json -Depth 80
