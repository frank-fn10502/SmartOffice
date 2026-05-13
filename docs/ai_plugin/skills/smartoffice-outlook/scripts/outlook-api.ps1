param(
    [string]$BaseUrl = $env:SMARTOFFICE_OUTLOOK_BASE_URL,
    [Parameter(Position = 0)][string]$Command,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Args
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($BaseUrl)) {
    $BaseUrl = "http://localhost:2805"
}
$BaseUrl = $BaseUrl.TrimEnd("/")
$script:Take = 100

function Show-Usage {
    @"
SmartOffice Outlook HTTP API helper.

Outputs JSON to stdout. Diagnostic text goes to stderr.

Usage:
  pwsh outlook-api.ps1 status
  pwsh outlook-api.ps1 post <path> <json-or-@file>
  pwsh outlook-api.ps1 fetch <fetch-result-path> <request-id> [-Take N]
  pwsh outlook-api.ps1 request-fetch <request-path> <json-or-@file> [-Take N]

Examples:
  pwsh ./scripts/outlook-api.ps1 status
  pwsh ./scripts/outlook-api.ps1 request-fetch /api/outlook/request-calendar '{ "daysForward": 31, "startDate": null, "endDate": null }'
  pwsh ./scripts/outlook-api.ps1 post /api/outlook/request-mail-search @request.json

Rules implemented by this helper:
  - request-* responses are not treated as success until paired fetch-result completes.
  - fetch-result pagination continues while next.hasMore=true, even when state=completed.
  - failed, unavailable, and timeout states stop the helper with a non-zero exit code.

Workflow recipes live in scripts/recipes/. Use those when a task has an established
sequence such as locating Inbox before reading recent mail.
"@
}

function ConvertTo-JsonText {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)

    if ($Value.StartsWith("@")) {
        $file = $Value.Substring(1)
        if (-not (Test-Path -LiteralPath $file)) {
            throw "JSON file not found: $file"
        }
        return Get-Content -LiteralPath $file -Raw
    }

    return $Value
}

function Invoke-ApiGet {
    param([Parameter(Mandatory = $true)][string]$Path)
    Invoke-RestMethod -Method Get -Uri "$BaseUrl$Path"
}

function Invoke-ApiPost {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Body
    )

    try {
        Invoke-RestMethod -Method Post -Uri "$BaseUrl$Path" -ContentType "application/json" -Body $Body
    }
    catch {
        if ($_.Exception.Response -and $_.ErrorDetails.Message) {
            $_.ErrorDetails.Message | ConvertFrom-Json
            exit 1
        }
        throw
    }
}

function Invoke-FetchAll {
    param(
        [Parameter(Mandatory = $true)][string]$Endpoint,
        [Parameter(Mandatory = $true)][string]$RequestId,
        [Parameter(Mandatory = $true)][int]$Take
    )

    $cursor = ""
    $pages = New-Object System.Collections.Generic.List[object]

    for ($i = 0; $i -lt 200; $i++) {
        $body = @{
            requestId = $RequestId
            cursor = $cursor
            take = $Take
        } | ConvertTo-Json -Depth 5 -Compress

        $page = Invoke-ApiPost -Path $Endpoint -Body $body

        if ($page.state -in @("failed", "unavailable", "timeout")) {
            $pages.Add($page)
            [pscustomobject]@{
                endpoint = $Endpoint
                requestId = $RequestId
                state = "failed"
                pages = $pages
            }
            exit 1
        }

        if ($page.next -and $page.next.hasMore) {
            $pages.Add($page)
            $cursor = [string]$page.next.cursor
            continue
        }

        if ($page.state -eq "completed") {
            $pages.Add($page)
            return [pscustomobject]@{
                endpoint = $Endpoint
                requestId = $RequestId
                state = "completed"
                pages = $pages
            }
        }

        Start-Sleep -Milliseconds 200
    }

    [pscustomobject]@{
        endpoint = $Endpoint
        requestId = $RequestId
        state = "timeout"
        message = "fetch-result loop exceeded 200 attempts"
        pages = $pages
    }
    exit 1
}

function Invoke-RequestFetch {
    param(
        [Parameter(Mandatory = $true)][string]$RequestPath,
        [Parameter(Mandatory = $true)][string]$RequestBody,
        [Parameter(Mandatory = $true)][int]$Take
    )

    $requestResponse = Invoke-ApiPost -Path $RequestPath -Body $RequestBody
    $requestId = [string]$requestResponse.requestId
    $endpoint = [string]$requestResponse.data.fetchResultEndpoint

    if ([string]::IsNullOrWhiteSpace($requestId) -or [string]::IsNullOrWhiteSpace($endpoint)) {
        $requestResponse
        throw "request response did not include requestId or data.fetchResultEndpoint"
    }

    [pscustomobject]@{
        requestResponse = $requestResponse
        fetchResult = Invoke-FetchAll -Endpoint $endpoint -RequestId $requestId -Take $Take
    }
}

function Read-OptionValue {
    param(
        [Parameter(Mandatory = $true)][string[]]$Values,
        [Parameter(Mandatory = $true)][ref]$Index,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if ($Index.Value + 1 -ge $Values.Count) {
        throw "$Name requires a value."
    }
    $Index.Value++
    return $Values[$Index.Value]
}

if ([string]::IsNullOrWhiteSpace($Command) -or $Command -in @("-h", "--help", "help")) {
    Show-Usage
    exit 0
}

switch ($Command) {
    "status" {
        Invoke-ApiGet -Path "/api/outlook/admin/status" | ConvertTo-Json -Depth 20
    }
    "post" {
        if ($Args.Count -lt 2) { Show-Usage; exit 2 }
        Invoke-ApiPost -Path $Args[0] -Body (ConvertTo-JsonText $Args[1]) | ConvertTo-Json -Depth 50
    }
    "fetch" {
        if ($Args.Count -lt 2) { Show-Usage; exit 2 }
        $endpoint = $Args[0]
        $requestId = $Args[1]
        for ($i = 2; $i -lt $Args.Count; $i++) {
            switch ($Args[$i]) {
                "-Take" { $script:Take = [int](Read-OptionValue -Values $Args -Index ([ref]$i) -Name "-Take") }
                "--take" { $script:Take = [int](Read-OptionValue -Values $Args -Index ([ref]$i) -Name "--take") }
                default { throw "Unknown option: $($Args[$i])" }
            }
        }
        Invoke-FetchAll -Endpoint $endpoint -RequestId $requestId -Take $script:Take | ConvertTo-Json -Depth 80
    }
    "request-fetch" {
        if ($Args.Count -lt 2) { Show-Usage; exit 2 }
        $requestPath = $Args[0]
        $requestBody = ConvertTo-JsonText $Args[1]
        for ($i = 2; $i -lt $Args.Count; $i++) {
            switch ($Args[$i]) {
                "-Take" { $script:Take = [int](Read-OptionValue -Values $Args -Index ([ref]$i) -Name "-Take") }
                "--take" { $script:Take = [int](Read-OptionValue -Values $Args -Index ([ref]$i) -Name "--take") }
                default { throw "Unknown option: $($Args[$i])" }
            }
        }
        Invoke-RequestFetch -RequestPath $requestPath -RequestBody $requestBody -Take $script:Take | ConvertTo-Json -Depth 80
    }
    default {
        Show-Usage
        exit 2
    }
}
