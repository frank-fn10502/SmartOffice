param(
    [switch]$User,
    [string]$Project,
    [string]$Dest,
    [string]$Tools,
    [string[]]$Tool,
    [switch]$Force,
    [switch]$DryRun,
    [switch]$Help
)

$ErrorActionPreference = "Stop"
$SkillName = "smartoffice-hub-outlook"
$SkillId = "smartoffice-hub-outlook.skill.smartoffice-hub.2026-05"
$MarkerFile = ".smartoffice-skill-id"

function Show-Usage {
    @"
安裝 SmartOffice Outlook Agents SKILL。

用法:
  pwsh ./install-smartoffice-hub-outlook-skill.ps1 [options]

直接呼叫 skill 內部 installer:
  pwsh ./docs/ai_plugin/skills/smartoffice-hub-outlook/scripts/install.ps1 [options]

預設:
  同時複製 SKILL folder 到 codex、copilot、opencode 的 user skill 位置。
  不會產生或修改 AGENTS.md、copilot-instructions.md、*.instructions.md 等規則檔。

User-level 目標:
  codex:   `$env:CODEX_HOME\skills\smartoffice-hub-outlook 或 `$HOME\.codex\skills\smartoffice-hub-outlook
  copilot: `$HOME\.copilot\skills\smartoffice-hub-outlook
  opencode:`$env:XDG_CONFIG_HOME\opencode\skills\smartoffice-hub-outlook 或 `$HOME\.config\opencode\skills\smartoffice-hub-outlook

Project-level 目標:
  codex:   <project>\.codex\skills\smartoffice-hub-outlook
  copilot: <project>\.github\skills\smartoffice-hub-outlook
  opencode:<project>\.opencode\skills\smartoffice-hub-outlook

Options:
  -User
      安裝到 user skill folder。這是預設行為。

  -Project <path>
      安裝到指定 project 的 tool-specific skill folder。

  -Tools <list>
      逗號分隔的工具清單。可用值: codex,copilot,opencode,all。
      預設: all。

  -Tool <name>
      加入單一工具。可重複使用。

  -Dest <path>
      只安裝 Codex skill 到指定 skills root 或完整 skill folder。
      若 path basename 是 smartoffice-hub-outlook，會直接使用該 path；
      否則會安裝到 <path>\smartoffice-hub-outlook。

  -Force
      保留相容參數；目前安裝預設就是全新重裝。

  -DryRun
      只顯示將會安裝的位置，不寫入檔案。

  -Help
      顯示說明。

範例:
  pwsh ./install-smartoffice-hub-outlook-skill.ps1
  pwsh ./install-smartoffice-hub-outlook-skill.ps1 -Project C:\path\to\project
  pwsh ./install-smartoffice-hub-outlook-skill.ps1 -Tools codex,opencode
  pwsh ./install-smartoffice-hub-outlook-skill.ps1 -Tool copilot -Project C:\path\to\project
"@
}

function Resolve-FullPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
}

function Resolve-Destination {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = Resolve-FullPath $Path
    if ((Split-Path -Leaf $fullPath) -eq $SkillName) {
        return $fullPath
    }

    return Join-Path $fullPath $SkillName
}

function Add-Tool {
    param([Parameter(Mandatory = $true)][string]$Name)

    switch ($Name) {
        "all" {
            $script:requestedCodex = $true
            $script:requestedCopilot = $true
            $script:requestedOpencode = $true
        }
        "codex" {
            $script:requestedCodex = $true
        }
        "copilot" {
            $script:requestedCopilot = $true
        }
        "opencode" {
            $script:requestedOpencode = $true
        }
        default {
            Write-Error "錯誤: 不支援的 tool: $Name"
        }
    }
}

function Add-ToolsCsv {
    param([Parameter(Mandatory = $true)][string]$Csv)

    foreach ($item in $Csv.Split(",")) {
        $name = $item.Trim()
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            Add-Tool $name
        }
    }
}

function Copy-SkillFolder {
    param(
        [Parameter(Mandatory = $true)][string]$SourceDir,
        [Parameter(Mandatory = $true)][string]$DestDir
    )

    if (Test-Path -LiteralPath $DestDir) {
        $targetMarker = Join-Path $DestDir $MarkerFile
        if (-not (Test-Path -LiteralPath $targetMarker)) {
            Write-Error "錯誤: 目標已存在，但缺少 $MarkerFile；為避免覆蓋其他同名 skill，已停止: $DestDir"
        }

        $targetSkillId = (Get-Content -LiteralPath $targetMarker -Raw).Trim()
        if ($targetSkillId -ne $SkillId) {
            Write-Error "錯誤: 目標 skill id 不符，為避免覆蓋其他同名 skill，已停止: $DestDir。預期: $SkillId；實際: $targetSkillId。"
        }

        Write-Output "移除既有安裝: $DestDir"
        Remove-Item -LiteralPath $DestDir -Recurse -Force
    }

    $parentDir = Split-Path -Parent $DestDir
    New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    Copy-Item -LiteralPath $SourceDir -Destination $DestDir -Recurse
}

if ($Help) {
    Show-Usage
    exit 0
}

if ($Project -and $Dest) {
    Write-Error "錯誤: -Project 與 -Dest 不能同時使用。"
}

$targetMode = "user"
if ($Project) {
    $targetMode = "project"
}
elseif ($Dest) {
    $targetMode = "dest"
}

$requestedCodex = $false
$requestedCopilot = $false
$requestedOpencode = $false
$toolsSet = $false

if ($Tools) {
    $toolsSet = $true
    Add-ToolsCsv $Tools
}

foreach ($toolName in $Tool) {
    if (-not $toolsSet) {
        $requestedCodex = $false
        $requestedCopilot = $false
        $requestedOpencode = $false
    }

    $toolsSet = $true
    Add-Tool $toolName
}

if (-not $toolsSet -and $targetMode -eq "dest") {
    $requestedCodex = $true
    $requestedCopilot = $false
    $requestedOpencode = $false
}
elseif (-not $toolsSet) {
    $requestedCodex = $true
    $requestedCopilot = $true
    $requestedOpencode = $true
}

if ($targetMode -eq "dest" -and ($requestedCopilot -or $requestedOpencode)) {
    Write-Error "錯誤: -Dest 只支援 Codex skill folder；請搭配 -Tools codex 使用。"
}

$scriptDir = Split-Path -Parent $PSCommandPath
$skillDir = Resolve-FullPath (Join-Path $scriptDir "..")
$sourceMarker = Join-Path $skillDir $MarkerFile

if (-not (Test-Path -LiteralPath $sourceMarker)) {
    Write-Error "錯誤: source skill 缺少識別檔 $MarkerFile。"
}

$sourceSkillId = (Get-Content -LiteralPath $sourceMarker -Raw).Trim()
if ($sourceSkillId -ne $SkillId) {
    Write-Error "錯誤: source skill id 不符，預期 $SkillId，實際 $sourceSkillId。"
}

$codexDest = $null
$copilotDest = $null
$opencodeDest = $null

switch ($targetMode) {
    "user" {
        $codexHome = $env:CODEX_HOME
        if ([string]::IsNullOrWhiteSpace($codexHome)) {
            $codexHome = Join-Path $HOME ".codex"
        }

        $xdgConfigHome = $env:XDG_CONFIG_HOME
        if ([string]::IsNullOrWhiteSpace($xdgConfigHome)) {
            $xdgConfigHome = Join-Path $HOME ".config"
        }

        $codexDest = Join-Path $codexHome "skills\$SkillName"
        $copilotDest = Join-Path $HOME ".copilot\skills\$SkillName"
        $opencodeDest = Join-Path $xdgConfigHome "opencode\skills\$SkillName"
    }
    "project" {
        $projectDir = Resolve-FullPath $Project
        $codexDest = Join-Path $projectDir ".codex\skills\$SkillName"
        $copilotDest = Join-Path $projectDir ".github\skills\$SkillName"
        $opencodeDest = Join-Path $projectDir ".opencode\skills\$SkillName"
    }
    "dest" {
        $codexDest = Resolve-Destination $Dest
    }
    default {
        Write-Error "錯誤: 未知 target mode: $targetMode"
    }
}

Write-Output "來源: $skillDir"
if ($requestedCodex) {
    Write-Output "Codex 目標: $codexDest"
}
if ($requestedCopilot) {
    Write-Output "Copilot 目標: $copilotDest"
}
if ($requestedOpencode) {
    Write-Output "opencode 目標: $opencodeDest"
}

if ($DryRun) {
    exit 0
}

if ($requestedCodex) {
    Copy-SkillFolder -SourceDir $skillDir -DestDir $codexDest
}
if ($requestedCopilot) {
    Copy-SkillFolder -SourceDir $skillDir -DestDir $copilotDest
}
if ($requestedOpencode) {
    Copy-SkillFolder -SourceDir $skillDir -DestDir $opencodeDest
}

Write-Output "已安裝 $SkillName。"
