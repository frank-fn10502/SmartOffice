param(
    [switch]$User,
    [string]$Project,
    [string]$Dest,
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
安裝 SmartOffice.Hub Outlook Agents SKILL。

用法:
  pwsh ./scripts/install.ps1 [options]

預設:
  安裝到 user skill folder:
  `$env:CODEX_HOME\skills\smartoffice-hub-outlook
  或 `$HOME\.codex\skills\smartoffice-hub-outlook`

Options:
  -User
      安裝到 user skill folder。這是預設行為。

  -Project <path>
      安裝到指定 project 的 .codex\skills folder。
      例如: -Project C:\path\to\project

  -Dest <path>
      安裝到指定 skills root 或完整 skill folder。
      若 path basename 是 smartoffice-hub-outlook，會直接使用該 path；
      否則會安裝到 <path>\smartoffice-hub-outlook。

  -Force
      保留相容參數；目前安裝預設就是全新重裝。

  -DryRun
      只顯示將會安裝的位置，不寫入檔案。

  -Help
      顯示說明。

範例:
  pwsh ./scripts/install.ps1
  pwsh ./scripts/install.ps1 -Project C:\path\to\project
  pwsh ./scripts/install.ps1 -Dest C:\temp\codex-skills
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

if ($Help) {
    Show-Usage
    exit 0
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

if ($Project -and $Dest) {
    Write-Error "錯誤: -Project 與 -Dest 不能同時使用。"
}

if ($Project) {
    $projectDir = Resolve-FullPath $Project
    $destDir = Join-Path $projectDir ".codex\skills\$SkillName"
}
elseif ($Dest) {
    $destDir = Resolve-Destination $Dest
}
else {
    $codexHome = $env:CODEX_HOME
    if ([string]::IsNullOrWhiteSpace($codexHome)) {
        $codexHome = Join-Path $HOME ".codex"
    }

    $destDir = Join-Path $codexHome "skills\$SkillName"
}

Write-Output "來源: $skillDir"
Write-Output "目標: $destDir"

if ($DryRun) {
    exit 0
}

if (Test-Path -LiteralPath $destDir) {
    $targetMarker = Join-Path $destDir $MarkerFile
    if (-not (Test-Path -LiteralPath $targetMarker)) {
        Write-Error "錯誤: 目標已存在，但缺少 $MarkerFile；為避免覆蓋其他同名 skill，已停止。"
    }

    $targetSkillId = (Get-Content -LiteralPath $targetMarker -Raw).Trim()
    if ($targetSkillId -ne $SkillId) {
        Write-Error "錯誤: 目標 skill id 不符，為避免覆蓋其他同名 skill，已停止。預期: $SkillId；實際: $targetSkillId。"
    }

    Write-Output "移除既有安裝: $destDir"
    Remove-Item -LiteralPath $destDir -Recurse -Force
}

$parentDir = Split-Path -Parent $destDir
New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
Copy-Item -LiteralPath $skillDir -Destination $destDir -Recurse

Write-Output "已全新安裝 $SkillName。"
