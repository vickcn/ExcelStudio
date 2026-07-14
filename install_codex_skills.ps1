param(
  [string]$CodexHome = $env:CODEX_HOME,
  [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $CodexHome) {
  $CodexHome = Join-Path $env:USERPROFILE '.codex'
}

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcSkillsDir = Join-Path $repoRoot '.codex\\skills'
if (-not (Test-Path $srcSkillsDir)) {
  throw \"Source skills dir not found: $srcSkillsDir\"
}

$dstSkillsDir = Join-Path $CodexHome 'skills'
New-Item -ItemType Directory -Force -Path $dstSkillsDir | Out-Null

$skillDirs = Get-ChildItem -Path $srcSkillsDir -Directory
if (-not $skillDirs) {
  throw \"No skill folders found under: $srcSkillsDir\"
}

foreach ($d in $skillDirs) {
  $dst = Join-Path $dstSkillsDir $d.Name

  if ((Test-Path $dst) -and (-not $Force)) {
    Write-Host \"Skip (exists, use -Force): $dst\"
    continue
  }

  if (Test-Path $dst) {
    Remove-Item -Recurse -Force -LiteralPath $dst
  }

  Copy-Item -Recurse -Force -LiteralPath $d.FullName -Destination $dst
  Write-Host \"Installed: $dst\"
}

