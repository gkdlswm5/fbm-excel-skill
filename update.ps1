# One-shot update script for the fbm-excel skill.
#
# Workflow:
#   1. You edit files under C:\Users\Andrew.Kim\.claude\skills\fbm-excel\ (the live skill).
#   2. Run this script. It will:
#      a. Mirror the live skill folder into this repo (C:\ak\fbm-excel-skill\)
#      b. Re-package the skill as fbm-excel.skill (single file for Claude.ai upload)
#      c. Commit + push to GitHub
#      d. Print a reminder to re-upload fbm-excel.skill to Claude.ai
#
# Usage:
#   .\update.ps1 "your commit message"
#   .\update.ps1                          # uses default "Update skill"

param(
    [string]$Message = "Update skill"
)

$ErrorActionPreference = 'Stop'

# Paths
$liveSkill   = "$env:USERPROFILE\.claude\skills\fbm-excel"
$repoDir     = $PSScriptRoot                      # this script lives in the repo
$skillPkgDir = "$env:TEMP\fbm-excel-build"        # disposable staging dir for the packaged .skill
New-Item -ItemType Directory -Force -Path $skillPkgDir | Out-Null
$creatorDir  = 'C:\Users\Andrew.Kim\AppData\Roaming\Claude\local-agent-mode-sessions\skills-plugin\b7fb31a3-a2c4-43e6-aae2-0e3853ab3d2f\cd99ea31-8d77-40cc-ae83-6b93b89950d8\skills\skill-creator'
$py          = "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe"

Write-Host "→ Mirroring live skill into repo..." -ForegroundColor Cyan
# Remove old skill content from repo (but keep README, .gitignore, update.ps1, .git, .skill)
Get-ChildItem $repoDir -Force | Where-Object {
    $_.Name -notin @('.git', 'README.md', '.gitignore', 'update.ps1', 'fbm-excel.skill', 'source')
} | Remove-Item -Recurse -Force

# Copy fresh content
Copy-Item "$liveSkill\*" $repoDir -Recurse -Force

Write-Host "→ Re-packaging skill as fbm-excel.skill..." -ForegroundColor Cyan
$env:PYTHONIOENCODING = 'utf-8'
Push-Location $creatorDir
try {
    & $py -m scripts.package_skill $liveSkill $skillPkgDir
    if ($LASTEXITCODE -ne 0) { throw "package_skill failed" }
}
finally {
    Pop-Location
}

# Copy the fresh .skill into the repo too (so the repo always has the latest distributable)
Copy-Item "$skillPkgDir\fbm-excel.skill" "$repoDir\fbm-excel.skill" -Force

Write-Host "→ Committing and pushing to GitHub..." -ForegroundColor Cyan
Push-Location $repoDir
try {
    git add -A
    $status = git status --porcelain
    if (-not $status) {
        Write-Host "✓ No changes to commit." -ForegroundColor Yellow
        return
    }
    git commit -m $Message
    if ($LASTEXITCODE -ne 0) { throw "git commit failed" }
    git push
    if ($LASTEXITCODE -ne 0) { throw "git push failed" }
}
finally {
    Pop-Location
}

Write-Host ""
Write-Host "✓ Done." -ForegroundColor Green
Write-Host "  Repo:   https://github.com/gkdlswm5/fbm-excel-skill" -ForegroundColor Green
Write-Host "  Skill:  $skillPkgDir\fbm-excel.skill" -ForegroundColor Green
Write-Host ""
Write-Host "→ Next step: re-upload fbm-excel.skill to Claude.ai" -ForegroundColor Yellow
Write-Host "  Settings → Capabilities → Skills → Upload (overwrites existing)" -ForegroundColor Yellow
