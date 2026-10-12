# build.ps1 - builds the critic-export release zip into dist\.
# critic-export runs from source (Python + Pandoc, see README), so the zip is the
# tracked files minus tests and CI, in one critic-export\ folder. Your own
# settings.toml is never tracked, so unzipping over an old copy keeps it.
# Works in Windows PowerShell 5.1 and 7.
#
# Version: -Version, else the pushed tag (v2026.10.12 -> 2026.10.12), else today.
param([string]$Version)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$dist = Join-Path $root 'dist'

if (-not $Version) {
    $Version = if ($env:GITHUB_REF_NAME -match '^v(.+)$') { $Matches[1] } else { Get-Date -Format 'yyyy.M.d' }
}
Write-Host "Building critic-export $Version"

$files = @(git -C $root ls-files) | Where-Object { $_ -notmatch '^(\.git|tests/|pytest\.ini$|build\.ps1$)' }
if ($files.Count -eq 0) { throw 'git ls-files returned nothing - run from a git checkout' }

if (Test-Path $dist) { Remove-Item $dist -Recurse -Force }
$stage = Join-Path $dist 'critic-export'
foreach ($f in $files) {
    $to = Join-Path $stage $f
    New-Item -ItemType Directory -Force (Split-Path $to) | Out-Null
    Copy-Item (Join-Path $root $f) $to
}

$zip = Join-Path $dist "critic-export-$Version.zip"
Compress-Archive -Path $stage -DestinationPath $zip -Force
Write-Host "Done: $zip"
