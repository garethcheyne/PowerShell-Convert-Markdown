<#
.SYNOPSIS
    Builds the release package: dist/ConvertMarkdown-<version>.zip, its .sha256
    and release-notes.md (the version's section of CHANGELOG.md).

.DESCRIPTION
    Used by the Release workflow and by CI (as a dry run), and runnable locally.

    Only files tracked by git are packaged, so private themes under theming/
    (gitignored) can never reach a release. Fails if the version strings disagree
    or CHANGELOG.md has no section for the version.

.EXAMPLE
    ./build/Build-Release.ps1
#>
[CmdletBinding()]
param(
    [string] $OutDir = (Join-Path (Split-Path -Parent $PSScriptRoot) 'dist')
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

# --- version: one value, everywhere -------------------------------------------------
$manifest = Import-PowerShellDataFile (Join-Path $root 'ConvertMarkdown.psd1')
$version = $manifest.PrivateData.DisplayVersion
if ($version -notmatch '^\d{4}\.\d{2}\.\d{2}(\.\d+)?$') { throw "DisplayVersion '$version' is not yyyy.MM.dd[.n]." }
if ($manifest.ModuleVersion -ne $version) { throw "ModuleVersion '$($manifest.ModuleVersion)' differs from DisplayVersion '$version'." }
$psm1 = Get-Content (Join-Path $root 'ConvertMarkdown.psm1') -Raw
if ($psm1 -notmatch [regex]::Escape("`$script:Version = '$version'")) { throw "ConvertMarkdown.psm1 `$script:Version is not '$version'." }
$readme = Get-Content (Join-Path $root 'README.md') -Raw
if ($readme -notmatch [regex]::Escape("**Version:** $version")) { throw "README.md does not show **Version:** $version." }

# --- release notes from the changelog ---------------------------------------------------
$changelog = Get-Content (Join-Path $root 'CHANGELOG.md') -Raw
$m = [regex]::Match($changelog, '(?ms)^## ' + [regex]::Escape($version) + '\s*$\r?\n(.*?)(?=^## |\z)')
if (-not $m.Success -or -not $m.Groups[1].Value.Trim()) { throw "CHANGELOG.md has no section '## $version'." }
$notes = $m.Groups[1].Value.Trim()

# --- stage tracked files ------------------------------------------------------------------
$include = '^(ConvertMarkdown\.ps[dm]1|Install\.ps1|README\.md|LICENSE|CHANGELOG\.md|style\.(example|schema)\.json|themes/.+|examples/.+)$'
$files = git -C $root ls-files | Where-Object { $_ -match $include }
if (-not $files) { throw 'git ls-files returned nothing to package.' }
if ($files -match '^theming/') { throw 'A private theme is tracked by git; refusing to package it.' }

if (Test-Path $OutDir) { Remove-Item $OutDir -Recurse -Force }
$stage = Join-Path $OutDir 'ConvertMarkdown'
foreach ($f in $files) {
    $dest = Join-Path $stage $f
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
    Copy-Item -LiteralPath (Join-Path $root $f) -Destination $dest
}

$zip = Join-Path $OutDir "ConvertMarkdown-$version.zip"
Compress-Archive -Path $stage -DestinationPath $zip -CompressionLevel Optimal
$hash = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -LiteralPath "$zip.sha256" -Value "$hash  $(Split-Path -Leaf $zip)" -Encoding ascii -NoNewline

$install = @"
## Install or update

``````powershell
Update-ConvertMarkdown          # already installed: fetch this release
``````

New install: download **ConvertMarkdown-$version.zip**, extract it, and run ``.\ConvertMarkdown\Install.ps1``.

SHA-256: ``$hash``
"@
Set-Content -LiteralPath (Join-Path $OutDir 'release-notes.md') -Value ($notes + "`n`n" + $install) -Encoding utf8

[pscustomobject]@{
    Version = $version
    Zip     = $zip
    Sha256  = $hash
    Files   = $files.Count
    Notes   = (Join-Path $OutDir 'release-notes.md')
}
