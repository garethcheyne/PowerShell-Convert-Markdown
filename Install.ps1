<#
.SYNOPSIS
    Installs the ConvertMarkdown module so Convert-Markdown can be called from any
    directory, in any repo.

.DESCRIPTION
    Copies (or symlinks) the module into a PSModulePath directory. Once installed,
    PowerShell autoloads it on first use -- no Import-Module and no path juggling.

    Default scope is CurrentUser, which needs no elevation.

.PARAMETER Scope
    CurrentUser (default) or AllUsers. AllUsers requires an elevated session.

.PARAMETER Link
    Create a directory symlink back to this source folder instead of copying.
    Edits to the source then take effect immediately -- useful while iterating.
    Requires Developer Mode or an elevated session.

.PARAMETER Force
    Replace an existing installation without prompting.

.PARAMETER SkipConnectivityTest
    Do not probe the mermaid rendering service during install.

.PARAMETER StylePath
    Create a style file at this path without prompting. Use 'none' to skip the
    prompt entirely -- useful for unattended installs.

.PARAMETER DefaultTheme
    Make this theme the default for every document on this machine, by writing
    { "theme": "<name>" } to ~/.convertmarkdown/style.json. An existing style file
    is kept as style.json.bak.

    Themes are folders under theming/ beside this script (gitignored, so brand
    assets never reach the public repo). Every one found is copied to
    ~/.convertmarkdown/themes on install, whether or not it is the default.

.EXAMPLE
    .\Install.ps1

.EXAMPLE
    .\Install.ps1 -Link -Force

.EXAMPLE
    .\Install.ps1 -StylePath none -SkipConnectivityTest

.EXAMPLE
    .\Install.ps1 -Force -DefaultTheme acme
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string] $Scope = 'CurrentUser',

    [switch] $Link,

    [switch] $Force,

    [switch] $SkipConnectivityTest,

    [string] $StylePath,

    [string] $DefaultTheme
)

$ErrorActionPreference = 'Stop'

$moduleName = 'ConvertMarkdown'
$sourceDir = $PSScriptRoot

foreach ($required in @("$moduleName.psd1", "$moduleName.psm1")) {
    if (-not (Test-Path -LiteralPath (Join-Path $sourceDir $required))) {
        throw "Missing $required in $sourceDir - run this from the module folder."
    }
}

if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw "PowerShell 7+ is required (found $($PSVersionTable.PSVersion)). ConvertFrom-Markdown does not exist in Windows PowerShell 5.1."
}

# ------------------------------------------------------------- target location
$base = if ($Scope -eq 'AllUsers') {
    Join-Path $env:ProgramFiles 'PowerShell\Modules'
}
else {
    Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules'
}

if ($Scope -eq 'AllUsers') {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) { throw 'AllUsers scope requires an elevated PowerShell session.' }
}

$destDir = Join-Path $base $moduleName

# ------------------------------------------------------------- remove existing
if (Test-Path -LiteralPath $destDir) {
    $existing = Get-Item -LiteralPath $destDir
    $isSymlink = $existing.LinkType -eq 'SymbolicLink'
    $label = if ($isSymlink) { 'symlink' } else { 'copy' }

    if (-not $Force) {
        $answer = Read-Host "An existing $label is at $destDir. Replace it? [y/N]"
        if ($answer -notmatch '^(y|yes)$') {
            Write-Host 'Cancelled.' -ForegroundColor Yellow
            return
        }
    }

    if ($PSCmdlet.ShouldProcess($destDir, 'Remove existing installation')) {
        # A symlink must be removed with Delete(), or Remove-Item follows it and
        # deletes the source contents.
        if ($isSymlink) { $existing.Delete() }
        else { Remove-Item -LiteralPath $destDir -Recurse -Force }
    }
}

# -------------------------------------------------------------------- install
if (-not (Test-Path -LiteralPath $base)) {
    New-Item -ItemType Directory -Path $base -Force | Out-Null
}

if ($Link) {
    if ($PSCmdlet.ShouldProcess($destDir, "Symlink to $sourceDir")) {
        try {
            New-Item -ItemType SymbolicLink -Path $destDir -Target $sourceDir -ErrorAction Stop | Out-Null
        }
        catch {
            throw "Could not create the symlink. Enable Windows Developer Mode, run elevated, or install without -Link.`n$($_.Exception.Message)"
        }
        Write-Host "Linked  $destDir -> $sourceDir" -ForegroundColor Green
    }
}
else {
    if ($PSCmdlet.ShouldProcess($destDir, 'Copy module')) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        Copy-Item -Path (Join-Path $sourceDir "$moduleName.ps*1") -Destination $destDir -Force
        Write-Host "Installed to $destDir" -ForegroundColor Green
    }
}

# --------------------------------------------------------------- verification
Remove-Module $moduleName -Force -ErrorAction SilentlyContinue
$manifest = Test-ModuleManifest -Path (Join-Path $destDir "$moduleName.psd1")
$display = $manifest.PrivateData.DisplayVersion

Write-Host ''
Write-Host "$moduleName $display installed ($Scope)." -ForegroundColor Cyan
Write-Host ''
Write-Host '  Convert-Markdown .\README.md -TableOfContents'
Write-Host '  Convert-Markdown .\CHANGELOG.md -Format Both -Force'
Write-Host '  Convert-Markdown .\Architecture.md -Format Pdf -DiagramService none'
Write-Host '  Get-ChildItem *.md | Convert-Markdown -Format Pdf -Force'
Write-Host ''

# ------------------------------------------------------------- prerequisites
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    Write-Host "Microsoft Word $($word.Version) found." -ForegroundColor Green
}
catch {
    Write-Host 'Microsoft Word NOT found - conversion will fail. Word desktop is required for COM automation.' -ForegroundColor Red
}
finally {
    if ($word) {
        try { $word.Quit() } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    }
}

if (-not $SkipConnectivityTest) {
    # Mermaid diagrams render through a web service. Probe it now rather than
    # letting the first real conversion discover it is unreachable.
    try {
        $probe = Invoke-WebRequest -Uri 'https://mermaid.ink/img/Z3JhcGggVEQ7QS0tPkI7' `
            -Method Get -TimeoutSec 20 -MaximumRedirection 2 -ErrorAction Stop
        if ($probe.StatusCode -eq 200) {
            Write-Host 'mermaid.ink reachable - diagrams will render.' -ForegroundColor Green
        }
    }
    catch {
        Write-Host 'mermaid.ink NOT reachable from this machine.' -ForegroundColor Yellow
        Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkYellow
        Write-Host '  Diagrams will keep their source block. Use -DiagramService none for a clean' -ForegroundColor Yellow
        Write-Host '  placeholder, or -DiagramServiceUrl for a self-hosted Kroki.' -ForegroundColor Yellow
    }
}

Write-Host ''
Write-Host 'PRIVACY: mermaid diagram source is sent to the rendering service.' -ForegroundColor DarkYellow
Write-Host 'For internal architecture documents use -DiagramService none, or point' -ForegroundColor DarkYellow
Write-Host '-DiagramServiceUrl at a self-hosted Kroki instance.' -ForegroundColor DarkYellow

# --------------------------------------------------------------- style file
# The module ships neutral defaults on purpose. Brand colours, fonts and margins
# belong in a style file on this machine, not in the module source.
Import-Module $moduleName -Force

$profileStyle = Join-Path $HOME '.convertmarkdown\style.json'

# -------------------------------------------------------------------- themes
# theming/ is gitignored: a theme carries brand colours and a logo, which have no
# place in a public repo. Copy whatever is there to the per-user theme folder, so
# the theme works from any directory however the module was installed.
$themeSource = Join-Path $sourceDir 'theming'
$themeDest = Join-Path $HOME '.convertmarkdown\themes'
if (Test-Path -LiteralPath $themeSource) {
    foreach ($t in Get-ChildItem -LiteralPath $themeSource -Directory) {
        if (-not (Test-Path -LiteralPath (Join-Path $t.FullName 'theme.json'))) { continue }
        $target = Join-Path $themeDest $t.Name
        if ($PSCmdlet.ShouldProcess($target, "Install theme '$($t.Name)'")) {
            New-Item -ItemType Directory -Path $target -Force | Out-Null
            Copy-Item -Path (Join-Path $t.FullName '*') -Destination $target -Recurse -Force
            Write-Host "Theme   $($t.Name) -> $target" -ForegroundColor Green
        }
    }
}

if ($DefaultTheme) {
    if ($PSCmdlet.ShouldProcess($profileStyle, "Make '$DefaultTheme' the default theme")) {
        if (Test-Path -LiteralPath $profileStyle) {
            Copy-Item -LiteralPath $profileStyle -Destination "$profileStyle.bak" -Force
            Write-Host "Kept the previous style file as $profileStyle.bak" -ForegroundColor DarkGray
        }
        $created = New-ConvertMarkdownStyle -Path $profileStyle -Name $DefaultTheme -Theme $DefaultTheme -Force
        Write-Host "Default theme: $DefaultTheme ($($created.FullName))" -ForegroundColor Green
    }
    # The theme choice is the style file; do not offer to create another.
    $StylePath = 'none'
}

$chosen = $null
if ($StylePath) {
    if ($StylePath -ne 'none') { $chosen = $StylePath }
}
elseif (Test-Path -LiteralPath $profileStyle) {
    Write-Host ''
    Write-Host "Existing style file: $profileStyle" -ForegroundColor Green
}
else {
    Write-Host ''
    Write-Host '--- Style ---' -ForegroundColor Cyan
    Write-Host 'The module ships neutral defaults. A style file lets you set your own'
    Write-Host 'colours, fonts and margins without editing the module.'
    Write-Host ''
    Write-Host "  1. Create one for this user   ($profileStyle)"
    Write-Host '     Applies to every document you convert, in any repo.'
    Write-Host '  2. Choose a different path'
    Write-Host '  3. Skip - use built-in defaults'
    Write-Host ''
    $answer = Read-Host 'Choice [1]'
    if ([string]::IsNullOrWhiteSpace($answer)) { $answer = '1' }

    switch ($answer) {
        '1' { $chosen = $profileStyle }
        '2' {
            $custom = Read-Host 'Path for the style file (.json, or .yaml with powershell-yaml installed)'
            if (-not [string]::IsNullOrWhiteSpace($custom)) { $chosen = $custom }
        }
        default { Write-Host 'Skipped - using built-in defaults.' -ForegroundColor Yellow }
    }
}

if ($chosen) {
    try {
        $created = New-ConvertMarkdownStyle -Path $chosen -Name 'Custom' -Force:$Force
        Write-Host ''
        Write-Host "Style file created: $($created.FullName)" -ForegroundColor Green
        Write-Host 'Edit the colours, fonts and margin in there. Every key is present at its'
        Write-Host 'default value, so the file doubles as the reference.'
    }
    catch {
        Write-Host ''
        Write-Host "Could not create the style file: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host 'Create it later with:  New-ConvertMarkdownStyle' -ForegroundColor Yellow
    }
}

Write-Host ''
Write-Host 'Style resolution order (first match wins):' -ForegroundColor DarkGray
Write-Host '  1. -StyleSheet <path>' -ForegroundColor DarkGray
Write-Host '  2. $env:CONVERTMARKDOWN_STYLE' -ForegroundColor DarkGray
Write-Host '  3. .convertmarkdown.json beside the document' -ForegroundColor DarkGray
Write-Host '  4. ~/.convertmarkdown/style.json' -ForegroundColor DarkGray
Write-Host '  5. built-in defaults' -ForegroundColor DarkGray
Write-Host 'A theme (-Theme, the style file''s "theme" key, or $env:CONVERTMARKDOWN_THEME)' -ForegroundColor DarkGray
Write-Host 'sits under the style file: its keys override the defaults, the style file''s override it.' -ForegroundColor DarkGray
