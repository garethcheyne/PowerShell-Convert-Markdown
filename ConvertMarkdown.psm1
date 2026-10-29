<#
    ConvertMarkdown
    Converts Markdown into a styled Word .docx and/or PDF.

    Lineage: began as Convert-MarkdownToWord.ps1, which was copied into seven repos
    and drifted. This module is the consolidation of that script at v1.1.1
    (2026-08-21, the newest of the copies) plus universal output format, 10mm
    margins and installable packaging.

    Versioned yyyy.MM.dd. Keep $script:Version, ModuleVersion in the manifest and
    CHANGELOG.md in step when bumping.
#>

Set-StrictMode -Version Latest

$script:Version = '2026.10.29'

#region ---------------------------------------------------------- diagrams

function ConvertTo-Base64Url {
    param([Parameter(Mandatory)][byte[]] $Bytes)
    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Invoke-MermaidRender {
    <#
    .SYNOPSIS
        Renders one mermaid diagram to PNG bytes via a rendering web service.
    #>
    param(
        [Parameter(Mandatory)][string] $Source,
        [Parameter(Mandatory)][string] $Service,
        [string] $ServiceUrl,
        [string] $Theme = 'default',
        [hashtable] $ThemeVariables,
        [int] $TimeoutSec = 90
    )

    $utf8 = [Text.Encoding]::UTF8.GetBytes($Source)

    # A diagram carrying its own %%{init: ...}%% directive has been styled
    # deliberately. Ours would be overridden by it anyway, so do not send it and
    # let the author's intent stand.
    $selfStyled = $Source -match '%%\{\s*init\s*:'

    if ($Service -eq 'kroki') {
        # POST, so diagram size is not capped by URL length.
        $base = if ($ServiceUrl) { $ServiceUrl.TrimEnd('/') } else { 'https://kroki.io' }
        $body = @{ diagram_source = $Source }
        if (-not $selfStyled) { $body.diagram_options = @{ theme = $Theme } }
        $resp = Invoke-WebRequest -Uri "$base/mermaid/png" -Method Post `
            -Body ($body | ConvertTo-Json -Compress -Depth 6) -ContentType 'application/json' `
            -TimeoutSec $TimeoutSec -MaximumRedirection 2 -ErrorAction Stop
    }
    else {
        # mermaid.ink carries the diagram in the URL path. Prefer the "pako:" form
        # (zlib-deflated JSON state, as mermaid.live emits) - it roughly halves the
        # URL, which keeps large diagrams under the server's URL limit.
        $base = if ($ServiceUrl) { $ServiceUrl.TrimEnd('/') } else { 'https://mermaid.ink' }
        $encoded = $null
        try {
            $cfg = @{ theme = $Theme }
            if (-not $selfStyled -and $ThemeVariables -and $ThemeVariables.Count -gt 0) {
                $cfg.themeVariables = $ThemeVariables
            }
            $json = @{ code = $Source; mermaid = $cfg } | ConvertTo-Json -Compress -Depth 6
            $jb = [Text.Encoding]::UTF8.GetBytes($json)
            $ms = [IO.MemoryStream]::new()
            $zs = [IO.Compression.ZLibStream]::new($ms, [IO.Compression.CompressionLevel]::SmallestSize)
            $zs.Write($jb, 0, $jb.Length)
            $zs.Dispose()
            $encoded = 'pako:' + (ConvertTo-Base64Url -Bytes $ms.ToArray())
        }
        catch {
            $encoded = ConvertTo-Base64Url -Bytes $utf8       # plain base64url also accepted
        }

        $resp = Invoke-WebRequest -Uri "$base/img/$encoded`?type=png&bgColor=FFFFFF" -Method Get `
            -TimeoutSec $TimeoutSec -MaximumRedirection 2 -ErrorAction Stop
    }

    $bytes = $resp.Content
    if ($bytes -isnot [byte[]]) { $bytes = [Text.Encoding]::UTF8.GetBytes([string]$bytes) }
    if ($bytes.Length -lt 8 -or $bytes[0] -ne 0x89 -or $bytes[1] -ne 0x50) {
        throw "response was not a PNG ($($bytes.Length) bytes)"
    }
    return $bytes
}

#endregion

#region ------------------------------------------------------------- style

function Get-DefaultStyle {
    <#
    .SYNOPSIS
        The built-in style. Deliberately neutral - brand colours belong in a
        style file, not in the module.
    #>
    # Ordered so a generated style file reads in the same order as the example
    # in the repo, rather than however the hashtable happened to enumerate.
    [ordered]@{
        name     = 'Default'
        page     = [ordered]@{
            marginMm    = 10
            orientation = 'Portrait'
        }
        fonts    = [ordered]@{
            body       = 'Segoe UI'
            heading    = 'Segoe UI Semibold'
            # Title page headline. Empty means the heading font.
            title      = ''
            mono       = 'Consolas'
            bodySizePt = 10.5
        }
        colours  = [ordered]@{
            accent     = '#1F3864'
            accentSoft = '#2E74B5'
            # H1/H2 colour. Empty means accent, which is how it has always been.
            heading    = ''
            muted      = '#595959'
            rule       = '#C8102E'
            border     = '#B4C6E7'
            band       = '#F2F5FA'
            text       = '#1F1F1F'
        }
        tables   = [ordered]@{
            # 'filled': accent header row. 'rule': small uppercase labels over a rule.
            headerStyle = 'filled'
        }
        cover    = [ordered]@{
            enabled     = $false
            # Image file, relative to the style or theme file that names it.
            logo        = ''
            logoWidthMm = 55
            kicker      = ''
            # Panel colour. Empty means colours.band.
            band        = ''
            # Frontmatter keys to show along the bottom, in order. Empty shows them all.
            fields      = @()
        }
        diagrams = [ordered]@{
            service              = 'mermaid.ink'
            serviceUrl           = ''
            maxWidthPx           = 640
            theme                = 'base'
            matchDocumentPalette = $true
        }
        images   = [ordered]@{
            maxWidthPx  = 640
            allowRemote = $true
        }
    }
}

function Resolve-MarkdownImage {
    <#
    .SYNOPSIS
        Turns a Markdown image src into a local file path Word can embed.

    .DESCRIPTION
        Handles a path relative to the document, an absolute path, a file:// URI,
        and - when allowed - an http(s) URL, which is downloaded to a temp file.
        Returns $null when the image cannot be resolved; the caller warns and
        leaves a visible placeholder rather than failing the conversion.
    #>
    param(
        [Parameter(Mandatory)][string] $Source,
        [Parameter(Mandatory)][string] $BaseDir,
        [bool] $AllowRemote = $true
    )

    # Markdown allows a title after the path: ![alt](cat.png "Tiddles")
    $src = $Source.Trim()
    if ($src -match '^(?<p>\S+)\s+.*$') { $src = $Matches['p'] }
    $src = $src.Trim('<', '>')
    if (-not $src) { return $null }

    if ($src -match '^(?i)https?://') {
        if (-not $AllowRemote) {
            Write-Warning "Remote image skipped (images.allowRemote is false): $src"
            return $null
        }
        try {
            $ext = [IO.Path]::GetExtension(([Uri]$src).AbsolutePath)
            if (-not $ext -or $ext.Length -gt 5) { $ext = '.png' }
            $tmp = Join-Path ([IO.Path]::GetTempPath()) ("md2img_{0}{1}" -f [Guid]::NewGuid().ToString('N'), $ext)
            Invoke-WebRequest -Uri $src -OutFile $tmp -UseBasicParsing -ErrorAction Stop | Out-Null
            $script:tempFiles.Add($tmp)
            return $tmp
        }
        catch {
            Write-Warning "Could not download image ${src}: $($_.Exception.Message)"
            return $null
        }
    }

    if ($src -match '^(?i)file:///') {
        try { $src = ([Uri]$src).LocalPath } catch { return $null }
    }

    # Editors percent-encode spaces and unicode in the path as a matter of course.
    if ($src -match '%[0-9A-Fa-f]{2}') {
        try { $src = [Uri]::UnescapeDataString($src) } catch { }
    }

    $full = if ([IO.Path]::IsPathRooted($src)) { $src } else { Join-Path $BaseDir $src }
    if (Test-Path -LiteralPath $full -PathType Leaf) {
        return (Resolve-Path -LiteralPath $full).Path
    }
    return $null
}

function Get-MermaidThemeVariables {
    <#
    .SYNOPSIS
        Maps the document palette onto mermaid theme variables.

    .DESCRIPTION
        Without this a diagram arrives in mauve-on-beige mermaid defaults and
        looks pasted in from another document. Mermaid only honours these under
        the 'base' theme; the named themes ignore them entirely.
    #>
    param([Parameter(Mandatory)] $Style)

    $c = $Style.colours
    @{
        # Node fill, border and label.
        primaryColor        = $c.band
        primaryBorderColor  = $c.accent
        primaryTextColor    = $c.text
        # Edges and their labels.
        lineColor           = $c.accentSoft
        textColor           = $c.text
        # Secondary/tertiary nodes - clusters, alternate branches.
        secondaryColor      = '#FFFFFF'
        secondaryBorderColor = $c.border
        secondaryTextColor  = $c.text
        tertiaryColor       = '#FFFFFF'
        tertiaryBorderColor = $c.border
        tertiaryTextColor   = $c.text
        # State diagram specifics.
        clusterBkg          = '#FFFFFF'
        clusterBorder       = $c.border
        # Match the document, so diagram labels are not visibly a different face.
        fontFamily          = "'$($Style.fonts.body)', sans-serif"
        background          = '#FFFFFF'
    }
}

function Import-StyleFile {
    <#
    .SYNOPSIS
        Reads a style file. JSON is native; YAML needs the powershell-yaml module.
    #>
    param([Parameter(Mandatory)][string] $LiteralPath)

    $ext = [IO.Path]::GetExtension($LiteralPath).ToLowerInvariant()
    $raw = Get-Content -LiteralPath $LiteralPath -Raw -Encoding UTF8

    switch ($ext) {
        { $_ -in '.yaml', '.yml' } {
            if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
                throw "YAML style files need the powershell-yaml module: Install-Module powershell-yaml -Scope CurrentUser`nOr use a .json style file, which needs nothing extra."
            }
            Import-Module powershell-yaml -ErrorAction Stop
            return ConvertFrom-Yaml $raw
        }
        '.psd1' { return Import-PowerShellDataFile -LiteralPath $LiteralPath }
        default { return $raw | ConvertFrom-Json -AsHashtable }
    }
}

function Get-ThemeSearchPath {
    <#
    .SYNOPSIS
        Folders searched for a theme by name, in order.

    .DESCRIPTION
        A theme is a folder holding theme.json and any files it names (a logo).
        Themes live outside the module on purpose: brand assets are private, the
        module is public. Install.ps1 copies the repo's gitignored theming/ folder
        to ~/.convertmarkdown/themes, which is what makes a theme machine-wide.
    #>
    $dirs = [Collections.Generic.List[string]]::new()
    if ($env:CONVERTMARKDOWN_THEMES) {
        foreach ($d in ($env:CONVERTMARKDOWN_THEMES -split [IO.Path]::PathSeparator)) { if ($d) { $dirs.Add($d) } }
    }
    $dirs.Add((Join-Path $HOME '.convertmarkdown\themes'))
    # A -Link install points the module at the working copy, so its theming/ folder
    # is usable without installing the theme separately.
    $dirs.Add((Join-Path $PSScriptRoot 'theming'))
    $dirs
}

function Resolve-ThemeFile {
    <#
    .SYNOPSIS
        Turns a theme name or path into the path of its theme.json.
    #>
    param([Parameter(Mandatory)][string] $Theme)

    if (Test-Path -LiteralPath $Theme) {
        $item = Get-Item -LiteralPath $Theme
        if ($item.PSIsContainer) {
            $f = Join-Path $item.FullName 'theme.json'
            if (-not (Test-Path -LiteralPath $f)) { throw "Theme folder has no theme.json: $($item.FullName)" }
            return $f
        }
        return $item.FullName
    }

    $searched = Get-ThemeSearchPath
    foreach ($dir in $searched) {
        $f = Join-Path (Join-Path $dir $Theme) 'theme.json'
        if (Test-Path -LiteralPath $f) { return (Resolve-Path -LiteralPath $f).Path }
    }
    throw "Theme '$Theme' not found. Looked for $Theme\theme.json in:`n  " + ($searched -join "`n  ")
}

function Merge-StyleLayer {
    <#
    .SYNOPSIS
        Copies the keys a style or theme file sets over the style being built.

    .DESCRIPTION
        Relative file paths (cover.logo) are resolved against the folder of the file
        that names them, so a theme folder can be moved or installed anywhere.
    #>
    param(
        [Parameter(Mandatory)] $Style,
        [Parameter(Mandatory)] $Layer,
        [Parameter(Mandatory)][string] $BaseDir
    )

    foreach ($section in @('page', 'fonts', 'colours', 'tables', 'cover', 'diagrams', 'images')) {
        if ($Layer.Contains($section) -and $Layer[$section]) {
            foreach ($k in $Layer[$section].Keys) {
                $Style[$section][$k] = $Layer[$section][$k]
            }
        }
    }
    # US spelling accepted, since half the world writes it that way.
    if ($Layer.Contains('colors') -and $Layer['colors']) {
        foreach ($k in $Layer['colors'].Keys) { $Style['colours'][$k] = $Layer['colors'][$k] }
    }
    if ($Layer.Contains('name')) { $Style['name'] = $Layer['name'] }

    if ($Layer.Contains('cover') -and $Layer['cover'] -and $Layer['cover'].Contains('logo')) {
        $logo = [string]$Layer['cover']['logo']
        if ($logo -and -not [IO.Path]::IsPathRooted($logo)) {
            $Style['cover']['logo'] = [IO.Path]::GetFullPath((Join-Path $BaseDir $logo))
        }
    }
}

function Resolve-Style {
    <#
    .SYNOPSIS
        Merges a theme and a style file over the built-in defaults.

    .DESCRIPTION
        The style file is found in this order, first match wins:
          1. -StyleSheet
          2. $env:CONVERTMARKDOWN_STYLE
          3. .convertmarkdown.json beside the source document
          4. ~/.convertmarkdown/style.json
          5. none - built-in defaults

        A theme sits between the defaults and the style file. It comes from
        -Theme, else the style file's "theme" key, else $env:CONVERTMARKDOWN_THEME.
        So ~/.convertmarkdown/style.json holding just { "theme": "acme" } makes that
        theme the default for every document on the machine, and any other key in
        the style file still overrides the theme.

        A style or theme file only needs the keys it changes; everything else falls back.
    #>
    param(
        [string] $StyleSheet,
        [string] $SourceDir,
        [string] $Theme
    )

    $style = Get-DefaultStyle
    $file = $null

    if ($StyleSheet) {
        if (-not (Test-Path -LiteralPath $StyleSheet)) { throw "Style file not found: $StyleSheet" }
        $file = (Resolve-Path -LiteralPath $StyleSheet).Path
    }
    elseif ($env:CONVERTMARKDOWN_STYLE -and (Test-Path -LiteralPath $env:CONVERTMARKDOWN_STYLE)) {
        $file = $env:CONVERTMARKDOWN_STYLE
    }
    elseif ($SourceDir -and (Test-Path -LiteralPath (Join-Path $SourceDir '.convertmarkdown.json'))) {
        $file = Join-Path $SourceDir '.convertmarkdown.json'
    }
    else {
        $profileStyle = Join-Path $HOME '.convertmarkdown\style.json'
        if (Test-Path -LiteralPath $profileStyle) { $file = $profileStyle }
    }

    $loaded = if ($file) { Import-StyleFile -LiteralPath $file } else { $null }

    $themeRef = if ($Theme) { $Theme }
    elseif ($loaded -and $loaded.Contains('theme') -and $loaded['theme']) { [string]$loaded['theme'] }
    elseif ($env:CONVERTMARKDOWN_THEME) { $env:CONVERTMARKDOWN_THEME }
    else { $null }

    if ($themeRef) {
        $themeFile = Resolve-ThemeFile -Theme $themeRef
        Merge-StyleLayer -Style $style -Layer (Import-StyleFile -LiteralPath $themeFile) -BaseDir (Split-Path -Parent $themeFile)
        $style['themeFile'] = $themeFile
    }

    if ($loaded) {
        Merge-StyleLayer -Style $style -Layer $loaded -BaseDir (Split-Path -Parent $file)
        $style['sourceFile'] = $file
    }

    return $style
}

function New-ConvertMarkdownStyle {
    <#
    .SYNOPSIS
        Writes a starter style file you can edit.

    .DESCRIPTION
        The generated file contains every setting at its default value, so it
        doubles as the reference for what can be changed. Delete any key you do
        not want to override.

    .PARAMETER Path
        Where to write it. Defaults to ~/.convertmarkdown/style.json, which
        Convert-Markdown picks up automatically for every document.

    .PARAMETER Name
        A label for the style, shown in verbose output.

    .PARAMETER Theme
        Write a style file that just selects this theme (by name or path), instead
        of one listing every default. Every default in the file would override the
        theme, so a theme-based style file should carry only the keys that differ.

    .PARAMETER Force
        Overwrite an existing file.

    .EXAMPLE
        New-ConvertMarkdownStyle

    .EXAMPLE
        # Per-repo style, picked up for documents in that folder
        New-ConvertMarkdownStyle -Path .\.convertmarkdown.json -Name 'Project docs'

    .EXAMPLE
        # Make an installed theme the default for every document on this machine
        New-ConvertMarkdownStyle -Theme acme -Force
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string] $Path,
        [string] $Name = 'Custom',
        [string] $Theme,
        [switch] $Force
    )

    if (-not $Path) { $Path = Join-Path $HOME '.convertmarkdown\style.json' }
    $Path = [IO.Path]::GetFullPath($Path)

    if ((Test-Path -LiteralPath $Path) -and -not $Force) {
        throw "Style file already exists: $Path`nUse -Force to overwrite."
    }

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    if (-not $PSCmdlet.ShouldProcess($Path, 'Create style file')) { return }

    $defaults = Get-DefaultStyle

    # Ordered, and with $schema first, so editors offer completion and hover help
    # on a generated file the same as on the example in the repo.
    $style = if ($Theme) {
        # Fail now, not at the first conversion, if the theme is not installed.
        Resolve-ThemeFile -Theme $Theme | Out-Null
        [ordered]@{
            '$schema' = 'https://raw.githubusercontent.com/garethcheyne/PowerShell-Convert-Markdown/main/style.schema.json'
            name      = $Name
            theme     = $Theme
        }
    }
    else {
        [ordered]@{
            '$schema' = 'https://raw.githubusercontent.com/garethcheyne/PowerShell-Convert-Markdown/main/style.schema.json'
            name      = $Name
            page      = $defaults.page
            fonts     = $defaults.fonts
            colours   = $defaults.colours
            tables    = $defaults.tables
            cover     = $defaults.cover
            diagrams  = $defaults.diagrams
            images    = $defaults.images
        }
    }

    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -in '.yaml', '.yml') {
        if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
            throw "YAML output needs the powershell-yaml module: Install-Module powershell-yaml -Scope CurrentUser`nOr use a .json path, which needs nothing extra."
        }
        Import-Module powershell-yaml -ErrorAction Stop
        $style.Remove('$schema')          # not meaningful in YAML
        ConvertTo-Yaml $style | Set-Content -LiteralPath $Path -Encoding UTF8
    }
    else {
        $style | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Path -Encoding UTF8
    }

    Write-Verbose "Style file written to $Path"
    Get-Item -LiteralPath $Path
}

#endregion

#region ------------------------------------------- extended (NextDocs) markdown

# NextDocs extends Markdown with ::: fenced blocks, inline :kbd[]/:badge[]
# directives and role-gated !variant! sections. None of it survives
# ConvertFrom-Markdown intact: Markdig's custom-container extension turns
# ":::api GET /users/:id" into a bare <div> and throws the header away, reads the
# "---" separator inside a block as a setext underline, breaks a diff into two
# bullet lists, and its generic-attributes extension swallows ":badge[x]{warning}"
# into an HTML attribute on the surrounding paragraph. So every construct is
# resolved here, ahead of Markdig, and parked behind a token Markdig will not touch.
#
# The web renderer is interactive; paper is not. Each block is re-imagined as
# something that still reads when nothing can be clicked - see each builder for
# what it becomes.

$script:NdBlockTypes = @('tabs', 'details', 'steps', 'filetree', 'api', 'diff', 'gallery', 'sketch', 'release')

# GitHub's own alert palette, so a printed callout is recognisable to anyone who
# has seen the same document on the web.
$script:NdCallout = @{
    note      = @{ Colour = '#0969DA'; Tint = '#DDF4FF'; Glyph = '&#9432;'; Label = 'Note' }
    tip       = @{ Colour = '#1A7F37'; Tint = '#DAFBE1'; Glyph = '&#9733;'; Label = 'Tip' }
    important = @{ Colour = '#8250DF'; Tint = '#FBEFFF'; Glyph = '&#10070;'; Label = 'Important' }
    warning   = @{ Colour = '#9A6700'; Tint = '#FFF8C5'; Glyph = '&#9888;'; Label = 'Warning' }
    caution   = @{ Colour = '#CF222E'; Tint = '#FFEBE9'; Glyph = '&#10006;'; Label = 'Caution' }
}

$script:NdMethodColour = @{
    GET    = '#0969DA'; POST = '#1A7F37'; PUT = '#9A6700'; PATCH = '#8250DF'
    DELETE = '#CF222E'; HEAD = '#57606A'; OPTIONS = '#57606A'
}

# The web renderer infers a badge colour from the word when no {variant} is given.
$script:NdBadgeWord = @{
    stable       = 'success'; released = 'success'; production = 'success'
    deprecated   = 'error'; removed = 'error'; breaking = 'error'
    beta         = 'warning'; preview = 'warning'; alpha = 'warning'
    new          = 'info'; updated = 'info'; added = 'info'
    experimental = 'purple'; internal = 'purple'; draft = 'purple'
}

$script:NdBadgeColour = @{
    default = @{ Fg = '#374151'; Bg = '#F1F2F4' }
    success = @{ Fg = '#116329'; Bg = '#DAFBE1' }
    warning = @{ Fg = '#7D4E00'; Bg = '#FFF8C5' }
    error   = @{ Fg = '#A40E26'; Bg = '#FFEBE9' }
    info    = @{ Fg = '#0A3069'; Bg = '#DDF4FF' }
    purple  = @{ Fg = '#5B21B6'; Bg = '#FBEFFF' }
}

function ConvertTo-NdHtml {
    # Inner markdown -> HTML. NextDocs blocks never nest, so recursing into Markdig
    # here cannot loop, and it keeps lists, tables and emphasis working inside a block.
    param([AllowEmptyString()][string]$Markdown)
    if ([string]::IsNullOrWhiteSpace($Markdown)) { return '' }
    # Markdig closes a fenced block as "...\n</code></pre>". Word turns that last
    # newline into an empty paragraph, which inside a bordered block cell shows up
    # as a stray blank row under every code sample.
    (ConvertFrom-Markdown -InputObject $Markdown).Html -replace '\r?\n</code></pre>', '</code></pre>'
}

function ConvertTo-NdInlineHtml {
    # Same, for something that must stay on one line - a step title, a tab label.
    param([AllowEmptyString()][string]$Markdown)
    $h = (ConvertTo-NdHtml $Markdown).Trim()
    if ($h.StartsWith('<p>') -and $h.EndsWith('</p>') -and $h.IndexOf('<p>', 3) -lt 0) {
        return $h.Substring(3, $h.Length - 7)
    }
    $h
}

function Get-NdFenceRange {
    # Index ranges of fenced code. A document that *documents* this syntax must keep
    # its examples as text, so nothing inside a fence is ever rewritten.
    param([string]$Text)
    $ranges = [Collections.Generic.List[object]]::new()
    $fenceRe = '(?ms)^[ \t]*(?<f>`{3,}|~{3,})[^\n]*\n.*?^[ \t]*\k<f>[ \t]*$'
    foreach ($m in [regex]::Matches($Text, $fenceRe)) {
        $ranges.Add([pscustomobject]@{ Start = $m.Index; End = $m.Index + $m.Length })
    }
    , $ranges
}

function Test-NdInFence {
    param($Ranges, [int]$Index)
    foreach ($r in $Ranges) { if ($Index -ge $r.Start -and $Index -lt $r.End) { return $true } }
    $false
}

function Get-NdStepGlyph {
    # U+2776..U+277F are the filled circled digits 1-10. Tinted with the accent they
    # print as a solid disc with the number knocked out of it, which is what the web
    # renderer draws. U+2460.. are the outline circles, used past ten. Both ranges
    # live in Segoe UI Symbol, which this module already falls back to for the
    # task-list checkboxes.
    param([int]$Number)
    if ($Number -ge 1 -and $Number -le 10) { '&#{0};' -f (10101 + $Number) }
    elseif ($Number -le 20) { '&#{0};' -f (9311 + $Number) }
    else { "$Number." }
}

function Split-NdHeader {
    # Several blocks put metadata above a "---" rule and content below it. Markdig
    # would read that rule as a setext underline, so it is split off by hand.
    param([string]$Body)
    $m = [regex]::Match($Body, '(?m)^[ \t]*---[ \t]*$')
    if ($m.Success) {
        [pscustomobject]@{
            Head = $Body.Substring(0, $m.Index)
            Rest = $Body.Substring($m.Index + $m.Length)
        }
    }
    else {
        [pscustomobject]@{ Head = ''; Rest = $Body }
    }
}

function New-NdCallout {
    # One cell with a thick left border, not a narrow spacer cell beside it: Word
    # ignores a sub-point column width and stretches the spacer into a slab.
    param([string]$Type, [string]$Body)
    $c = $script:NdCallout[$Type]
    '<table class="nd callout" summary="nd">' +
    ('<tr><td class="callout-body" style="background:{0};border-left:4pt solid {1};">' -f $c.Tint, $c.Colour) +
    ('<p class="callout-title" style="color:{0};"><span class="sym">{1}</span>&#160;{2}</p>' -f $c.Colour, $c.Glyph, $c.Label) +
    (ConvertTo-NdHtml $Body) + '</td></tr></table>'
}

function New-NdSteps {
    # A numbered disc in the gutter, the step title in bold beside it and the body
    # under that - the print equivalent of the web renderer's connected step list.
    param([string]$Body)
    $parts = [regex]::Split($Body, '(?m)^[ \t]*#{2,4}[ \t]+')
    $rows = [Collections.Generic.List[string]]::new()
    $lead = ''
    $n = 0
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ($i -eq 0) {
            if (-not [string]::IsNullOrWhiteSpace($parts[0])) { $lead = ConvertTo-NdHtml $parts[0] }
            continue
        }
        $n++
        $split = $parts[$i] -split "`n", 2
        $title = ConvertTo-NdInlineHtml $split[0].Trim()
        $desc = if ($split.Count -gt 1) { ConvertTo-NdHtml $split[1] } else { '' }
        # The whole -f expression needs its own parentheses: inside .Add(..) a bare
        # comma is read as the next method argument, not as the next format value.
        $rows.Add((('<tr><td class="num"><span class="step-num">{0}</span></td>' +
                    '<td class="step"><p class="step-title">{1}</p>{2}</td></tr>') -f (Get-NdStepGlyph $n), $title, $desc))
    }
    # No headings inside: nothing to number, so leave the content as written.
    if ($n -eq 0) { return (ConvertTo-NdHtml $Body) }
    $lead + '<table class="nd steps" summary="nd">' + ($rows -join '') + '</table>'
}

function New-NdTabs {
    # Nothing can be clicked on paper. Every panel is printed, each behind its own
    # label chip, so the reader keeps all the alternatives instead of losing four of
    # five to whichever tab happened to be open.
    param([string]$Body)
    $parts = [regex]::Split($Body, '(?m)^[ \t]*@tab[ \t]*')
    $rows = [Collections.Generic.List[string]]::new()
    for ($i = 1; $i -lt $parts.Count; $i++) {
        $split = $parts[$i] -split "`n", 2
        $label = ConvertTo-NdInlineHtml $split[0].Trim()
        $panel = if ($split.Count -gt 1) { ConvertTo-NdHtml $split[1] } else { '' }
        $rows.Add(('<tr><td class="tab-label">{0}</td></tr><tr><td class="tab-body">{1}</td></tr>' -f $label, $panel))
    }
    if ($rows.Count -eq 0) { return (ConvertTo-NdHtml $Body) }
    '<table class="nd tabs" summary="nd">' + ($rows -join '') + '</table>'
}

function New-NdDetails {
    # A disclosure triangle has nothing to disclose in print, so the panel is shown
    # open with its summary kept as the heading it already is on the web.
    param([string]$Summary, [string]$Body)
    $label = if ($Summary.Trim()) { ConvertTo-NdInlineHtml $Summary.Trim() } else { 'Details' }
    '<table class="nd details" summary="nd">' +
    ('<tr><td class="details-head"><span class="sym">&#9662;</span>&#160;{0}</td></tr>' -f $label) +
    ('<tr><td class="details-body">{0}</td></tr>' -f (ConvertTo-NdHtml $Body)) +
    '</table>'
}

function New-NdFiletree {
    # The web draws indent guides. Print gets the box-drawing tree everyone already
    # reads in terminal output, which needs no colour and survives photocopying.
    param([string]$Body)
    $items = [Collections.Generic.List[object]]::new()
    foreach ($line in ($Body -split "`n")) {
        if ($line -notmatch '^(?<pad>[ \t]*)[-*+][ \t]+(?<name>.+?)[ \t]*$') { continue }
        $pad = $Matches['pad'] -replace "`t", '  '
        $items.Add([pscustomobject]@{
                Depth = [int][Math]::Floor($pad.Length / 2)
                Name  = $Matches['name']
            })
    }
    if ($items.Count -eq 0) { return (ConvertTo-NdHtml $Body) }

    $out = [Collections.Generic.List[string]]::new()
    $cont = @{}   # depth -> a sibling still follows, so keep drawing that trunk
    for ($i = 0; $i -lt $items.Count; $i++) {
        $d = $items[$i].Depth
        $last = $true
        for ($j = $i + 1; $j -lt $items.Count; $j++) {
            if ($items[$j].Depth -lt $d) { break }
            if ($items[$j].Depth -eq $d) { $last = $false; break }
        }
        $cont[$d] = -not $last

        $prefix = ''
        for ($a = 0; $a -lt $d; $a++) {
            $prefix += $(if ($cont.ContainsKey($a) -and $cont[$a]) { '&#9474;&#160;&#160;&#160;' } else { '&#160;&#160;&#160;&#160;' })
        }
        $prefix += $(if ($last) { '&#9492;&#9472;&#9472;&#160;' } else { '&#9500;&#9472;&#9472;&#160;' })

        $name = [Net.WebUtility]::HtmlEncode($items[$i].Name)
        # A trailing slash is how NextDocs marks a folder; keep it and weight it.
        if ($items[$i].Name.EndsWith('/')) { $name = '<strong>' + $name + '</strong>' }
        $out.Add($prefix + $name)
    }
    '<pre class="filetree">' + ($out -join "`n") + '</pre>'
}

function New-NdApi {
    param([string]$Header, [string]$Body)
    $method = 'GET'
    $path = ''
    if ($Header.Trim() -match '^(?<m>[A-Za-z]+)[ \t]+(?<p>.+)$') {
        $method = $Matches['m'].ToUpperInvariant()
        $path = $Matches['p'].Trim()
    }
    else {
        $path = $Header.Trim()
    }
    $colour = if ($script:NdMethodColour.ContainsKey($method)) { $script:NdMethodColour[$method] } else { '#57606A' }
    $split = Split-NdHeader $Body

    # Method and path share one cell. As two cells the chip is a table column, and
    # Word autofits each table on its own content - so consecutive blocks came out
    # with chips of visibly different widths. As a shaded span it is the same size
    # every time, whatever sits underneath it.
    '<table class="nd api" summary="nd">' +
    ('<tr><td class="api-head"><span class="api-method" style="background:{0};">&#160;{1}&#160;</span>' -f
        $colour, [Net.WebUtility]::HtmlEncode($method)) +
    ('&#160;&#160;<span class="api-path">{0}</span></td></tr>' -f [Net.WebUtility]::HtmlEncode($path)) +
    ('<tr><td class="api-body">{0}{1}</td></tr>' -f (ConvertTo-NdHtml $split.Head), (ConvertTo-NdHtml $split.Rest)) +
    '</table>'
}

function New-NdDiff {
    # Tinted rows carry the colour; the +/- stays in the text so the block still
    # reads when the page is printed in grey-scale or read aloud.
    param([string]$Header, [string]$Body)
    $split = Split-NdHeader $Body
    $caption = ''
    if ($split.Head -match '(?im)^[ \t]*filename:[ \t]*(?<f>.+?)[ \t]*$') { $caption = $Matches['f'] }
    $lang = $Header.Trim()

    $rows = [Collections.Generic.List[string]]::new()
    if ($caption -or $lang) {
        $bar = [Net.WebUtility]::HtmlEncode($caption)
        # Hard spaces, not padding on the span: Word drops the padding and the
        # filename runs straight into the language.
        if ($lang) { $bar += ('&#160;&#160;<span class="diff-lang">{0}</span>' -f [Net.WebUtility]::HtmlEncode($lang)) }
        $rows.Add('<tr><td class="diff-head">' + $bar + '</td></tr>')
    }
    foreach ($line in ($split.Rest -split "`n")) {
        if ([string]::IsNullOrWhiteSpace($line) -and $rows.Count -le 1) { continue }
        # The row tint has to be inline: Word flattens CSS to direct formatting and
        # drops descendant selectors, so "table.diff tr.add td" would never land.
        $tint = ''
        if ($line -match '^\+') { $tint = 'background:#E6FFEC;color:#0A5C1E;' }
        elseif ($line -match '^-') { $tint = 'background:#FFEBE9;color:#8B1A22;' }
        # Leading spaces collapse in Word's HTML import, so hard-space the indent.
        $text = [Net.WebUtility]::HtmlEncode($line.TrimEnd())
        $text = [regex]::Replace($text, '^ +', { param($m) '&#160;' * $m.Length })
        if (-not $text) { $text = '&#160;' }
        # Bare text in a cell inherits the body paragraph spacing, which opens a
        # visible gap between every line of the diff; the paragraph is explicit here
        # so its margin can be zeroed.
        $rows.Add('<tr><td class="diff-line" style="' + $tint + '"><p style="margin:0;">' + $text + '</p></td></tr>')
    }
    '<table class="nd diff" summary="nd">' + ($rows -join '') + '</table>'
}

function New-NdGallery {
    # The image pass upstream has already swapped every ![](..) for a token, so the
    # gallery only has to lay the tokens out; AddPicture fills them in after import.
    param([string]$Header, [string]$Body, [hashtable]$Alt)
    $cols = 3                                   # elastic: what fits an A4 text column
    $h = $Header.Trim()
    if ($h -match '^(?<c>\d+)[xX](?<r>\d+)$') { $cols = [int]$Matches['c'] }
    elseif ($h -match '^(?<c>\d+)$') { $cols = [int]$Matches['c'] }
    $cols = [Math]::Max(1, [Math]::Min(6, $cols))

    $cells = [Collections.Generic.List[string]]::new()
    foreach ($m in [regex]::Matches($Body, '@@MDIMAGE\d+@@|<span class="image-missing">.*?</span>')) {
        # Remember the column count against the token. The picture is not inserted
        # until Word has the document open, and by then the cell is still empty and
        # sized on nothing - so a picture fitted to the page width silently widens
        # its column and pushes the whole table off the right margin. The embed step
        # divides the text width by this instead.
        $script:ndGalleryCols[$m.Value] = $cols
        $cap = ''
        if ($Alt.ContainsKey($m.Value) -and $Alt[$m.Value]) {
            $cap = '<p class="gallery-cap">' + [Net.WebUtility]::HtmlEncode($Alt[$m.Value]) + '</p>'
        }
        $cells.Add('<td class="gallery-cell" width="' + [int](100 / $cols) + '%">' + $m.Value + $cap + '</td>')
    }
    if ($cells.Count -eq 0) { return (ConvertTo-NdHtml $Body) }

    $rows = [Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $cells.Count; $i += $cols) {
        $slice = @($cells[$i..([Math]::Min($i + $cols - 1, $cells.Count - 1))])
        # Pad the last row so the columns keep the same width all the way down.
        $pad = $cols - $slice.Count
        $rows.Add('<tr>' + ($slice -join '') + ('<td class="gallery-cell"></td>' * $pad) + '</tr>')
    }
    '<table class="nd gallery" summary="nd">' + ($rows -join '') + '</table>'
}

function New-NdSketch {
    # The drawing itself is a database row, not markdown - there is nothing in the
    # file to render, so say so rather than printing a bare UUID.
    param([string]$Header, [string]$Body)
    $cap = $Body.Trim()
    $note = '<strong>Sketch</strong> &#8212; drawn in NextDocs and not stored in the Markdown, so it cannot be printed.'
    if ($cap) { $note += ' ' + [Net.WebUtility]::HtmlEncode($cap) }
    '<p class="block-missing">' + $note + '</p>'
}

function New-NdRelease {
    param([string]$Body)
    $split = Split-NdHeader $Body
    $teams = ''
    $version = ''
    foreach ($line in ($split.Head -split "`n")) {
        if ($line -match '(?i)^[ \t]*teams:[ \t]*(?<v>.+?)[ \t]*$') { $teams = $Matches['v'] }
        elseif ($line -match '(?i)^[ \t]*version:[ \t]*(?<v>.+?)[ \t]*$') { $version = $Matches['v'] }
    }
    $bar = 'Release'
    if ($version) { $bar += ' ' + [Net.WebUtility]::HtmlEncode($version) }

    '<table class="nd release" summary="nd">' +
    ('<tr><td class="release-head">{0}</td></tr>' -f $bar) +
    $(if ($teams) { '<tr><td class="release-teams">' + [Net.WebUtility]::HtmlEncode($teams) + '</td></tr>' } else { '' }) +
    ('<tr><td class="release-body">{0}</td></tr>' -f (ConvertTo-NdHtml $split.Rest)) +
    '</table>'
}

function New-NdRestricted {
    # Role-gated content the website hides from most readers. It is printed, but it
    # is never printed silently - the banner says who it was written for, and
    # -ExcludeRestricted drops it before anything else looks at the file.
    param([string]$Role, [string]$Body)
    '<table class="nd restricted" summary="nd">' +
    ('<tr><td class="restricted-head"><span class="sym">&#9679;</span>&#160;Restricted &#8212; {0}</td></tr>' -f [Net.WebUtility]::HtmlEncode($Role.Trim())) +
    ('<tr><td class="restricted-body">{0}</td></tr>' -f (ConvertTo-NdHtml $Body)) +
    '</table>'
}

$script:NdFrontMatterLabel = [ordered]@{
    author         = 'Author'
    date           = 'Date'
    publishedat    = 'Published'
    published      = 'Published'
    category       = 'Category'
    tags           = 'Tags'
    draft          = 'Draft'
    restricted     = 'Restricted'
    restrictedroles = 'Restricted roles'
}

function Split-NdFrontMatter {
    # Every NextDocs page opens with YAML frontmatter. Markdig has no frontmatter
    # extension enabled here, so it reads the closing "---" as a setext underline and
    # welds the whole block into one enormous H2. NextDocs itself renders `title` and
    # `excerpt` as the page header and never repeats them in the body, so the same is
    # done here: title becomes the document H1 (which is also what the running header
    # and the .docx Title property are taken from), excerpt becomes the standfirst,
    # and everything else joins the key/value box this module already draws for
    # "**Key:** value" runs.
    #
    # -NoHeader is for a title page: the fields go on the cover instead, so none of
    # them is written into the body.
    param(
        [string]$Markdown,
        [switch]$NoHeader
    )

    $m = [regex]::Match($Markdown, '(?s)^---[ \t]*\n(?<yaml>.*?)\n---[ \t]*(?:\n|$)')
    if (-not $m.Success) {
        return [pscustomobject]@{ Markdown = $Markdown; Found = $false; Title = ''; Excerpt = ''; Fields = [ordered]@{} }
    }

    # Flat "key: value" only, plus "key:" followed by a "- item" list. That is the
    # whole of the NextDocs frontmatter contract; anything nested deeper is left out
    # rather than half-rendered.
    $fields = [ordered]@{}
    $listKey = ''
    foreach ($line in ($m.Groups['yaml'].Value -split "`n")) {
        if ($line -match '^[ \t]+-[ \t]*(?<v>.+?)[ \t]*$' -and $listKey) {
            $fields[$listKey] = @($fields[$listKey], $Matches['v']) -ne '' -join ', '
            continue
        }
        if ($line -notmatch '^(?<k>[A-Za-z_][A-Za-z0-9_-]*):[ \t]*(?<v>.*?)[ \t]*$') { continue }
        $key = $Matches['k']
        $val = $Matches['v'].Trim('"', "'")
        $listKey = if ($val) { '' } else { $key }
        # tags: [a, b] and tags: "a, b" both mean the same list.
        if ($val -match '^\[(?<i>.*)\]$') { $val = ($Matches['i'] -split ',' | ForEach-Object { $_.Trim().Trim('"', "'") }) -join ', ' }
        $fields[$key] = $val
    }

    $head = [Collections.Generic.List[string]]::new()
    $title = ''
    $excerpt = ''
    foreach ($k in @($fields.Keys)) {
        switch ($k.ToLowerInvariant()) {
            'title' { $title = $fields[$k] }
            'excerpt' { if (-not $excerpt) { $excerpt = $fields[$k] } }
            'description' { if (-not $excerpt) { $excerpt = $fields[$k] } }
        }
    }
    # Emitted as an encoded <h1> rather than "# $title": a frontmatter title is
    # metadata, not markdown, and NextDocs renders it as plain text. Left as markdown
    # a title containing markup would be interpreted rather than shown. The module's
    # own title extraction reads <h1> either way.
    if ($title) { $head.Add('<h1>' + [Net.WebUtility]::HtmlEncode($title) + '</h1>'); $head.Add('') }
    if ($excerpt) { $head.Add('<p class="excerpt">' + [Net.WebUtility]::HtmlEncode($excerpt) + '</p>'); $head.Add('') }

    $rows = [Collections.Generic.List[string]]::new()
    foreach ($k in @($fields.Keys)) {
        $lk = $k.ToLowerInvariant()
        if ($lk -in @('title', 'excerpt', 'description')) { continue }
        $v = $fields[$k]
        if (-not $v) { continue }
        # "draft: false" and "restricted: false" say nothing; only the true case is
        # worth a row on the page.
        if ($lk -in @('draft', 'restricted') -and $v -notmatch '^(?i)true$') { continue }
        $label = if ($script:NdFrontMatterLabel.Contains($lk)) { $script:NdFrontMatterLabel[$lk] } else { $k }
        # Same as the "**Key:** value" box below: the cells are raw HTML by the time
        # Markdig runs, so inline markup in a frontmatter value is converted here.
        $rows.Add(('<tr><td class="k">{0}</td><td class="v">{1}</td></tr>' -f
                [Net.WebUtility]::HtmlEncode($label), (ConvertTo-NdInlineHtml $v)))
    }
    if ($rows.Count -gt 0) {
        $head.Add('<table class="meta">' + ($rows -join '') + '</table>')
        $head.Add('')
    }
    if ($NoHeader) { $head.Clear() }

    [pscustomobject]@{
        Markdown = ($head -join "`n") + $Markdown.Substring($m.Length)
        Found    = $true
        Title    = $title
        Excerpt  = $excerpt
        Fields   = $fields
    }
}

function Test-NdRestrictedBalance {
    # An unterminated !variant! is a confidentiality problem, not a formatting one:
    # the strip below is regex-paired, so a missing (or misspelt) !endvariant! means
    # -ExcludeRestricted silently removes nothing and the section ships anyway.
    # Counted with code removed first, so a document that documents the syntax - this
    # module's own README among them - is not mistaken for an unbalanced one.
    param([string]$Markdown)
    $bare = [regex]::Replace($Markdown, '(?ms)^[ \t]*(?<f>`{3,}|~{3,})[^\n]*\n.*?^[ \t]*\k<f>[ \t]*$', '')
    $bare = [regex]::Replace($bare, '`[^`\n]+`', '')
    [pscustomobject]@{
        Open  = [regex]::Matches($bare, '(?m)^[ \t]*!variant!').Count
        Close = [regex]::Matches($bare, '(?m)^[ \t]*!endvariant!').Count
    }
}

function Remove-NdRestricted {
    # Runs before the diagram and image passes so a restricted section's mermaid
    # source is never sent to a rendering service it was not meant to reach.
    param([string]$Markdown)
    [regex]::Replace($Markdown, '(?ms)^[ \t]*!variant!\s*#\s*[^\r\n]+\r?\n.*?^[ \t]*!endvariant![ \t]*$', '')
}

function Convert-NdInline {
    # :kbd[..], :badge[..]{..} and :#fluentui name: become inline HTML, which
    # CommonMark passes straight through, so these need no token. Fenced and inline
    # code is skipped: a document showing the syntax must keep showing it.
    param([string]$Markdown)
    $script:ndKbd = 0
    $script:ndBadge = 0
    $inFence = $false
    $out = foreach ($line in ($Markdown -split "`n")) {
        if ($line -match '^[ \t]*(```|~~~)') { $inFence = -not $inFence; $line; continue }
        if ($inFence -or $line -notmatch ':kbd\[|:badge\[|:#fluentui') { $line; continue }

        $work = $line
        $work = [regex]::Replace($work, ':kbd\[(?<k>[^\]]+)\]', {
                param($m)
                $before = $work.Substring(0, $m.Index)
                if (([regex]::Matches($before, '`').Count % 2) -eq 1) { return $m.Value }
                $script:ndKbd++
                $keys = @($m.Groups['k'].Value -split '\+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                ($keys | ForEach-Object { '<span class="kbd">' + [Net.WebUtility]::HtmlEncode($_) + '</span>' }) -join '<span class="kbd-plus">+</span>'
            })

        $work = [regex]::Replace($work, ':badge\[(?<t>[^\]]+)\](?:\{(?<v>[^}]+)\})?', {
                param($m)
                $before = $work.Substring(0, $m.Index)
                if (([regex]::Matches($before, '`').Count % 2) -eq 1) { return $m.Value }
                $script:ndBadge++
                $text = $m.Groups['t'].Value
                $v = $m.Groups['v'].Value
                if (-not $v) {
                    $key = $text.Trim().ToLowerInvariant()
                    $v = if ($script:NdBadgeWord.ContainsKey($key)) { $script:NdBadgeWord[$key] } else { 'default' }
                }
                if (-not $script:NdBadgeColour.ContainsKey($v)) { $v = 'default' }
                $c = $script:NdBadgeColour[$v]
                '<span class="badge" style="background:{0};color:{1};">{2}</span>' -f $c.Bg, $c.Fg, [Net.WebUtility]::HtmlEncode($text)
            })

        # Fluent icons are an icon font that is not on the page. The name alone is
        # noise, so the directive is dropped. Lucide's ":name:" form is deliberately
        # left alone - that pattern is too close to ordinary text such as "10:30:00".
        $work = [regex]::Replace($work, ':#fluentui[ \t]+[A-Za-z0-9-]+:[ \t]?', '')
        $work
    }
    $out -join "`n"
}

function Expand-NdBlock {
    # Replaces every ::: block, callout and !variant! section with a token and
    # returns the finished HTML for each one keyed by that token. Markdig then sees
    # only the tokens, so it has nothing left to mangle.
    param(
        [string]$Markdown,
        [hashtable]$ImageAlt,
        [switch]$LabelRestricted
    )
    $stash = [ordered]@{}
    $counts = @{ Block = 0; Callout = 0; Restricted = 0 }
    $script:ndGalleryCols = @{}

    $emit = {
        param([string]$Html)
        $token = '@@NDBLOCK{0}@@' -f ($stash.Count + 1)
        $stash[$token] = $Html
        "`n$token`n"
    }

    # --- callouts: > [!TYPE] plus the blockquote lines under it
    $fences = Get-NdFenceRange $Markdown
    $calloutRe = '(?m)^[ \t]*>[ \t]*\[!(?<t>NOTE|TIP|IMPORTANT|WARNING|CAUTION)\][ \t]*$(?<rest>(?:\r?\n[ \t]*>[^\n]*)*)'
    $Markdown = [regex]::Replace($Markdown, $calloutRe, {
            param($m)
            if (Test-NdInFence $fences $m.Index) { return $m.Value }
            $body = (($m.Groups['rest'].Value -split "`n") | ForEach-Object { $_ -replace '^[ \t]*>[ \t]?', '' }) -join "`n"
            $counts.Callout++
            & $emit (New-NdCallout -Type $m.Groups['t'].Value.ToLowerInvariant() -Body $body)
        })

    # --- ::: blocks, spliced back in reverse so earlier indices stay valid
    $fences = Get-NdFenceRange $Markdown
    $blockRe = '(?ms)^[ \t]*:::[ \t]*(?<type>' + ($script:NdBlockTypes -join '|') +
    ')(?<hdr>[^\n]*)\n(?<body>.*?)\n[ \t]*:::[ \t]*$'
    $found = @([regex]::Matches($Markdown, $blockRe) | Where-Object { -not (Test-NdInFence $fences $_.Index) })
    [Array]::Reverse($found)

    foreach ($m in $found) {
        $type = $m.Groups['type'].Value
        $hdr = $m.Groups['hdr'].Value
        $body = $m.Groups['body'].Value
        $html = switch ($type) {
            'tabs' { New-NdTabs -Body $body }
            'details' { New-NdDetails -Summary $hdr -Body $body }
            'steps' { New-NdSteps -Body $body }
            'filetree' { New-NdFiletree -Body $body }
            'api' { New-NdApi -Header $hdr -Body $body }
            'diff' { New-NdDiff -Header $hdr -Body $body }
            'gallery' { New-NdGallery -Header $hdr -Body $body -Alt $ImageAlt }
            'sketch' { New-NdSketch -Header $hdr -Body $body }
            'release' { New-NdRelease -Body $body }
            default { ConvertTo-NdHtml $body }
        }
        $counts.Block++
        $Markdown = $Markdown.Remove($m.Index, $m.Length).Insert($m.Index, (& $emit $html))
    }

    # --- role-gated sections (only reached when they were not stripped outright)
    if ($LabelRestricted) {
        $fences = Get-NdFenceRange $Markdown
        $variantRe = '(?ms)^[ \t]*!variant!\s*#\s*(?<role>[^\r\n]+)\r?\n(?<body>.*?)^[ \t]*!endvariant![ \t]*$'
        $Markdown = [regex]::Replace($Markdown, $variantRe, {
                param($m)
                if (Test-NdInFence $fences $m.Index) { return $m.Value }
                $counts.Restricted++
                & $emit (New-NdRestricted -Role $m.Groups['role'].Value -Body $m.Groups['body'].Value)
            })
    }

    [pscustomobject]@{
        Markdown   = $Markdown
        Blocks     = $stash
        Count      = $counts.Block
        Callouts   = $counts.Callout
        Restricted = $counts.Restricted
    }
}

function Get-NdExtendedCss {
    # Word's HTML import is unreliable for anything laid out with floats, flexbox or
    # border-radius, but it is solid on tables, cell shading and borders. Every block
    # above is therefore a table.
    #
    # Every rule below is a single-element class selector. Word flattens CSS to
    # direct formatting as it imports and does not resolve descendant selectors, so
    # "table.nd td.tab-label" silently loses to the plain "td" rule further up the
    # sheet and every cell comes back with the generic table border around it. Each
    # cell class therefore restates its own border and padding, and anything whose
    # colour varies per row is written inline by the builder instead.
    param([hashtable]$Style)
    $accent = $Style.colours.accent
    $accentSoft = $Style.colours.accentSoft
    $border = $Style.colours.border
    $band = $Style.colours.band
    $text = $Style.colours.text
    $fontBody = $Style.fonts.body
    $fontMono = $Style.fonts.mono
    $sizeBody = $Style.fonts.bodySizePt

    $mono = "$fontMono, 'Cascadia Mono', 'DejaVu Sans Mono', 'Courier New', monospace"
    $sym = "'Segoe UI Symbol', 'Segoe UI', $fontBody, sans-serif"

    @"
table.nd { border-collapse: collapse; width: 100%; margin: 8pt 0 12pt 0;
     font-size: ${sizeBody}pt; page-break-inside: avoid; }
span.sym { font-family: $sym; }

p.excerpt { color: #595959; font-size: 11.5pt; margin: -6pt 0 12pt 0; }

td.callout-body { border: none; padding: 7pt 10pt 3pt 10pt; vertical-align: top; }
p.callout-title { font-weight: 700; margin: 0 0 3pt 0; }

td.num { border: none; padding: 1pt 7pt 9pt 0; vertical-align: top; width: 22pt; }
td.step { border: none; padding: 0 0 9pt 0; vertical-align: top; }
span.step-num { font-family: $sym; font-size: 15pt; color: $accent; }
p.step-title { font-weight: 700; color: $text; margin: 0 0 2pt 0; }

td.tab-label { background: $accentSoft; color: #FFFFFF; font-weight: 700; font-size: 8.5pt;
     border: none; padding: 2pt 10pt; vertical-align: top; }
td.tab-body { border: 0.5pt solid $border; border-top: none; padding: 7pt 10pt 3pt 10pt;
     vertical-align: top; }

td.details-head { background: $band; color: $accentSoft; font-weight: 700;
     border: 0.5pt solid $border; padding: 3pt 10pt; vertical-align: top; }
td.details-body { border: 0.5pt solid $border; border-top: none; padding: 7pt 10pt 3pt 10pt;
     vertical-align: top; }

pre.filetree { font-family: $mono; font-size: 9pt; line-height: 1.1; background: #FBFBFC;
     border: 0.5pt solid $border; border-left: 0.5pt solid $border;
     padding: 7pt 10pt; margin: 8pt 10pt 12pt 10pt; page-break-inside: avoid; }

td.api-head { background: $band; border: none; padding: 4pt 10pt; vertical-align: top; }
span.api-method { color: #FFFFFF; font-weight: 700; font-size: 8.5pt; }
span.api-path { font-family: $mono; font-size: 9.5pt; font-weight: 600; }
td.api-body { border: 0.5pt solid $border; border-top: none; padding: 7pt 10pt 3pt 10pt;
     vertical-align: top; }

td.diff-head { background: $band; font-family: $mono; font-size: 9pt; font-weight: 700;
     border: none; border-bottom: 0.5pt solid $border; padding: 3pt 10pt; vertical-align: top; }
span.diff-lang { color: #595959; font-weight: 400; font-size: 8pt; }
td.diff-line { font-family: $mono; font-size: 8.5pt; border: none;
     padding: 0.5pt 10pt; vertical-align: top; }

td.gallery-cell { border: none; text-align: center; padding: 4pt; vertical-align: top; }
p.gallery-cap { font-size: 8pt; color: #595959; margin: 2pt 0 0 0; }

td.release-head { background: $accent; color: #FFFFFF; font-weight: 700; border: none;
     padding: 4pt 10pt; vertical-align: top; }
td.release-teams { background: $band; color: $accentSoft; font-size: 8.5pt; font-weight: 600;
     border: none; padding: 2pt 10pt; vertical-align: top; }
td.release-body { border: 0.5pt solid $accent; border-top: none; padding: 7pt 10pt 3pt 10pt;
     vertical-align: top; }

td.restricted-head { background: $band; color: $accentSoft; font-weight: 700; font-size: 8.5pt;
     border: 0.75pt dashed $accentSoft; padding: 3pt 10pt; vertical-align: top; }
td.restricted-body { border: 0.75pt dashed $accentSoft; border-top: none;
     padding: 7pt 10pt 3pt 10pt; vertical-align: top; }

span.kbd { font-family: $mono; font-size: 8.5pt; color: $text; background: #F4F4F4;
     border: 0.5pt solid $border; padding: 0 3pt; }
span.kbd-plus { font-size: 8pt; color: #767676; }
span.badge { font-size: 8pt; font-weight: 700; padding: 0 5pt; }

p.block-missing { border: 0.5pt dashed $border; background: $band; color: #595959;
     padding: 8pt 10pt; margin: 8pt 10pt 12pt 10pt; font-size: 9.5pt; }
"@
}

#endregion

#region -------------------------------------------------------------- title page

# A title page is a full-bleed panel (logo, kicker, title, standfirst) over a white
# band of key/value fields, then a section break. Word draws it from two tables -
# HTML import cannot express a page with no margins - and Convert-Markdown gives
# that first section zero margins and no running header once the document exists.

$script:CoverSkipField = @('title', 'excerpt', 'description', 'subtitle', 'kicker', 'draft', 'restricted', 'restrictedroles')

function ConvertTo-CoverLabel {
    # "preparedFor", "prepared_for" and "prepared-for" all read "Prepared for".
    param([Parameter(Mandatory)][string]$Key)
    $lk = $Key.ToLowerInvariant()
    if ($script:NdFrontMatterLabel.Contains($lk)) { return $script:NdFrontMatterLabel[$lk] }
    $words = (($Key -creplace '([a-z0-9])([A-Z])', '$1 $2') -replace '[_-]+', ' ').Trim().ToLowerInvariant()
    if (-not $words) { return $Key }
    $words.Substring(0, 1).ToUpperInvariant() + $words.Substring(1)
}

function Get-CoverField {
    # The label/value pairs printed along the bottom of the title page. cover.fields
    # picks and orders them; left empty, every frontmatter field that is not already
    # on the page somewhere else is shown, plus today's date when there is none.
    param(
        [AllowNull()] $Fields,
        [Parameter(Mandatory)] $Style
    )
    $out = [ordered]@{}
    if (-not $Fields) { $Fields = [ordered]@{} }
    $wanted = @($Style.cover.fields | Where-Object { $_ })

    if ($wanted.Count) {
        foreach ($k in $wanted) {
            $hit = @($Fields.Keys) | Where-Object { $_ -ieq $k } | Select-Object -First 1
            if ($hit -and $Fields[$hit]) { $out[(ConvertTo-CoverLabel $hit)] = [string]$Fields[$hit] }
        }
        return $out
    }

    foreach ($k in $Fields.Keys) {
        if ($k.ToLowerInvariant() -in $script:CoverSkipField) { continue }
        if (-not $Fields[$k]) { continue }
        $out[(ConvertTo-CoverLabel $k)] = [string]$Fields[$k]
    }
    if (-not (@($out.Keys) | Where-Object { $_ -match '^(?i)(date|published)$' })) {
        $out['Date'] = (Get-Date).ToString('d MMMM yyyy', [Globalization.CultureInfo]::InvariantCulture)
    }
    $out
}

function New-CoverHtml {
    # The title page markup. Labels and the kicker are upper-cased here rather than
    # with text-transform, which Word's HTML import does not reliably honour.
    param(
        [Parameter(Mandatory)][string]$Title,
        [AllowEmptyString()][string]$Subtitle,
        [AllowEmptyString()][string]$Kicker,
        [AllowNull()] $Fields,
        [AllowEmptyString()][string]$LogoToken,
        # Without a logo: a short two-colour bar, accent then rule.
        [string[]]$BarColours = @('#1F3864', '#C8102E')
    )
    $enc = { param($s) [Net.WebUtility]::HtmlEncode([string]$s) }
    $panel = [Text.StringBuilder]::new()
    if ($LogoToken) { [void]$panel.Append("<p class=""cover-logo"">$LogoToken</p>") }
    else {
        # Shaded runs of non-breaking spaces. A short paragraph border would be the
        # obvious way, but Word's import drops the indent that keeps it short.
        $run = { param($colour, $n) "<span style=""background:$colour;font-size:5pt"">" + ('&nbsp;' * $n) + '</span>' }
        [void]$panel.Append('<p class="cover-bar">' + (& $run $BarColours[0] 36) + (& $run $BarColours[-1] 12) + '</p>')
    }
    if ($Kicker) { [void]$panel.Append("<p class=""cover-kicker"">$(& $enc $Kicker.ToUpperInvariant())</p>") }
    [void]$panel.Append("<p class=""cover-title"">$(& $enc $Title)</p>")
    if ($Subtitle) { [void]$panel.Append("<p class=""cover-sub"">$(ConvertTo-NdInlineHtml $Subtitle)</p>") }

    $meta = '<p class="cover-value">&nbsp;</p>'
    $keys = @(if ($Fields) { $Fields.Keys })
    if ($keys.Count) {
        # Four to a row: more than that and the values start wrapping into each other.
        $rows = for ($i = 0; $i -lt $keys.Count; $i += 4) {
            $cells = foreach ($k in $keys[$i..([Math]::Min($i + 3, $keys.Count - 1))]) {
                '<td class="cover-field"><p class="cover-label">{0}</p><p class="cover-value">{1}</p></td>' -f
                (& $enc $k.ToUpperInvariant()), (ConvertTo-NdInlineHtml $Fields[$k])
            }
            '<tr>' + ($cells -join '') + '</tr>'
        }
        $meta = '<table class="cover-fields" summary="covermeta">' + ($rows -join '') + '</table>'
    }

    "<table class=""cover"" summary=""cover""><tr><td class=""cover-panel"">$panel</td></tr>" +
    "<tr><td class=""cover-meta"">$meta</td></tr></table>`n<p class=""cover-break"">@@MDCOVERBREAK@@</p>`n"
}

function Get-CoverCss {
    param([Parameter(Mandatory)] $Style)
    $c = $Style.colours
    $f = $Style.fonts
    $band = if ($Style.cover.band) { $Style.cover.band } else { $c.band }
    $headCol = if ($c.heading) { $c.heading } else { $c.accent }
    $titleFont = if ($f.title) { $f.title } else { $f.heading }
    @"
table.cover { width: 100%; border-collapse: collapse; margin: 0; }
table.cover td { border: none; padding: 0; }
td.cover-panel { background: $band; vertical-align: top; }
td.cover-meta { background: #FFFFFF; vertical-align: top; }
table.cover-fields { width: 100%; border-collapse: collapse; margin: 0; }
table.cover-fields td { border: none; padding: 0 12pt 10pt 0; vertical-align: top; background: #FFFFFF; }
p.cover-logo { margin: 0 0 30pt 0; }
p.cover-bar { margin: 0 0 34pt 0; font-size: 5pt; line-height: 5pt; }
p.cover-kicker { font-family: '$($f.heading)', '$($f.body)', sans-serif; font-size: 8.5pt; color: $($c.accent);
     letter-spacing: 1.2pt; margin: 0 0 10pt 0; }
p.cover-title { font-family: '$titleFont', '$($f.body)', sans-serif; font-size: 34pt; color: $headCol;
     line-height: 1.1; margin: 0 0 14pt 0; }
p.cover-sub { font-size: 13pt; color: $($c.muted); line-height: 1.3; margin: 0; }
p.cover-label { font-family: '$($f.heading)', '$($f.body)', sans-serif; font-size: 7pt; color: $($c.muted);
     letter-spacing: 1pt; margin: 0 0 3pt 0; }
p.cover-value { font-size: 10pt; color: $($c.text); margin: 0; }
p.cover-break { font-size: 1pt; line-height: 1pt; margin: 0; }
"@
}

#endregion

function Get-ConvertMarkdownVersion {
    <#
    .SYNOPSIS
        Returns the module version (yyyy.MM.dd).
    #>
    [CmdletBinding()]
    param()
    [pscustomobject]@{
        Module        = 'ConvertMarkdown'
        Version       = $script:Version
        ModuleVersion = (Get-Module ConvertMarkdown).Version
        Path          = $PSScriptRoot
    }
}

function Convert-Markdown {
    <#
    .SYNOPSIS
        Converts a Markdown file into a styled Word .docx, a PDF, or both.

    .DESCRIPTION
        Renders Markdown to HTML with Markdig (built into PowerShell 7's
        ConvertFrom-Markdown), wraps it in a corporate stylesheet, then uses Word
        automation to produce a real .docx with print layout, 10mm margins, running
        header and footer. PDF output is exported from the same document, so both
        formats are always identical.

        ```mermaid fenced blocks are rendered to PNG and embedded as pictures.
        Without this they land in the document as a grey box full of "graph TB"
        source.

        Markdown images - ![alt](screenshot.png) - are embedded as pictures too.
        Word silently discards <img src="file:///..."> from imported HTML, so
        without this they land in the document as empty placeholder boxes. Paths
        are resolved relative to the document; http(s) URLs are downloaded.

        Requires PowerShell 7 and Microsoft Word (desktop, for COM automation).

    .PARAMETER Path
        Path to the source .md file. Accepts pipeline input.

    .PARAMETER OutputPath
        Output file path. Defaults to the source name with the format's extension,
        written next to the source. When -Format Both, this names the .docx and the
        PDF sits beside it.

    .PARAMETER Format
        Docx (default), Pdf, or Both.

    .PARAMETER StyleSheet
        A style file to use instead of the one found automatically.

    .PARAMETER Theme
        A theme to build on, by name (a folder under ~/.convertmarkdown/themes) or
        path. Overrides the style file's own "theme" key; the style file's other keys
        still apply on top.

    .PARAMETER TitlePage
        Start the document with a title page: logo, kicker, title and standfirst on
        a full-bleed panel, frontmatter fields along the bottom, no running header.
        The title, excerpt and fields then come off the top of the first page.
        A theme can turn this on by default (cover.enabled).

    .PARAMETER NoTitlePage
        No title page, even when the style or theme enables one.

    .PARAMETER Orientation
        Portrait (default) or Landscape. Also settable as page.orientation.

    .PARAMETER Accent
        Primary heading / table header colour as a hex string. Default Office navy.

    .PARAMETER Margin
        Page margin in millimetres, all four sides. Default 10.

    .PARAMETER DiagramService
        Where mermaid diagrams are rendered. Default 'mermaid.ink'.

        !! PRIVACY !! mermaid.ink and the public kroki.io are third-party services.
        The source of every mermaid block is sent over the internet to be rendered,
        and may be cached or logged by the provider. For internal architecture
        documents, either point -DiagramServiceUrl at a self-hosted Kroki, or use
        -DiagramService none. Nothing else in the document is transmitted.

        'none' inserts a labelled placeholder box instead and makes no network call.

    .PARAMETER DiagramServiceUrl
        Base URL of a self-hosted renderer, e.g. https://kroki.example.internal. Keeps
        diagram source inside the network. Pair with -DiagramService kroki.

    .PARAMETER DiagramMaxWidth
        Maximum rendered diagram width in pixels. Default 640 (A4 text width at 96dpi).

    .PARAMETER ImageMaxWidth
        Maximum embedded image width in pixels. Default 640. Wider images are
        scaled down with their aspect ratio preserved; narrower ones are left at
        their natural size.

    .PARAMETER NoRemoteImages
        Do not download images referenced by http(s) URL. They are skipped with a
        warning and a placeholder. Local images are unaffected.

        Pair this with -DiagramService none for a conversion that makes no network
        call at all.

    .PARAMETER ExcludeRestricted
        Drop NextDocs !variant!# ROLE ... !endvariant! sections instead of printing
        them. Without this they are printed behind a "Restricted" banner and the
        conversion warns that the file carries them.

        The removal happens before diagrams are rendered, so a mermaid block inside
        a restricted section is never sent to a rendering service either.

    .PARAMETER TableOfContents
        Insert a Table of Contents after the title.

    .PARAMETER Open
        Leave the finished document open in Word instead of closing it.

    .PARAMETER Force
        Overwrite existing output files.

    .EXAMPLE
        Convert-Markdown .\CHANGELOG.md -TableOfContents

    .EXAMPLE
        Convert-Markdown .\Development-Cycle.md -Format Both -Force

    .EXAMPLE
        # Internal architecture doc - nothing leaves the machine:
        Convert-Markdown .\TECHNICAL-PROFILE.md -Format Pdf -DiagramService none

    .EXAMPLE
        Get-ChildItem *.md | Convert-Markdown -Format Pdf -Force

    .EXAMPLE
        # A proposal: title page, landscape, a named theme
        Convert-Markdown .\proposal.md -Format Both -TitlePage -Orientation Landscape -Theme acme
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [string] $Path,

        [string] $OutputPath,

        [ValidateSet('Docx', 'Pdf', 'Both')]
        [string] $Format = 'Docx',

        [string] $StyleSheet,

        [string] $Theme,

        [switch] $TitlePage,

        [switch] $NoTitlePage,

        [ValidateSet('Portrait', 'Landscape')]
        [string] $Orientation,

        [ValidatePattern('^#[0-9A-Fa-f]{6}$')]
        [string] $Accent,

        [ValidateRange(0, 50)]
        [Nullable[double]] $Margin,

        [ValidateSet('mermaid.ink', 'kroki', 'none')]
        [string] $DiagramService,

        [string] $DiagramServiceUrl,

        [ValidateRange(200, 2000)]
        [Nullable[int]] $DiagramMaxWidth,

        [ValidateRange(100, 2000)]
        [Nullable[int]] $ImageMaxWidth,

        [switch] $NoRemoteImages,

        [switch] $ExcludeRestricted,

        [switch] $TableOfContents,

        [switch] $Open,

        [switch] $Force
    )

    begin { $ErrorActionPreference = 'Stop' }

    process {
        $script:tempFiles = [Collections.Generic.List[string]]::new()
        Write-Verbose "ConvertMarkdown v$($script:Version)"

        # ------------------------------------------------------ resolve paths
        $source = (Resolve-Path -LiteralPath $Path).Path

        # ------------------------------------------------------------- style
        # Explicit parameters beat the style file; the style file beats defaults.
        $style = Resolve-Style -StyleSheet $StyleSheet -SourceDir (Split-Path -Parent $source) -Theme $Theme
        if ($style.Contains('themeFile')) {
            Write-Verbose "Theme from $($style.themeFile)"
        }
        if ($style.Contains('sourceFile')) {
            Write-Verbose "Style '$($style.name)' from $($style.sourceFile)"
        }

        if ($PSBoundParameters.ContainsKey('Orientation')) { $style.page.orientation = $Orientation }
        $landscape = [string]$style.page.orientation -match '^(?i)landscape$'

        $coverOn = if ($NoTitlePage) { $false } elseif ($TitlePage) { $true } else { [bool]$style.cover.enabled }

        if ($PSBoundParameters.ContainsKey('Accent')) { $style.colours.accent = $Accent }
        if ($PSBoundParameters.ContainsKey('Margin')) { $style.page.marginMm = $Margin }
        if ($PSBoundParameters.ContainsKey('DiagramService')) { $style.diagrams.service = $DiagramService }
        if ($PSBoundParameters.ContainsKey('DiagramServiceUrl')) { $style.diagrams.serviceUrl = $DiagramServiceUrl }
        if ($PSBoundParameters.ContainsKey('DiagramMaxWidth')) { $style.diagrams.maxWidthPx = $DiagramMaxWidth }

        # Belt and braces: Resolve-Style always starts from the defaults, but a
        # caller handing in a hand-built style should not trip StrictMode here.
        if (-not $style.Contains('images')) { $style['images'] = (Get-DefaultStyle).images }
        if ($PSBoundParameters.ContainsKey('ImageMaxWidth')) { $style.images.maxWidthPx = $ImageMaxWidth }
        if ($NoRemoteImages) { $style.images.allowRemote = $false }

        $Margin = [double]$style.page.marginMm
        $DiagramService = [string]$style.diagrams.service
        $DiagramServiceUrl = [string]$style.diagrams.serviceUrl
        $DiagramMaxWidth = [int]$style.diagrams.maxWidthPx
        $ImageMaxWidth = [int]$style.images.maxWidthPx
        $allowRemoteImages = -not $style.images.Contains('allowRemote') -or $style.images.allowRemote

        $mermaidTheme = if ($style.diagrams.Contains('theme')) { [string]$style.diagrams.theme } else { 'base' }
        $mermaidVars = $null
        # themeVariables are only honoured under the 'base' theme; sending them
        # with a named theme is silently ignored, so do not bother.
        $matchPalette = -not $style.diagrams.Contains('matchDocumentPalette') -or $style.diagrams.matchDocumentPalette
        if ($matchPalette -and $mermaidTheme -eq 'base') {
            $mermaidVars = Get-MermaidThemeVariables -Style $style
        }

        $wantDocx = $Format -in @('Docx', 'Both')
        $wantPdf = $Format -in @('Pdf', 'Both')

        if ($OutputPath) {
            $primary = [IO.Path]::GetFullPath($OutputPath)
        }
        else {
            $ext = if ($Format -eq 'Pdf') { '.pdf' } else { '.docx' }
            $primary = [IO.Path]::ChangeExtension($source, $ext)
        }

        $docxPath = if ($Format -eq 'Pdf') { $null } else { $primary }
        $pdfPath = if ($Format -eq 'Pdf') { $primary } else { [IO.Path]::ChangeExtension($primary, '.pdf') }
        if (-not $wantPdf) { $pdfPath = $null }

        $outDir = Split-Path -Parent $primary
        if (-not (Test-Path -LiteralPath $outDir)) {
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null
        }

        foreach ($existing in @($docxPath, $pdfPath | Where-Object { $_ })) {
            if ((Test-Path -LiteralPath $existing) -and -not $Force) {
                throw "Output already exists: $existing`nUse -Force to overwrite."
            }
        }

        if (-not $PSCmdlet.ShouldProcess($primary, "Convert Markdown to $Format")) { return }

        # Normalise to LF before anything looks at the text. Several patterns below
        # anchor with (?m)^...$, and .NET matches $ immediately before the \n - on a
        # CRLF file the \r sits between the text and that position, so "[ \t]*$"
        # never matches and the pattern silently finds nothing. That is how mermaid
        # fences were reaching Word as raw source.
        $markdown = (Get-Content -LiteralPath $source -Raw -Encoding UTF8) -replace "`r`n", "`n"

        # ------------------------------------------------------- YAML frontmatter
        $fm = Split-NdFrontMatter -Markdown $markdown -NoHeader:$coverOn
        $markdown = $fm.Markdown
        if ($fm.Found) {
            $where = if ($coverOn) { 'title page' } else { 'document header' }
            Write-Verbose "Frontmatter: $($fm.Fields.Count) field(s) lifted into the $where"
        }

        # --------------------------------------------- restricted sections (first)
        # Before the diagram pass, so a mermaid block inside a section the reader is
        # not entitled to see is never sent to a third-party renderer.
        $balance = Test-NdRestrictedBalance -Markdown $markdown
        if ($balance.Open -ne $balance.Close) {
            $msg = "$($balance.Open) !variant! marker(s) but $($balance.Close) !endvariant!. " +
            'An unterminated restricted section cannot be identified reliably.'
            # Refusing beats shipping. -ExcludeRestricted is an explicit instruction to
            # keep this content out of the output, and the pairing is what makes that
            # possible - so a mismatch here would remove nothing and say nothing.
            if ($ExcludeRestricted) {
                throw "$msg Fix the markers, or convert without -ExcludeRestricted."
            }
            Write-Warning "$msg The markers will appear in the document as literal text."
        }

        if ($ExcludeRestricted) {
            $before = $markdown.Length
            $markdown = Remove-NdRestricted -Markdown $markdown
            if ($markdown.Length -lt $before) { Write-Verbose 'Restricted !variant! section(s) removed' }
        }

        # ------------------------------------------------------ mermaid -> PNG
        # Word silently drops <img src="file:///..."> from imported HTML, so the
        # markup carries a text token and InlineShapes.AddPicture swaps each one for
        # a properly embedded picture after import.
        $script:diagramNo = 0
        $script:diagramsRendered = 0
        $script:diagramsFailed = 0
        $script:diagrams = [Collections.Generic.List[object]]::new()

        $markdown = [regex]::Replace($markdown, '(?ims)^[ \t]*```[ \t]*mermaid[ \t]*\r?$\r?\n(.*?)^[ \t]*```[ \t]*\r?$', {
                param($m)
                $script:diagramNo++
                $src = $m.Groups[1].Value.TrimEnd()
                $label = "Diagram $script:diagramNo"

                if ($DiagramService -eq 'none') {
                    return "`n<p class=""diagram-missing""><strong>$label</strong> &#8212; mermaid diagram, " +
                    "not rendered. See the Markdown source.</p>`n"
                }

                try {
                    $bytes = Invoke-MermaidRender -Source $src -Service $DiagramService `
                        -ServiceUrl $DiagramServiceUrl -Theme $mermaidTheme -ThemeVariables $mermaidVars
                    $png = Join-Path ([IO.Path]::GetTempPath()) ("md2doc_{0}.png" -f [Guid]::NewGuid().ToString('N'))
                    [IO.File]::WriteAllBytes($png, $bytes)
                    $script:tempFiles.Add($png)

                    $token = "@@MDDIAGRAM$($script:diagramNo)@@"
                    $script:diagrams.Add([pscustomobject]@{ Token = $token; Path = $png; Label = $label })

                    $script:diagramsRendered++
                    Write-Verbose "$label rendered ($($bytes.Length) bytes)"
                    return "`n<p class=""diagram"">$token</p>`n"
                }
                catch {
                    $script:diagramsFailed++
                    Write-Warning "$label failed to render via ${DiagramService}: $($_.Exception.Message). Keeping the source block."
                    return $m.Value
                }
            })

        if ($script:diagramNo -gt 0 -and $DiagramService -ne 'none') {
            Write-Verbose "$($script:diagramsRendered) of $($script:diagramNo) diagram(s) rendered via $DiagramService"
        }

        # Any ```mermaid fence still standing will reach Word as raw diagram source.
        # A failed render hands its fence back deliberately and is already counted;
        # anything above that means the fence pattern did not recognise the block.
        $leftover = [regex]::Matches($markdown, '(?im)^[ \t]*```[ \t]*mermaid[ \t]*\r?$').Count
        $unrecognised = $leftover - $script:diagramsFailed
        if ($unrecognised -gt 0) {
            $script:diagramsFailed += $unrecognised
            Write-Warning ("$unrecognised mermaid block(s) were not recognised and will appear in the document as " +
                'raw source. Check the fence spelling (```mermaid) and that the fence lines carry no trailing text.')
        }

        # ---------------------------------------------------- images -> picture
        # Same reason as the diagrams above: Word drops <img src="file:///...">
        # on HTML import and leaves an empty placeholder box, so each Markdown
        # image becomes a text token that AddPicture swaps for a real, embedded
        # picture once the document exists.
        $script:imageNo = 0
        $script:imagesEmbedded = 0
        $script:imagesMissing = 0
        $script:images = [Collections.Generic.List[object]]::new()
        $sourceDir = Split-Path -Parent $source

        # Fenced code is skipped so a document that *documents* Markdown keeps its
        # examples as text. Inline code is handled per-line below.
        $inFence = $false
        $imgLines = foreach ($line in ($markdown -split "`n")) {
            if ($line -match '^[ \t]*(```|~~~)') { $inFence = -not $inFence; $line; continue }
            if ($inFence -or $line -notmatch '!\[') { $line; continue }

            $out = [regex]::Replace($line, '!\[(?<alt>[^\]]*)\]\((?<src>[^)]+)\)', {
                    param($m)
                    # An image inside inline code is an example, not a picture.
                    # Match exposes no public Input property, so read the line the
                    # evaluator was built for - it is in scope for this iteration.
                    $before = $line.Substring(0, $m.Index)
                    if (([regex]::Matches($before, '`').Count % 2) -eq 1) { return $m.Value }

                    $alt = $m.Groups['alt'].Value
                    $file = Resolve-MarkdownImage -Source $m.Groups['src'].Value -BaseDir $sourceDir `
                        -AllowRemote $allowRemoteImages
                    if (-not $file) {
                        $script:imagesMissing++
                        Write-Warning "Image not found, leaving a placeholder: $($m.Groups['src'].Value)"
                        $label = if ($alt) { $alt } else { $m.Groups['src'].Value }
                        return "<span class=""image-missing"">[image not found: $([Net.WebUtility]::HtmlEncode($label))]</span>"
                    }

                    $script:imageNo++
                    $token = "@@MDIMAGE$($script:imageNo)@@"
                    $script:images.Add([pscustomobject]@{
                            Token        = $token
                            Path         = $file
                            Alt          = $alt
                            Label        = "Image $($script:imageNo)"
                            FixedWidthPt = $null
                        })
                    return $token
                })

            # An image alone on its line is a figure, not a run of text - centre it
            # and keep it off a page break, exactly as a diagram is treated.
            if ($out.Trim() -match '^@@MDIMAGE\d+@@$') { "<p class=""image"">$($out.Trim())</p>" }
            else { $out }
        }
        $markdown = $imgLines -join "`n"

        if ($script:imageNo -gt 0) {
            Write-Verbose "$($script:imageNo) image(s) queued for embedding"
        }

        # The title page logo goes in by the same token mechanism, at its own width.
        $logoToken = ''
        if ($coverOn -and $style.cover.logo) {
            if (Test-Path -LiteralPath $style.cover.logo -PathType Leaf) {
                $logoToken = '@@MDCOVERLOGO@@'
                $script:images.Add([pscustomobject]@{
                        Token        = $logoToken
                        Path         = (Resolve-Path -LiteralPath $style.cover.logo).Path
                        Alt          = 'Logo'
                        Label        = 'Title page logo'
                        FixedWidthPt = [double]$style.cover.logoWidthMm * 72 / 25.4
                    })
            }
            else {
                Write-Warning "Title page logo not found, using the accent bar instead: $($style.cover.logo)"
            }
        }

        # ------------------------------------------ extended (NextDocs) markdown
        # After the image pass, so a :::gallery is laying out picture tokens rather
        # than raw ![](..), and before Markdig, which would otherwise mangle every
        # one of these constructs. See the "extended markdown" region for the why.
        $imageAlt = @{}
        foreach ($img in $script:images) { $imageAlt[$img.Token] = $img.Alt }

        $markdown = Convert-NdInline -Markdown $markdown
        $nd = Expand-NdBlock -Markdown $markdown -ImageAlt $imageAlt -LabelRestricted:(-not $ExcludeRestricted)
        $markdown = $nd.Markdown

        $script:ndBlocks = $nd.Count
        $script:ndCallouts = $nd.Callouts
        $script:ndRestricted = $nd.Restricted

        if ($nd.Count -or $nd.Callouts) {
            Write-Verbose "$($nd.Count) extended block(s) and $($nd.Callouts) callout(s) rendered"
        }
        if ($script:ndKbd -or $script:ndBadge) {
            Write-Verbose "$($script:ndKbd) keyboard shortcut(s) and $($script:ndBadge) badge(s) rendered"
        }
        if ($nd.Restricted -gt 0) {
            Write-Warning ("$($nd.Restricted) restricted section(s) are included in the output behind a " +
                'Restricted banner. Use -ExcludeRestricted to leave them out.')
        }

        # ------------------------------------------- pre-process: metadata block
        # Consecutive "**Key:** value" lines before the first horizontal rule become
        # a key/value box rather than one run-together paragraph.
        $lines = $markdown -split "\r?\n"
        $firstRule = [Array]::FindIndex([string[]]$lines, [Predicate[string]] { param($l) $l.Trim() -eq '---' })
        if ($firstRule -lt 0) { $firstRule = $lines.Count }

        $out = [Collections.Generic.List[string]]::new()
        $i = 0
        while ($i -lt $lines.Count) {
            $run = [Collections.Generic.List[string[]]]::new()
            $j = $i
            while ($j -lt $firstRule -and $lines[$j] -match '^\*\*(.+?):\*\*\s*(.*)$') {
                $run.Add(@($Matches[1], $Matches[2]))
                $j++
            }
            if ($run.Count -ge 2) {
                $out.Add('<table class="meta">')
                foreach ($kv in $run) {
                    # This is raw HTML by the time Markdig sees it, and CommonMark
                    # does not look inside an HTML block - so `code`, *emphasis* and
                    # links in a value would reach the page as literal punctuation
                    # unless the cells are converted here.
                    $out.Add(('<tr><td class="k">{0}</td><td class="v">{1}</td></tr>' -f
                            (ConvertTo-NdInlineHtml $kv[0]), (ConvertTo-NdInlineHtml $kv[1])))
                }
                $out.Add('</table>')
                $out.Add('')
                $i = $j
            }
            else {
                $out.Add($lines[$i])
                $i++
            }
        }
        $markdown = $out -join "`n"

        # ------------------------------------------------------ markdown -> html
        $body = (ConvertFrom-Markdown -InputObject $markdown).Html

        # Extended blocks back in. Each token sat alone on its line, so Markdig wrapped
        # it in its own paragraph; that wrapper goes too, or every block ends up nested
        # inside a <p> and Word inherits the paragraph spacing twice.
        #
        # Repeated until nothing changes, because blocks nest: a callout inside a tab
        # panel is tokenised first, so its token is carried inside the tabs block's
        # stashed HTML rather than in the body. One pass substitutes the tabs block and
        # leaves the callout token sitting in the text it just inserted.
        for ($pass = 0; $pass -lt 8; $pass++) {
            $before = $body
            foreach ($token in $nd.Blocks.Keys) {
                $body = $body.Replace("<p>$token</p>", $nd.Blocks[$token]).Replace($token, $nd.Blocks[$token])
            }
            if ($body -eq $before) { break }
        }
        if ($body -match '@@NDBLOCK\d+@@') {
            Write-Warning 'An extended block was nested too deeply to render and its placeholder remains in the document.'
        }

        # Safety net: Markdig emits an inline <svg> icon for any GitHub alert the
        # callout pass did not claim (an indented one, say). Word imports the markup
        # as literal path data, so drop it and keep the text.
        $body = [regex]::Replace($body, '(?is)<svg\b.*?</svg>', '')

        # Markdig emits <del> for ~~strikethrough~~ and <ins> for ++inserted++. Word
        # reads those as tracked-change revisions rather than formatting: the text is
        # imported as a deletion, and because revisions are hidden on export it is
        # absent from the PDF entirely. Silent text loss, so map them to the
        # presentational tags Word treats as ordinary character formatting.
        $body = $body -replace '(?i)<del>', '<s>' -replace '(?i)</del>', '</s>'
        $body = $body -replace '(?i)<ins>', '<u>' -replace '(?i)</ins>', '</u>'

        # Task list checkboxes -> printable glyphs (Word renders <input> as a form field).
        $body = [regex]::Replace($body, '<input\b[^>]*type="checkbox"[^>]*>', {
                param($m)
                $glyph = if ($m.Value -match 'checked') { '&#9745;' } else { '&#9744;' }
                '<span class="cb">' + $glyph + '</span>'
            })

        # Banded table rows - Word ignores :nth-child(), so stripe them explicitly.
        $body = [regex]::Replace($body, '(?s)<tbody>.*?</tbody>', {
                param($m)
                $script:rowNo = 0
                [regex]::Replace($m.Value, '<tr>', {
                        $script:rowNo++
                        if ($script:rowNo % 2 -eq 0) { '<tr class="alt">' } else { '<tr>' }
                    })
            })

        # ----------------------------------------------------------- palette
        $accentRgb = $style.colours.accent
        $accentSoft = $style.colours.accentSoft
        $rule = $style.colours.rule
        $border = $style.colours.border
        $band = $style.colours.band
        $textCol = $style.colours.text
        $fontBody = $style.fonts.body
        $fontHead = $style.fonts.heading
        $fontMono = $style.fonts.mono
        $sizeBody = $style.fonts.bodySizePt

        $title = [IO.Path]::GetFileNameWithoutExtension($source)
        if ($coverOn -and $fm.Found -and $fm.Title) {
            $title = $fm.Title
        }
        elseif ($body -match '(?s)<h1[^>]*>(.*?)</h1>') {
            # Strip inline markup, then decode entities - otherwise "&amp;" reaches the header.
            $title = [Net.WebUtility]::HtmlDecode(($Matches[1] -replace '<[^>]+>', '')).Trim()
            # On a title page the heading is the cover title; printing it again as the
            # first line of page two would just repeat it.
            if ($coverOn) { $body = ([regex]'(?s)<h1[^>]*>.*?</h1>').Replace($body, '', 1) }
        }

        $coverCss = ''
        if ($coverOn) {
            $fields = if ($fm.Found) { $fm.Fields } else { [ordered]@{} }
            $pick = { param($k) $hit = @($fields.Keys) | Where-Object { $_ -ieq $k } | Select-Object -First 1; if ($hit) { [string]$fields[$hit] } else { '' } }
            $subtitle = & $pick 'subtitle'
            if (-not $subtitle -and $fm.Found) { $subtitle = $fm.Excerpt }
            $kicker = & $pick 'kicker'
            if (-not $kicker) { $kicker = [string]$style.cover.kicker }
            $body = (New-CoverHtml -Title $title -Subtitle $subtitle -Kicker $kicker `
                    -Fields (Get-CoverField -Fields $fields -Style $style) -LogoToken $logoToken `
                    -BarColours @($style.colours.accent, $style.colours.rule)) + $body
            $coverCss = Get-CoverCss -Style $style
        }

        $marginCss = '{0}mm' -f $Margin
        $pageSize = if ($landscape) { 'A4 landscape' } else { 'A4' }
        $encodedTitle = [Net.WebUtility]::HtmlEncode($title)
        $ndCss = Get-NdExtendedCss -Style $style
        $muted = $style.colours.muted
        $headCol = if ($style.colours.heading) { $style.colours.heading } else { $accentRgb }

        # Table header row. 'filled' is the original accent bar; 'rule' is a lighter,
        # report-style header: small uppercase grey labels over a single rule, rows
        # divided by hairlines rather than boxed.
        $tableCss = if ([string]$style.tables.headerStyle -eq 'rule') {
            @"
th { background: #FFFFFF; color: $muted; font-family: '$fontHead', '$fontBody', sans-serif; font-weight: 600;
     font-size: 7.5pt; text-transform: uppercase; letter-spacing: 0.8pt; text-align: left;
     border: none; border-bottom: 1pt solid $border; padding: 3pt 6pt 4pt 6pt; }
td { border: none; border-bottom: 0.5pt solid $border; padding: 3pt 6pt; vertical-align: top; }
"@
        }
        else {
            @"
th { background: $accentRgb; color: #FFFFFF; font-weight: 600; text-align: left;
     border: 0.5pt solid $accentRgb; padding: 2.5pt 6pt; }
td { border: 0.5pt solid $border; padding: 2.5pt 6pt; vertical-align: top; }
"@
        }

        # A note on the horizontal margins below. Word's HTML import turns CSS
        # padding on a bordered or shaded block into Borders.DistanceFromLeft/Right,
        # and it draws that box OUTWARD from the text - so a <pre> with 8pt of
        # horizontal padding keeps its text on the margin and hangs its grey box 8pt
        # past it on both sides, while tables and rules sit flush. Every such block
        # therefore carries a horizontal margin equal to its horizontal padding,
        # which becomes the paragraph indent and cancels the overhang exactly.
        # Change one and change the other.
        $html = @"
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
<head>
<meta charset="utf-8" />
<title>$encodedTitle</title>
<style>
@page { size: $pageSize; margin: $marginCss; }
body { font-family: '$fontBody', Calibri, sans-serif; font-size: ${sizeBody}pt; color: $textCol; line-height: 1.35; }
p { margin: 0 0 8pt 0; }

h1 { font-family: '$fontHead', '$fontBody', sans-serif; font-size: 21pt; color: $headCol;
     border-bottom: 2.25pt solid $rule; padding-bottom: 5pt; margin: 0 0 14pt 0; page-break-after: avoid; }
h2 { font-family: '$fontHead', '$fontBody', sans-serif; font-size: 14pt; color: $headCol;
     border-bottom: 0.75pt solid $border; padding-bottom: 3pt; margin: 20pt 0 8pt 0; page-break-after: avoid; }
h3 { font-family: '$fontHead', '$fontBody', sans-serif; font-size: 11.5pt; color: $accentSoft;
     margin: 14pt 0 5pt 0; page-break-after: avoid; }
h4 { font-size: ${sizeBody}pt; color: #333333; margin: 12pt 0 4pt 0; page-break-after: avoid; }

hr { border: none; border-top: 0.75pt solid $border; margin: 14pt 0; }

ul, ol { margin: 0 0 8pt 0; padding-left: 20pt; }
li { margin: 0 0 3pt 0; }
ul.contains-task-list { list-style: none; padding-left: 6pt; }
span.cb { font-family: 'Segoe UI Symbol'; font-size: 12pt; }

table { border-collapse: collapse; width: 100%; margin: 6pt 0 12pt 0; font-size: 9.5pt; }
$tableCss
tr.alt td { background: $band; }

table.meta { width: auto; margin: 0 0 14pt 0; font-size: 10pt; }
table.meta td { border: none; padding: 1pt 14pt 1pt 0; }
table.meta td.k { color: #595959; font-weight: 600; white-space: nowrap; }
table.meta td.v { color: #1F1F1F; }

code { font-family: $fontMono, 'Cascadia Mono', 'DejaVu Sans Mono', 'Courier New', monospace; font-size: 9pt; color: #A31515;
       background: #F4F4F4; padding: 0 2pt; }
pre { font-family: $fontMono, 'Cascadia Mono', 'DejaVu Sans Mono', 'Courier New', monospace; font-size: 9pt; background: #F7F7F9;
      border: 0.5pt solid #DDDDDD; border-left: 3pt solid $accentSoft; padding: 6pt 8pt;
      margin: 6pt 8pt 12pt 8pt; page-break-inside: avoid; }
pre code { background: none; color: $textCol; padding: 0; }

blockquote { border-left: 3pt solid $border; margin: 6pt 0 12pt 10pt; padding: 2pt 0 2pt 10pt; color: #444444; }

p.diagram { text-align: center; margin: 10pt 0 12pt 0; page-break-inside: avoid; }
p.image { text-align: center; margin: 10pt 0 12pt 0; page-break-inside: avoid; }
p.diagram-missing { border: 0.5pt dashed $border; background: $band; color: #595959;
      text-align: center; padding: 10pt; margin: 10pt 10pt 12pt 10pt; font-size: 9.5pt; }
span.image-missing { border: 0.5pt dashed $border; background: $band; color: #595959;
      padding: 2pt 6pt; font-size: 9pt; font-style: italic; }
strong { color: $textCol; }

$ndCss
$coverCss
</style>
</head>
<body>
$body
</body>
</html>
"@

        $tempHtml = Join-Path ([IO.Path]::GetTempPath()) ("md2doc_{0}.html" -f [Guid]::NewGuid().ToString('N'))
        [IO.File]::WriteAllText($tempHtml, $html, [Text.UTF8Encoding]::new($true))
        $script:tempFiles.Add($tempHtml)

        # ------------------------------------------------------ word automation
        # "Was Word already running" decides whether we are allowed to quit it at the
        # end - quitting an instance the user is typing in would be rude. Count only
        # instances with a window: a headless WINWORD is somebody's automation, not a
        # person. Without that distinction one leaked instance makes every later run
        # believe Word was already open, so none of them clean up either, and they
        # pile up until the file ends up locked by an orphan.
        $wordWasRunning = [bool](
            Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowTitle }
        )
        $word = $null; $doc = $null
        try {
            try { $word = New-Object -ComObject Word.Application }
            catch { throw "Microsoft Word is required but could not be started. $($_.Exception.Message)" }

            if (-not $wordWasRunning) { $word.Visible = $false }
            $word.DisplayAlerts = 0

            $doc = $word.Documents.Open($tempHtml, $false, $false)
            $doc.ActiveWindow.View.Type = 3          # wdPrintView

            # Page setup first, so diagram fitting measures the real text width.
            $ps = $doc.PageSetup
            $ps.PaperSize = 7                        # wdPaperA4
            # After the paper size: setting A4 again would put it back upright.
            $ps.Orientation = if ($landscape) { 1 } else { 0 }   # wdOrientLandscape / Portrait
            $marginPt = $word.MillimetersToPoints($Margin)
            $ps.TopMargin = $marginPt
            $ps.BottomMargin = $marginPt
            $ps.LeftMargin = $marginPt
            $ps.RightMargin = $marginPt

            # Header and footer must sit inside the margin or they collide with the
            # body. Half the margin, floored at 8pt, keeps them clear at any value.
            $chrome = [Math]::Max($word.MillimetersToPoints($Margin / 2), 8)
            $ps.HeaderDistance = $chrome
            $ps.FooterDistance = $chrome

            # Title page: a section break after the cover tables, and that first
            # section edge to edge. Done here, before any picture is fitted, because
            # once there are two sections with different margins the document-wide
            # PageSetup reports them as undefined - so $ps becomes the body's own.
            $coverSplit = $false
            if ($coverOn) {
                $rng = $doc.Content
                $find = $rng.Find
                $find.ClearFormatting()
                $find.Text = '@@MDCOVERBREAK@@'
                $find.Forward = $true
                $find.Wrap = 0                       # wdFindStop
                $find.MatchCase = $true
                if ($find.Execute()) {
                    $rng.Text = ''
                    $rng.InsertBreak(2)              # wdSectionBreakNextPage
                    $coverSplit = $doc.Sections.Count -ge 2
                }
                if ($coverSplit) {
                    $cps = $doc.Sections.Item(1).PageSetup
                    $cps.TopMargin = 0
                    $cps.BottomMargin = 0
                    $cps.LeftMargin = 0
                    $cps.RightMargin = 0
                    $cps.HeaderDistance = 0
                    $cps.FooterDistance = 0
                    $ps = $doc.Sections.Item($doc.Sections.Count).PageSetup
                }
                else {
                    Write-Warning 'The title page could not be separated from the body; it is printed inline.'
                }
            }

            # Swap each diagram token for its picture, embedded (LinkToFile = false)
            # so the document stands alone once the temp PNGs are cleaned up.
            if ($script:diagrams.Count -gt 0) {
                $maxW = [Math]::Min($DiagramMaxWidth * 0.75, $ps.PageWidth - $ps.LeftMargin - $ps.RightMargin)
                $maxH = $ps.PageHeight - $ps.TopMargin - $ps.BottomMargin - 24

                foreach ($d in $script:diagrams) {
                    $rng = $doc.Content
                    $find = $rng.Find
                    $find.ClearFormatting()
                    $find.Text = $d.Token
                    $find.Forward = $true
                    $find.Wrap = 0                   # wdFindStop
                    $find.MatchCase = $true
                    if (-not $find.Execute()) {
                        Write-Warning "$($d.Label): placeholder not found in the document."
                        continue
                    }
                    $rng.Text = ''                   # AddPicture inserts at the range,
                    $shape = $doc.InlineShapes.AddPicture($d.Path, $false, $true, $rng)   # it does not replace it
                    $shape.LockAspectRatio = -1      # msoTrue
                    if ($shape.Width -gt $maxW) { $shape.Width = $maxW }
                    if ($shape.Height -gt $maxH) { $shape.Height = $maxH }
                }
            }

            # Markdown images, by the same mechanism.
            if ($script:images.Count -gt 0) {
                $imgMaxW = [Math]::Min($ImageMaxWidth * 0.75, $ps.PageWidth - $ps.LeftMargin - $ps.RightMargin)
                $imgMaxH = $ps.PageHeight - $ps.TopMargin - $ps.BottomMargin - 24

                foreach ($i in $script:images) {
                    $rng = $doc.Content
                    $find = $rng.Find
                    $find.ClearFormatting()
                    $find.Text = $i.Token
                    $find.Forward = $true
                    $find.Wrap = 0                   # wdFindStop
                    $find.MatchCase = $true
                    if (-not $find.Execute()) {
                        Write-Warning "$($i.Label): placeholder not found in the document."
                        continue
                    }
                    $rng.Text = ''
                    try {
                        $shape = $doc.InlineShapes.AddPicture($i.Path, $false, $true, $rng)
                    }
                    catch {
                        Write-Warning "$($i.Label): Word could not embed $($i.Path) - $($_.Exception.Message)"
                        continue
                    }
                    $shape.LockAspectRatio = -1

                    # A picture in a :::gallery cell has to fit its column, not the
                    # page. Word cannot be asked for the column width here - the cell
                    # is empty until this picture lands in it - so the width comes
                    # from the column count the gallery block recorded.
                    $fitW = $imgMaxW
                    if ($script:ndGalleryCols.ContainsKey($i.Token)) {
                        $textW = $ps.PageWidth - $ps.LeftMargin - $ps.RightMargin
                        $fitW = ($textW / $script:ndGalleryCols[$i.Token]) - 10    # cell padding
                    }

                    # Alt text is what a screen reader announces, and what Word's
                    # accessibility checker looks for.
                    if ($i.Alt) { $shape.AlternativeText = $i.Alt }

                    # The title page logo is set to its own width, larger or smaller.
                    if ($i.FixedWidthPt) {
                        $shape.Width = $i.FixedWidthPt
                        continue
                    }

                    if ($shape.Width -gt $fitW) { $shape.Width = $fitW }
                    if ($shape.Height -gt $imgMaxH) { $shape.Height = $imgMaxH }
                    $script:imagesEmbedded++
                }
                Write-Verbose "$($script:imagesEmbedded) of $($script:imageNo) image(s) embedded"
            }

            # The running header and footer belong to the body. With a title page that
            # is the second section, unlinked first so the cover keeps a clean page.
            $section = $doc.Sections.Item($doc.Sections.Count)
            if ($coverSplit) {
                $section.Headers.Item(1).LinkToPrevious = $false
                $section.Footers.Item(1).LinkToPrevious = $false
            }

            # Header: document title, right aligned, hairline rule underneath.
            $hdr = $section.Headers.Item(1).Range
            $hdr.Text = $title
            $hdr.ParagraphFormat.Alignment = 2
            $hdr.Font.Name = $fontBody
            $hdr.Font.Size = 8
            $hdr.Font.Color = 8421504
            $hdr.Borders.Item(-3).LineStyle = 1
            $hdr.Borders.Item(-3).LineWidth = 4

            # Footer: "<date>            Page X of Y"
            $ftr = $section.Footers.Item(1).Range
            $ftr.Text = ''
            $ftr.Font.Name = $fontBody
            $ftr.Font.Size = 8
            $ftr.Font.Color = 8421504
            $ftr.ParagraphFormat.Alignment = 0
            $ftr.InsertAfter((Get-Date -Format 'yyyy-MM-dd') + "`t")
            $end = $section.Footers.Item(1).Range; $end.Collapse(0)
            $end.InsertAfter('Page ')
            $end = $section.Footers.Item(1).Range; $end.Collapse(0)
            $doc.Fields.Add($end, -1, 'PAGE', $false) | Out-Null
            $end = $section.Footers.Item(1).Range; $end.Collapse(0)
            $end.InsertAfter(' of ')
            $end = $section.Footers.Item(1).Range; $end.Collapse(0)
            $doc.Fields.Add($end, -1, 'NUMPAGES', $false) | Out-Null
            $section.Footers.Item(1).Range.ParagraphFormat.TabStops.ClearAll()
            $section.Footers.Item(1).Range.ParagraphFormat.TabStops.Add(
                ($ps.PageWidth - $ps.LeftMargin - $ps.RightMargin), 2) | Out-Null

            if ($coverSplit) {
                $cover = $doc.Sections.Item(1)
                # An empty header is still a line tall, and with no margin Word pushes
                # the page body below it - so the cover would start a line down and its
                # last row land on the next page. Shrink both to a 1pt sliver.
                foreach ($hf in @($cover.Headers.Item(1).Range, $cover.Footers.Item(1).Range)) {
                    $hf.Text = ''
                    $hf.Borders.Item(-3).LineStyle = 0   # no rule under an empty header
                    $hf.Font.Size = 1
                    $hf.ParagraphFormat.SpaceBefore = 0
                    $hf.ParagraphFormat.SpaceAfter = 0
                    $hf.ParagraphFormat.LineSpacingRule = 4   # wdLineSpaceExactly
                    $hf.ParagraphFormat.LineSpacing = 1
                }

                # Size the cover to the page: the panel takes the top 70%, the field
                # band the rest. Row heights are exact, a little short of the page, so
                # the paragraph carrying the section break cannot spill onto a page of
                # its own. Padding is set here because Word drops it from imported cells.
                $cps = $cover.PageSetup
                $coverTbl = $null
                foreach ($t in $doc.Tables) { if ($t.Descr -eq 'cover') { $coverTbl = $t; break } }
                if ($coverTbl) {
                    $padX = $word.MillimetersToPoints(22)
                    $coverTbl.Rows.LeftIndent = 0
                    $coverTbl.PreferredWidthType = 3                 # wdPreferredWidthPoints
                    $coverTbl.PreferredWidth = $cps.PageWidth
                    $panelH = [Math]::Round($cps.PageHeight * 0.70)
                    $row = $coverTbl.Rows.Item(1); $row.HeightRule = 2; $row.Height = $panelH    # wdRowHeightExactly
                    $row = $coverTbl.Rows.Item(2); $row.HeightRule = 2; $row.Height = $cps.PageHeight - $panelH - 20
                    $cell = $coverTbl.Cell(1, 1)
                    $cell.TopPadding = $word.MillimetersToPoints(26)
                    $cell.LeftPadding = $padX; $cell.RightPadding = $padX
                    $cell = $coverTbl.Cell(2, 1)
                    $cell.TopPadding = $word.MillimetersToPoints(12)
                    $cell.LeftPadding = $padX; $cell.RightPadding = $padX
                    foreach ($inner in $coverTbl.Tables) {
                        $inner.Descr = ''
                        $inner.Rows.AllowBreakAcrossPages = $false
                        $inner.Range.ParagraphFormat.SpaceBefore = 0
                    }
                }
            }

            # Repeat table header rows across page breaks, and stop rows splitting
            # mid-cell. Long reference tables are close to unreadable otherwise.
            #
            # The extended-markdown blocks are the exception: they are laid out as
            # tables but their first row is content, not a header, so repeating it
            # reprints (say) the "npm" tab label above the yarn panel on the next
            # page. They carry summary="nd", which Word imports as the table
            # description; that is read here and then cleared, so the marker never
            # reaches a reader as table alt text.
            foreach ($tbl in $doc.Tables) {
                try {
                    $tbl.Rows.AllowBreakAcrossPages = $false
                    # The title page is layout, sized above; nothing here applies.
                    if ($tbl.Descr -eq 'cover') {
                        $tbl.Descr = ''
                        $tbl.Rows.Item(1).HeadingFormat = $false
                        continue
                    }
                    if ($tbl.Descr -eq 'nd') {
                        $tbl.Descr = ''
                        $tbl.Rows.Item(1).HeadingFormat = $false
                        # Layout blocks hold prose - steps, API details, release
                        # notes - so their paragraph spacing is left alone.
                        continue
                    }
                    if ($tbl.Rows.Count -gt 1) { $tbl.Rows.Item(1).HeadingFormat = $true }

                    # A data-table cell holds one value, but Word still gives its
                    # paragraph the body's 8pt space-after - which is more vertical
                    # room than the cell padding itself and is what makes reference
                    # tables look loosely set. It cannot be done in CSS: "td p" is a
                    # descendant selector and Word drops those when it flattens the
                    # sheet to direct formatting.
                    $tbl.Range.ParagraphFormat.SpaceAfter = 0
                    $tbl.Range.ParagraphFormat.SpaceBefore = 0
                }
                catch { }
            }

            try { $doc.Content.ParagraphFormat.WidowControl = $true } catch { }

            if ($TableOfContents) {
                # The TOC 1-3 styles are Word's own, not ours: the field is built from
                # built-in styles that follow the document theme, so a TOC dropped into
                # a document whose body font came from our stylesheet arrives in
                # Calibri or Times and reads as though it were pasted in. Set the
                # styles before the field is added so the entries are right the first
                # time, and again after Update() because rebuilding the field re-applies
                # them.
                $tocStyles = { param($d)
                    # wdStyleTOC1, TOC2, TOC3. Colour is left to Word - the defaults
                    # are already near-black and the document ink is a hex string,
                    # not the BGR integer Font.Color wants.
                    foreach ($sid in -20, -21, -22) {
                        try {
                            $st = $d.Styles.Item($sid)
                            $st.Font.Name = $fontBody
                            $st.Font.Size = $sizeBody
                        }
                        catch { }
                    }
                }

                & $tocStyles $doc

                # The TOC goes after the whole document header, not merely after the
                # title. A document can carry a frontmatter standfirst and a key/value
                # box between the two, and dropping the field in at paragraph 2 splits
                # that header apart - title, contents, then subtitle. So walk forward
                # past anything that is still header: the excerpt paragraph, the
                # key/value table, and any blank paragraph between them.
                $anchor = 1
                $lastPara = $doc.Paragraphs.Count
                while ($anchor -lt $lastPara) {
                    $p = $doc.Paragraphs.Item($anchor + 1)
                    $isHeader = $false
                    try {
                        if ($p.Range.Information(12)) { $isHeader = $true }          # wdWithInTable
                        elseif ([string]$p.Style.NameLocal -eq 'excerpt') { $isHeader = $true }
                        elseif (-not $p.Range.Text.Trim()) { $isHeader = $true }
                    }
                    catch { }
                    if (-not $isHeader) { break }
                    $anchor++
                }

                # Insert BEFORE the first body paragraph rather than after the last
                # header one. The last header paragraph is the key/value table's own
                # end-of-row marker, and InsertParagraphAfter on that puts the new
                # paragraph inside the table, which swallows the heading that follows.
                if ($anchor -lt $lastPara) {
                    $bodyIdx = $anchor + 1
                    $doc.Paragraphs.Item($bodyIdx).Range.InsertParagraphBefore()
                    $tocPara = $doc.Paragraphs.Item($bodyIdx)
                    # It inherited the style of the paragraph it was pushed in front
                    # of; left as a heading it would list itself as a blank entry.
                    try { $tocPara.Style = -1 } catch { }        # wdStyleNormal
                    $tocRange = $tocPara.Range
                }
                else {
                    $doc.Paragraphs.Item($anchor).Range.InsertParagraphAfter()
                    $tocRange = $doc.Paragraphs.Item($anchor + 1).Range
                }
                $doc.TablesOfContents.Add($tocRange, $true, 1, 3) | Out-Null
                $doc.TablesOfContents.Item(1).Update()
                & $tocStyles $doc
            }

            try {
                $doc.BuiltInDocumentProperties.Item('Title').Value = $title
                $doc.BuiltInDocumentProperties.Item('Comments').Value =
                "Generated by ConvertMarkdown $($script:Version) from $(Split-Path -Leaf $source)"
            }
            catch { }

            # ------------------------------------------------------------ save
            if ($wantDocx) {
                $doc.SaveAs2($docxPath, 16)          # wdFormatDocumentDefault
            }

            if ($wantPdf) {
                # ExportAsFixedFormat keeps heading bookmarks and doc structure tags,
                # so the PDF is navigable and accessible rather than a flat print.
                $doc.ExportAsFixedFormat(
                    $pdfPath,
                    17, # wdExportFormatPDF
                    $false, # OpenAfterExport
                    0, # wdExportOptimizeForPrint
                    0, # wdExportAllDocument
                    1, 1, # From, To (ignored for whole document)
                    0, # wdExportDocumentContent
                    $true, # IncludeDocProps
                    $true, # KeepIRM
                    1, # wdExportCreateHeadingBookmarks
                    $true, # DocStructureTags
                    $true, # BitmapMissingFonts
                    $false               # UseISO19005_1
                )
            }

            # Word lays out lazily, and ComputeStatistics reports only what has
            # been paginated so far. On a long document that returns something
            # short and plausible - 3 or 4 - rather than an error, so it is worse
            # than useless. Reading Panes.Pages.Count forces a real layout pass;
            # ComputeStatistics is then correct and is kept as the fallback.
            $pages = 0
            try { $doc.Repaginate() } catch { }
            try { $pages = $doc.Windows.Item(1).Panes.Item(1).Pages.Count } catch { }
            if (-not $pages) { $pages = $doc.ComputeStatistics(2) }

            if ($Open) {
                if (-not $wordWasRunning) { $word.Visible = $true }
                $word.Activate()
            }
            else {
                $doc.Close(0)
                $doc = $null
                if (-not $wordWasRunning) { $word.Quit() }
            }

            [pscustomobject]@{
                Source           = $source
                Format           = $Format
                DocxPath         = $docxPath
                PdfPath          = $pdfPath
                Title            = $title
                Pages            = $pages
                Diagrams         = $script:diagramNo
                DiagramsRendered = $script:diagramsRendered
                DiagramsFailed   = $script:diagramsFailed
                DiagramService   = $DiagramService
                Images           = $script:imageNo
                ImagesEmbedded   = $script:imagesEmbedded
                ImagesMissing    = $script:imagesMissing
                ExtendedBlocks   = $script:ndBlocks
                Callouts         = $script:ndCallouts
                Restricted       = $script:ndRestricted
                MarginMm         = $Margin
                Orientation      = $(if ($landscape) { 'Landscape' } else { 'Portrait' })
                TitlePage        = $coverSplit
                Theme            = $(if ($style.Contains('themeFile')) { Split-Path -Leaf (Split-Path -Parent $style.themeFile) } else { $null })
                Style            = $style.name
                StyleFile        = $(if ($style.Contains('sourceFile')) { $style.sourceFile } else { '(built-in)' })
                Version          = $script:Version
            }
        }
        finally {
            if ($doc -and -not $Open) { try { $doc.Close(0) } catch { } }
            foreach ($o in @($doc, $word)) {
                if ($o) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null } catch { } }
            }
            foreach ($t in $script:tempFiles) {
                Remove-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

function Update-ConvertMarkdown {
    <#
    .SYNOPSIS
        Updates the installed module from GitHub.

    .DESCRIPTION
        Downloads the latest release of the repository, compares its version with
        the installed one, and replaces the installed module if it is newer.

        This works regardless of where the module was originally installed from,
        and needs no git client -- it fetches the branch archive over HTTPS.

        Installed with Install.ps1 -Link? Then the install points at your working
        copy and this is not the right tool: pull in that folder instead, and the
        change is live immediately.

    .PARAMETER CheckOnly
        Report what is available without changing anything.

    .PARAMETER Force
        Reinstall even when the installed version is already current.

    .PARAMETER Branch
        Branch to fetch. Default 'main'.

    .EXAMPLE
        Update-ConvertMarkdown -CheckOnly

    .EXAMPLE
        Update-ConvertMarkdown
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch] $CheckOnly,
        [switch] $Force,
        [string] $Branch = 'main'
    )

    $ErrorActionPreference = 'Stop'
    $repo = 'garethcheyne/PowerShell-Convert-Markdown'
    $installed = $script:Version
    $moduleDir = $PSScriptRoot

    # A symlinked install points back at a working copy; overwriting it would
    # clobber uncommitted work.
    $parent = Split-Path -Parent $moduleDir
    $self = Get-Item -LiteralPath $moduleDir -ErrorAction SilentlyContinue
    if ($self -and $self.LinkType -eq 'SymbolicLink') {
        Write-Warning "This module is a symlink to $($self.Target). Update it with 'git pull' there instead."
        return
    }

    $work = Join-Path ([IO.Path]::GetTempPath()) ('cmupd_{0}' -f [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $work -Force | Out-Null

    try {
        $zip = Join-Path $work 'src.zip'
        Write-Verbose "Fetching $repo ($Branch)"

        # codeload serves the archive directly. The github.com/archive URL relies
        # on a redirect that can return an HTML interstitial instead of the zip,
        # which surfaces as an opaque FileFormatException from Expand-Archive.
        Invoke-WebRequest -Uri "https://codeload.github.com/$repo/zip/refs/heads/$Branch" `
            -OutFile $zip -TimeoutSec 60 -Headers @{ 'User-Agent' = 'ConvertMarkdown' } -ErrorAction Stop

        # Verify it really is a zip before handing it to Expand-Archive, so a
        # captive portal or rate-limit page gives a comprehensible error.
        $magic = [byte[]]::new(4)
        $fs = [IO.File]::OpenRead($zip)
        try { $null = $fs.Read($magic, 0, 4) } finally { $fs.Dispose() }
        if ($magic[0] -ne 0x50 -or $magic[1] -ne 0x4B) {
            throw "Download was not a zip archive (got $([BitConverter]::ToString($magic))). A proxy or captive portal may have intercepted the request."
        }

        Expand-Archive -LiteralPath $zip -DestinationPath $work -Force
        $src = Get-ChildItem -LiteralPath $work -Directory |
            Where-Object { $_.Name -like 'PowerShell-Convert-Markdown*' } |
            Select-Object -First 1
        if (-not $src) { throw 'Downloaded archive did not contain the expected folder.' }

        $newManifest = Join-Path $src.FullName 'ConvertMarkdown.psd1'
        if (-not (Test-Path -LiteralPath $newManifest)) { throw 'Downloaded archive has no module manifest.' }

        $available = (Import-PowerShellDataFile -LiteralPath $newManifest).PrivateData.DisplayVersion

        $isNewer = [version]($available -replace '\.0*', '.') -gt [version]($installed -replace '\.0*', '.')

        $result = [pscustomobject]@{
            Installed = $installed
            Available = $available
            IsNewer   = $isNewer
            Action    = 'None'
            Path      = $moduleDir
        }

        if ($CheckOnly) {
            $result.Action = if ($isNewer) { 'Update available' } else { 'Up to date' }
            return $result
        }

        if (-not $isNewer -and -not $Force) {
            Write-Host "ConvertMarkdown $installed is current." -ForegroundColor Green
            $result.Action = 'Up to date'
            return $result
        }

        if (-not $PSCmdlet.ShouldProcess($moduleDir, "Update $installed -> $available")) { return }

        Copy-Item -Path (Join-Path $src.FullName 'ConvertMarkdown.ps*1') -Destination $moduleDir -Force

        Write-Host "ConvertMarkdown updated: $installed -> $available" -ForegroundColor Green
        Write-Host 'Open a new session, or run: Import-Module ConvertMarkdown -Force' -ForegroundColor Yellow

        $result.Action = "Updated to $available"
        return $result
    }
    finally {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Set-Alias -Name md2docx -Value Convert-Markdown
Set-Alias -Name md2pdf -Value Convert-Markdown
# Back-compat with the script this module replaces.
Set-Alias -Name Convert-MarkdownToWord -Value Convert-Markdown

Export-ModuleMember `
    -Function Convert-Markdown, New-ConvertMarkdownStyle, Update-ConvertMarkdown, Get-ConvertMarkdownVersion `
    -Alias md2docx, md2pdf, Convert-MarkdownToWord
