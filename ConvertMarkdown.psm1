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

$script:Version = '2026.09.02'

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
        page     = [ordered]@{ marginMm = 10 }
        fonts    = [ordered]@{
            body       = 'Segoe UI'
            heading    = 'Segoe UI Semibold'
            mono       = 'Consolas'
            bodySizePt = 10.5
        }
        colours  = [ordered]@{
            accent     = '#1F3864'
            accentSoft = '#2E74B5'
            rule       = '#C8102E'
            border     = '#B4C6E7'
            band       = '#F2F5FA'
            text       = '#1F1F1F'
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

function Resolve-Style {
    <#
    .SYNOPSIS
        Merges a style file over the built-in defaults.

    .DESCRIPTION
        Resolution order, first match wins:
          1. -StyleSheet
          2. $env:CONVERTMARKDOWN_STYLE
          3. .convertmarkdown.json beside the source document
          4. ~/.convertmarkdown/style.json
          5. built-in defaults

        A style file only needs the keys it changes; everything else falls back.
    #>
    param(
        [string] $StyleSheet,
        [string] $SourceDir
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

    if ($file) {
        $loaded = Import-StyleFile -LiteralPath $file
        foreach ($section in @('page', 'fonts', 'colours', 'diagrams', 'images')) {
            if ($loaded.Contains($section) -and $loaded[$section]) {
                foreach ($k in $loaded[$section].Keys) {
                    $style[$section][$k] = $loaded[$section][$k]
                }
            }
        }
        # US spelling accepted, since half the world writes it that way.
        if ($loaded.Contains('colors') -and $loaded['colors']) {
            foreach ($k in $loaded['colors'].Keys) { $style['colours'][$k] = $loaded['colors'][$k] }
        }
        if ($loaded.Contains('name')) { $style['name'] = $loaded['name'] }
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

    .PARAMETER Force
        Overwrite an existing file.

    .EXAMPLE
        New-ConvertMarkdownStyle

    .EXAMPLE
        # Per-repo style, picked up for documents in that folder
        New-ConvertMarkdownStyle -Path .\.convertmarkdown.json -Name 'Project docs'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string] $Path,
        [string] $Name = 'Custom',
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
    $style = [ordered]@{
        '$schema' = 'https://raw.githubusercontent.com/garethcheyne/PowerShell-Convert-Markdown/main/style.schema.json'
        name      = $Name
        page      = $defaults.page
        fonts     = $defaults.fonts
        colours   = $defaults.colours
        diagrams  = $defaults.diagrams
        images    = $defaults.images
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
        Base URL of a self-hosted renderer, e.g. https://kroki.internal.hn. Keeps
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
        $style = Resolve-Style -StyleSheet $StyleSheet -SourceDir (Split-Path -Parent $source)
        if ($style.Contains('sourceFile')) {
            Write-Verbose "Style '$($style.name)' from $($style.sourceFile)"
        }

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
                            Token = $token
                            Path  = $file
                            Alt   = $alt
                            Label = "Image $($script:imageNo)"
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
                    $out.Add(('<tr><td class="k">{0}</td><td class="v">{1}</td></tr>' -f $kv[0], $kv[1]))
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
        if ($body -match '(?s)<h1[^>]*>(.*?)</h1>') {
            # Strip inline markup, then decode entities - otherwise "&amp;" reaches the header.
            $title = [Net.WebUtility]::HtmlDecode(($Matches[1] -replace '<[^>]+>', '')).Trim()
        }

        $marginCss = '{0}mm' -f $Margin
        $encodedTitle = [Net.WebUtility]::HtmlEncode($title)

        $html = @"
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
<head>
<meta charset="utf-8" />
<title>$encodedTitle</title>
<style>
@page { size: A4; margin: $marginCss; }
body { font-family: '$fontBody', Calibri, sans-serif; font-size: ${sizeBody}pt; color: $textCol; line-height: 1.35; }
p { margin: 0 0 8pt 0; }

h1 { font-family: '$fontHead', '$fontBody', sans-serif; font-size: 21pt; color: $accentRgb;
     border-bottom: 2.25pt solid $rule; padding-bottom: 5pt; margin: 0 0 14pt 0; page-break-after: avoid; }
h2 { font-family: '$fontHead', '$fontBody', sans-serif; font-size: 14pt; color: $accentRgb;
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
th { background: $accentRgb; color: #FFFFFF; font-weight: 600; text-align: left;
     border: 0.5pt solid $accentRgb; padding: 4pt 6pt; }
td { border: 0.5pt solid $border; padding: 4pt 6pt; vertical-align: top; }
tr.alt td { background: $band; }

table.meta { width: auto; margin: 0 0 14pt 0; font-size: 10pt; }
table.meta td { border: none; padding: 1pt 14pt 1pt 0; }
table.meta td.k { color: #595959; font-weight: 600; white-space: nowrap; }
table.meta td.v { color: #1F1F1F; }

code { font-family: $fontMono, 'Cascadia Mono', 'DejaVu Sans Mono', 'Courier New', monospace; font-size: 9pt; color: #A31515;
       background: #F4F4F4; padding: 0 2pt; }
pre { font-family: $fontMono, 'Cascadia Mono', 'DejaVu Sans Mono', 'Courier New', monospace; font-size: 9pt; background: #F7F7F9;
      border: 0.5pt solid #DDDDDD; border-left: 3pt solid $accentSoft; padding: 6pt 8pt;
      margin: 6pt 0 12pt 0; page-break-inside: avoid; }
pre code { background: none; color: $textCol; padding: 0; }

blockquote { border-left: 3pt solid $border; margin: 6pt 0 12pt 0; padding: 2pt 0 2pt 10pt; color: #444444; }

p.diagram { text-align: center; margin: 10pt 0 12pt 0; page-break-inside: avoid; }
p.image { text-align: center; margin: 10pt 0 12pt 0; page-break-inside: avoid; }
p.diagram-missing { border: 0.5pt dashed $border; background: $band; color: #595959;
      text-align: center; padding: 10pt; margin: 10pt 0 12pt 0; font-size: 9.5pt; }
span.image-missing { border: 0.5pt dashed $border; background: $band; color: #595959;
      padding: 2pt 6pt; font-size: 9pt; font-style: italic; }
strong { color: $textCol; }
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
        $wordWasRunning = [bool](Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
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
                    if ($shape.Width -gt $imgMaxW) { $shape.Width = $imgMaxW }
                    if ($shape.Height -gt $imgMaxH) { $shape.Height = $imgMaxH }
                    # Alt text is what a screen reader announces, and what Word's
                    # accessibility checker looks for.
                    if ($i.Alt) { $shape.AlternativeText = $i.Alt }
                    $script:imagesEmbedded++
                }
                Write-Verbose "$($script:imagesEmbedded) of $($script:imageNo) image(s) embedded"
            }

            $section = $doc.Sections.Item(1)

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

            # Repeat table header rows across page breaks, and stop rows splitting
            # mid-cell. Long reference tables are close to unreadable otherwise.
            foreach ($tbl in $doc.Tables) {
                try {
                    $tbl.Rows.AllowBreakAcrossPages = $false
                    if ($tbl.Rows.Count -gt 1) { $tbl.Rows.Item(1).HeadingFormat = $true }
                }
                catch { }
            }

            try { $doc.Content.ParagraphFormat.WidowControl = $true } catch { }

            if ($TableOfContents) {
                $first = $doc.Paragraphs.Item(1).Range
                $first.InsertParagraphAfter()
                $tocRange = $doc.Paragraphs.Item(2).Range
                $doc.TablesOfContents.Add($tocRange, $true, 1, 3) | Out-Null
                $doc.TablesOfContents.Item(1).Update()
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
                MarginMm         = $Margin
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
