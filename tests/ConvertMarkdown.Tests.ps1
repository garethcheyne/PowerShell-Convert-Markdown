<#
    Pester 5+ suite for ConvertMarkdown.

    Two halves, split by tag:

      (default)  Pure string -> string tests over the parsing and rendering
                 functions. No Word, no network. Runs in a couple of seconds and
                 is what CI should gate on.

      -Tag Word  End-to-end conversions. Needs Microsoft Word on the machine, so
                 exclude it anywhere Word is not installed:
                     Invoke-Pester .\tests -ExcludeTagFilter Word

    The corpus is examples/*.md. Those files are documentation first and fixtures
    second, which is deliberate: a construct nobody thought to document is a
    construct nobody is testing either. Deliberately malformed input lives in
    tests/fixtures instead, so the examples stay publishable.
#>

# Discovery-time: -ForEach on an It is evaluated before any BeforeAll runs, so the
# file list has to exist at file scope as well as inside BeforeAll.
$ExampleNames = (Get-ChildItem (Join-Path (Split-Path -Parent $PSScriptRoot) 'examples\*.md') -Exclude 'README.md').Name

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:ExampleDir = Join-Path $Root 'examples'
    $script:FixtureDir = Join-Path $PSScriptRoot 'fixtures'
    $script:OutDir = Join-Path ([IO.Path]::GetTempPath()) ('cm-tests-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force $script:OutDir | Out-Null

    Import-Module (Join-Path $Root 'ConvertMarkdown.psd1') -Force

    # Source text, normalised the way Convert-Markdown normalises it. Several
    # patterns anchor with (?m)^...$ and .NET puts $ before the \n, so a stray \r
    # makes them silently match nothing.
    function Get-Md {
        param([string]$Path)
        (Get-Content -LiteralPath $Path -Raw -Encoding UTF8) -replace "`r`n", "`n"
    }

    # Mirrors Convert-Markdown's pre-Word stages so the HTML can be asserted on
    # without automating Word. The diagram and image passes are left out; they are
    # covered by the Word-tagged tests further down.
    #
    # If the order of passes in Convert-Markdown changes, change it here too.
    function Get-BodyHtml {
        param([string]$Markdown, [switch]$ExcludeRestricted)
        InModuleScope ConvertMarkdown -Parameters @{ md = $Markdown; excl = [bool]$ExcludeRestricted } {
            $m = $md
            if ($excl) { $m = Remove-NdRestricted -Markdown $m }
            $m = (Split-NdFrontMatter -Markdown $m).Markdown
            $m = Convert-NdInline -Markdown $m
            $nd = Expand-NdBlock -Markdown $m -ImageAlt @{} -LabelRestricted:(-not $excl)
            $html = (ConvertFrom-Markdown -InputObject $nd.Markdown).Html
            for ($pass = 0; $pass -lt 8; $pass++) {
                $prev = $html
                foreach ($t in $nd.Blocks.Keys) {
                    $html = $html.Replace("<p>$t</p>", $nd.Blocks[$t]).Replace($t, $nd.Blocks[$t])
                }
                if ($html -eq $prev) { break }
            }
            $html = [regex]::Replace($html, '(?is)<svg\b.*?</svg>', '')
            $html = $html -replace '(?i)<del>', '<s>' -replace '(?i)</del>', '</s>'
            [pscustomobject]@{ Html = $html; Blocks = $nd.Count; Callouts = $nd.Callouts; Restricted = $nd.Restricted }
        }
    }

    $script:Examples = Get-ChildItem (Join-Path $ExampleDir '*.md') -Exclude 'README.md'
}

AfterAll {
    Remove-Item $script:OutDir -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'Corpus' {
    It 'has example documents to test against' {
        $script:Examples.Count | Should -BeGreaterThan 0
    }

    It 'covers every block type somewhere in the examples' {
        $all = ($script:Examples | ForEach-Object { Get-Md $_.FullName }) -join "`n"
        foreach ($type in 'tabs', 'details', 'steps', 'filetree', 'api', 'diff', 'gallery', 'sketch', 'release') {
            $all | Should -Match ":::$type" -Because "no example exercises :::$type"
        }
    }

    It 'covers every callout type somewhere in the examples' {
        $all = ($script:Examples | ForEach-Object { Get-Md $_.FullName }) -join "`n"
        foreach ($type in 'NOTE', 'TIP', 'IMPORTANT', 'WARNING', 'CAUTION') {
            $all | Should -Match "\[!$type\]" -Because "no example exercises a $type callout"
        }
    }

    It 'covers the inline directives and role-gated sections' {
        $all = ($script:Examples | ForEach-Object { Get-Md $_.FullName }) -join "`n"
        $all | Should -Match ':kbd\['
        $all | Should -Match ':badge\['
        $all | Should -Match ':#fluentui'
        $all | Should -Match '!variant!'
    }
}

Describe 'Frontmatter' {
    It 'lifts title and excerpt out of the body' {
        $r = InModuleScope ConvertMarkdown {
            Split-NdFrontMatter -Markdown "---`ntitle: The Title`nexcerpt: The excerpt.`n---`n`n## Body"
        }
        $r.Found | Should -BeTrue
        $r.Title | Should -Be 'The Title'
        $r.Markdown | Should -Match '<h1>The Title</h1>'
        $r.Markdown | Should -Match 'class="excerpt">The excerpt\.'
        $r.Markdown | Should -Match '## Body'
    }

    It 'accepts description and publishedAt as aliases' {
        $r = InModuleScope ConvertMarkdown {
            Split-NdFrontMatter -Markdown "---`ntitle: T`ndescription: Alias text.`npublishedAt: 2026-09-03`n---`n`nBody"
        }
        $r.Markdown | Should -Match 'class="excerpt">Alias text\.'
        $r.Markdown | Should -Match '>Published<'
    }

    It 'reads a YAML list into a single value' {
        $r = InModuleScope ConvertMarkdown {
            Split-NdFrontMatter -Markdown "---`ntitle: T`nrestrictedRoles:`n  - Alpha`n  - Beta`n---`n`nBody"
        }
        $r.Markdown | Should -Match 'Alpha, Beta'
    }

    It 'suppresses draft:false and restricted:false but keeps the true case' {
        $off = InModuleScope ConvertMarkdown {
            Split-NdFrontMatter -Markdown "---`ntitle: T`ndraft: false`nrestricted: false`n---`n`nBody"
        }
        $off.Markdown | Should -Not -Match '>Draft<'
        $off.Markdown | Should -Not -Match '>Restricted<'

        $on = InModuleScope ConvertMarkdown {
            Split-NdFrontMatter -Markdown "---`ntitle: T`nrestricted: true`n---`n`nBody"
        }
        $on.Markdown | Should -Match '>Restricted<'
    }

    It 'encodes markup in a title rather than interpreting it' {
        $r = InModuleScope ConvertMarkdown {
            Split-NdFrontMatter -Markdown "---`ntitle: A <b>bold</b> and R&D title`n---`n`nBody"
        }
        $r.Markdown | Should -Match '&lt;b&gt;'
        $r.Markdown | Should -Not -Match '<h1>A <b>'
    }

    It 'leaves a document without frontmatter untouched' {
        $md = "# Plain`n`nText."
        $r = InModuleScope ConvertMarkdown -Parameters @{ m = $md } { Split-NdFrontMatter -Markdown $m }
        $r.Found | Should -BeFalse
        $r.Markdown | Should -Be $md
    }

    It 'does not treat an unterminated block as frontmatter' {
        $r = InModuleScope ConvertMarkdown {
            Split-NdFrontMatter -Markdown "---`ntitle: Never Closed`n`n## Body"
        }
        $r.Found | Should -BeFalse
    }
}

Describe 'Block rendering' {
    BeforeAll {
        $script:Edge = Get-BodyHtml (Get-Md (Join-Path $ExampleDir 'edge-cases.md'))
        $script:Ext = Get-BodyHtml (Get-Md (Join-Path $ExampleDir 'extended-markdown.md'))
    }

    It 'renders steps as a numbered table with a disc glyph and a bold title' {
        $script:Ext.Html | Should -Match '<table class="nd steps"'
        $script:Ext.Html | Should -Match 'class="step-num">&#10102;'      # U+2776, step 1
        $script:Ext.Html | Should -Match 'class="step-title"'
    }

    It 'switches glyph range past step ten' {
        $script:Edge.Html | Should -Match '&#10111;'                      # U+277F, step 10
        $script:Edge.Html | Should -Match '&#9322;'                       # U+246A, step 11
    }

    It 'prints every tab panel, each behind its own label' {
        $labels = [regex]::Matches($script:Ext.Html, 'class="tab-label">([^<]+)<')
        $labels.Count | Should -BeGreaterOrEqual 3
        $bodies = [regex]::Matches($script:Ext.Html, 'class="tab-body"')
        $bodies.Count | Should -Be $labels.Count -Because 'every label needs its panel'
    }

    It 'colours the API method chip and keeps the path' {
        $script:Ext.Html | Should -Match 'class="api-method"[^>]*background:#0969DA'   # GET
        $script:Ext.Html | Should -Match 'class="api-path">/documents/:id<'
    }

    It 'gives an unrecognised HTTP method the neutral chip' {
        $script:Edge.Html | Should -Match 'background:#57606A;">&#160;TRACE'
    }

    It 'tints diff rows and keeps the +/- markers in the text' {
        $script:Ext.Html | Should -Match 'background:#E6FFEC[^>]*><p style="margin:0;">\+'
        $script:Ext.Html | Should -Match 'background:#FFEBE9[^>]*><p style="margin:0;">-'
    }

    It 'draws the file tree with box-drawing connectors and bold folders' {
        $script:Ext.Html | Should -Match '<pre class="filetree">'
        $script:Ext.Html | Should -Match '&#9500;&#9472;&#9472;'          # tee
        $script:Ext.Html | Should -Match '&#9492;&#9472;&#9472;'          # elbow
        $script:Ext.Html | Should -Match '<strong>ConvertMarkdown/</strong>'
    }

    It 'lays a gallery out at the requested column count' {
        # The image pass upstream turns every ![](..) into a token before the gallery
        # is built, so the builder is fed the tokens directly here.
        $body = "@@MDIMAGE1@@`n@@MDIMAGE2@@`n@@MDIMAGE3@@`n@@MDIMAGE4@@"
        foreach ($case in @(@{ H = ''; Cols = 3 }, @{ H = '2'; Cols = 2 }, @{ H = '3x2'; Cols = 3 }, @{ H = '1'; Cols = 1 })) {
            $html = InModuleScope ConvertMarkdown -Parameters @{ h = $case.H; b = $body } {
                New-NdGallery -Header $h -Body $b -Alt @{ '@@MDIMAGE1@@' = 'First' }
            }
            $rows = [regex]::Matches($html, '<tr>').Count
            $rows | Should -Be ([Math]::Ceiling(4 / $case.Cols)) -Because "header '$($case.H)' means $($case.Cols) column(s)"
            # Short rows are padded so the columns keep their width down the table.
            ([regex]::Matches($html, '<td class="gallery-cell"').Count) | Should -Be ($rows * $case.Cols)
        }
        $html = InModuleScope ConvertMarkdown { New-NdGallery -Header '' -Body '@@MDIMAGE1@@' -Alt @{ '@@MDIMAGE1@@' = 'Caption text' } }
        $html | Should -Match 'class="gallery-cap">Caption text<'
    }

    It 'renders all five callout types with their own colour and label' {
        foreach ($pair in @(@('#0969DA', 'Note'), @('#1A7F37', 'Tip'), @('#8250DF', 'Important'),
                            @('#9A6700', 'Warning'), @('#CF222E', 'Caution'))) {
            $pattern = 'color:{0};"><span class="sym">.+?</span>&#160;{1}</p>' -f $pair[0], $pair[1]
            $script:Ext.Html | Should -Match $pattern
        }
    }

    It 'says a sketch cannot be drawn rather than printing its id' {
        $script:Ext.Html | Should -Match 'class="block-missing"'
        $script:Ext.Html | Should -Not -Match '0f8fad5b-d9cb-469f-a165-70867728950e'
    }

    It 'puts the version and teams on the release bar' {
        $script:Ext.Html | Should -Match 'class="release-head">Release 2026\.09\.03\.1<'
        $script:Ext.Html | Should -Match 'class="release-teams">'
    }

    It 'substitutes a block nested inside another block' {
        # A callout is tokenised before the tab block that contains it, so its token
        # travels inside the tabs block's stashed HTML rather than in the body. A
        # single substitution pass leaves it sitting there as literal text.
        $r = Get-BodyHtml ":::tabs`n@tab One`n> [!TIP]`n> Nested callout.`n:::"
        $r.Html | Should -Not -Match '@@NDBLOCK'
        $r.Html | Should -Match 'class="tab-label">One<'
        $r.Html | Should -Match '&#160;Tip</p><p>Nested callout\.'
    }

    It 'leaves no placeholder token anywhere in the example corpus' {
        foreach ($ex in $script:Examples) {
            (Get-BodyHtml (Get-Md $ex.FullName)).Html |
                Should -Not -Match '@@(NDBLOCK|MDDIAGRAM)\d+@@' -Because "$($ex.Name) must render every block"
        }
    }

    It 'marks every layout table so Word does not repeat its first row' {
        $tables = [regex]::Matches($script:Ext.Html, '<table class="nd [^"]+"([^>]*)>')
        $tables.Count | Should -BeGreaterThan 0
        foreach ($t in $tables) { $t.Groups[1].Value | Should -Match 'summary="nd"' }
    }
}

Describe 'Inline directives' {
    BeforeAll { $script:Edge = Get-BodyHtml (Get-Md (Join-Path $ExampleDir 'edge-cases.md')) }

    It 'renders each key of a chord as its own cap' {
        $script:Edge.Html | Should -Match '<span class="kbd">Ctrl</span><span class="kbd-plus">\+</span><span class="kbd">Alt</span>'
    }

    It 'infers a badge variant from the word when none is given' {
        $r = Get-BodyHtml 'Status :badge[stable] and :badge[deprecated] and :badge[beta].'
        $r.Html | Should -Match 'background:#DAFBE1;color:#116329;">stable'
        $r.Html | Should -Match 'background:#FFEBE9;color:#A40E26;">deprecated'
        $r.Html | Should -Match 'background:#FFF8C5;color:#7D4E00;">beta'
    }

    It 'falls back to the default badge for an unknown variant' {
        $script:Edge.Html | Should -Match 'background:#F1F2F4;color:#374151;">unknown'
    }

    It 'removes a Fluent icon but leaves a Lucide one alone' {
        $script:Edge.Html | Should -Not -Match '#fluentui'
        $script:Edge.Html | Should -Match '<code>:Rocket:</code>'
    }

    It 'leaves a YouTube link as an ordinary hyperlink' {
        $script:Edge.Html | Should -Match '<a href="https://www\.youtube\.com/watch\?v='
    }
}

Describe 'Code is never rewritten' {
    It 'leaves extended syntax inside a fenced block as text' {
        $r = Get-BodyHtml (Get-Md (Join-Path $ExampleDir 'edge-cases.md'))
        $fence = [regex]::Match($r.Html, '(?s)<pre><code class="language-markdown">(.*?)</code></pre>').Groups[1].Value
        $fence | Should -Match ':::steps'
        $fence | Should -Match '\[!NOTE\]'
        $fence | Should -Match ':kbd\[Ctrl\+S\]'
        $fence | Should -Match '!variant!'
    }

    It 'leaves extended syntax inside inline code as text' {
        $r = Get-BodyHtml 'Literal: `:kbd[Ctrl+S]` and `:badge[beta]{warning}` stay put.'
        $r.Html | Should -Match '<code>:kbd\[Ctrl\+S\]</code>'
        $r.Html | Should -Match '<code>:badge\[beta\]\{warning\}</code>'
    }

    It 'finds no blocks at all in documentation about the syntax' {
        foreach ($doc in 'README.md', 'CHANGELOG.md') {
            $r = Get-BodyHtml (Get-Md (Join-Path $Root $doc))
            $r.Blocks | Should -Be 0 -Because "$doc only documents the syntax"
        }
        $r = Get-BodyHtml (Get-Md (Join-Path $ExampleDir 'README.md'))
        $r.Blocks | Should -Be 0
    }
}

Describe 'Restricted sections' {
    It 'labels a restricted section by default' {
        $r = Get-BodyHtml (Get-Md (Join-Path $ExampleDir 'nextdocs-page.md'))
        $r.Restricted | Should -Be 1
        $r.Html | Should -Match 'Restricted &#8212; Platform-Ops'
    }

    It 'removes the section entirely when excluded' {
        $md = Get-Md (Join-Path $ExampleDir 'nextdocs-page.md')
        $r = Get-BodyHtml $md -ExcludeRestricted
        $r.Restricted | Should -Be 0
        $r.Html | Should -Not -Match 'Rotation is quarterly'
        $r.Html | Should -Match 'not in the'   # public text after the block survives
    }

    It 'counts balanced markers, ignoring fenced and inline code' {
        foreach ($doc in 'README.md', 'CHANGELOG.md') {
            $b = InModuleScope ConvertMarkdown -Parameters @{ m = (Get-Md (Join-Path $Root $doc)) } {
                Test-NdRestrictedBalance -Markdown $m
            }
            $b.Open | Should -Be $b.Close -Because "$doc documents the syntax and must not look unbalanced"
        }
    }

    It 'detects an unterminated marker' {
        $b = InModuleScope ConvertMarkdown -Parameters @{ m = (Get-Md (Join-Path $FixtureDir 'unclosed-variant.md')) } {
            Test-NdRestrictedBalance -Markdown $m
        }
        $b.Open | Should -Be 1
        $b.Close | Should -Be 0
    }
}

Describe 'Malformed input is survivable' {
    It 'survives <File> without throwing or leaking a placeholder token' -ForEach @(
        @{ File = 'unclosed-block.md' }
        @{ File = 'empty-frontmatter.md' }
        @{ File = 'unclosed-frontmatter.md' }
        @{ File = 'markup-in-title.md' }
        @{ File = 'empty-blocks.md' }
        @{ File = 'no-frontmatter.md' }
        @{ File = 'crlf.md' }
        @{ File = 'unclosed-variant.md' }
    ) {
        $md = Get-Md (Join-Path $FixtureDir $File)
        { Get-BodyHtml $md } | Should -Not -Throw
        (Get-BodyHtml $md).Html | Should -Not -Match '@@ND(BLOCK|TOC)'
    }

    It 'leaves an unclosed block as ordinary text rather than eating the rest of the file' {
        $r = Get-BodyHtml (Get-Md (Join-Path $FixtureDir 'unclosed-block.md'))
        $r.Blocks | Should -Be 0
        $r.Html | Should -Match 'No closing fence follows'
    }

    It 'handles CRLF sources identically to LF' {
        $r = Get-BodyHtml (Get-Md (Join-Path $FixtureDir 'crlf.md'))
        $r.Blocks | Should -Be 1
        $r.Html | Should -Match '<h1>CRLF Document</h1>'
    }
}

Describe 'Themes and title page' {
    BeforeAll {
        # A throwaway theme: a folder with theme.json and a "logo" beside it.
        $script:ThemeDir = Join-Path $script:OutDir 'themes\testbrand'
        New-Item -ItemType Directory -Force $script:ThemeDir | Out-Null
        Set-Content -LiteralPath (Join-Path $script:ThemeDir 'logo.png') -Value 'not really a png'
        @{
            name    = 'Test brand'
            colours = @{ accent = '#112233'; heading = '#445566' }
            tables  = @{ headerStyle = 'rule' }
            cover   = @{ enabled = $true; logo = 'logo.png' }
        } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $script:ThemeDir 'theme.json')
    }

    It 'layers a theme over the defaults and resolves its logo beside it' {
        # An empty style file, so the machine's own ~/.convertmarkdown/style.json
        # cannot land on top of the theme.
        $empty = Join-Path $script:OutDir 'empty-style.json'
        Set-Content -LiteralPath $empty -Value '{}'
        $s = InModuleScope ConvertMarkdown -Parameters @{ t = $script:ThemeDir; e = $empty } {
            Resolve-Style -StyleSheet $e -SourceDir $null -Theme $t
        }
        $s.colours.accent | Should -Be '#112233'
        $s.colours.border | Should -Be '#B4C6E7'          # untouched default
        $s.tables.headerStyle | Should -Be 'rule'
        $s.cover.enabled | Should -BeTrue
        $s.cover.logo | Should -Be (Join-Path $script:ThemeDir 'logo.png')
    }

    It 'lets a style file name the theme and still override it' {
        $file = Join-Path $script:OutDir 'themed-style.json'
        @{ theme = $script:ThemeDir; colours = @{ accent = '#ABCDEF' } } | ConvertTo-Json | Set-Content -LiteralPath $file
        $s = InModuleScope ConvertMarkdown -Parameters @{ f = $file } { Resolve-Style -StyleSheet $f -SourceDir $null }
        $s.colours.accent | Should -Be '#ABCDEF'           # style file wins
        $s.colours.heading | Should -Be '#445566'          # theme still applies
    }

    It 'says where it looked when a theme is missing' {
        { InModuleScope ConvertMarkdown { Resolve-ThemeFile -Theme 'no-such-theme-anywhere' } } |
            Should -Throw -ExpectedMessage "*Theme 'no-such-theme-anywhere' not found*"
    }

    It 'labels frontmatter keys readably' {
        InModuleScope ConvertMarkdown {
            ConvertTo-CoverLabel 'preparedFor' | Should -Be 'Prepared for'
            ConvertTo-CoverLabel 'prepared_for' | Should -Be 'Prepared for'
            ConvertTo-CoverLabel 'author' | Should -Be 'Author'
        }
    }

    It 'shows the remaining frontmatter fields on the cover, and dates an undated one' {
        $f = InModuleScope ConvertMarkdown {
            $fields = [ordered]@{ title = 'T'; excerpt = 'E'; author = 'A'; status = 'Draft' }
            Get-CoverField -Fields $fields -Style (Get-DefaultStyle)
        }
        @($f.Keys) | Should -Be @('Author', 'Status', 'Date')
    }

    It 'uses cover.fields to pick and order the fields, with no added date' {
        $f = InModuleScope ConvertMarkdown {
            $style = Get-DefaultStyle
            $style.cover.fields = @('status', 'Author')
            Get-CoverField -Fields ([ordered]@{ author = 'A'; status = 'Draft'; version = '2' }) -Style $style
        }
        @($f.Keys) | Should -Be @('Status', 'Author')
    }

    It 'keeps the title, excerpt and fields out of the body when they go on a title page' {
        $r = InModuleScope ConvertMarkdown {
            Split-NdFrontMatter -Markdown "---`ntitle: Hello`nexcerpt: Standfirst`nauthor: A`n---`nBody text" -NoHeader
        }
        $r.Markdown | Should -Be 'Body text'
        $r.Title | Should -Be 'Hello'
        $r.Excerpt | Should -Be 'Standfirst'
    }

    It 'builds the cover markup with its section-break placeholder' {
        $h = InModuleScope ConvertMarkdown {
            New-CoverHtml -Title 'Plan & budget' -Subtitle 'Why' -Kicker 'Proposal' `
                -Fields ([ordered]@{ Author = 'A' }) -LogoToken ''
        }
        $h | Should -Match 'summary="cover"'
        $h | Should -Match 'Plan &amp; budget'
        $h | Should -Match 'PROPOSAL'
        $h | Should -Match '@@MDCOVERBREAK@@'
        $h | Should -Match 'class="cover-bar"'                # no logo: the accent bar
    }
}

Describe 'End-to-end conversion' -Tag Word {
    It 'converts <_> to a PDF' -ForEach $ExampleNames {
        $src = Join-Path $script:ExampleDir $_
        $out = Join-Path $script:OutDir ([IO.Path]::ChangeExtension($_, '.pdf'))
        $r = Convert-Markdown $src -OutputPath $out -Format Pdf -Force -DiagramService none -WarningAction SilentlyContinue
        Test-Path $out | Should -BeTrue
        $r.Pages | Should -BeGreaterThan 0
        $r.Version | Should -Match '^\d{4}\.\d{2}\.\d{2}$'
    }

    It 'embeds every resolvable image and reports the missing one' {
        $src = Join-Path $script:ExampleDir 'edge-cases.md'
        $out = Join-Path $script:OutDir 'edge-img.docx'
        $r = Convert-Markdown $src -OutputPath $out -Format Docx -Force -DiagramService none -WarningAction SilentlyContinue
        $r.ImagesMissing | Should -Be 1 -Because 'edge-cases.md references one absent file on purpose'
        $r.ImagesEmbedded | Should -Be ($r.Images)
        $r.ImagesEmbedded | Should -BeGreaterThan 0
    }

    It 'produces a document with no tracked revisions and no repair prompt' {
        $out = Join-Path $script:OutDir 'integrity.docx'
        $null = Convert-Markdown (Join-Path $script:ExampleDir 'core-features.md') -OutputPath $out `
            -Format Docx -Force -DiagramService none -WarningAction SilentlyContinue
        $w = New-Object -ComObject Word.Application
        try {
            $w.Visible = $false; $w.DisplayAlerts = 0
            $d = $w.Documents.Open($out, $false, $true)
            # <del> imported as a tracked deletion is how ~~strikethrough~~ used to
            # vanish from the PDF: present in the .docx, hidden on export.
            $d.Revisions.Count | Should -Be 0
            $rng = $d.Content
            $f = $rng.Find; $f.ClearFormatting(); $f.Text = 'strikethrough'; $f.Forward = $true; $f.Wrap = 0
            $f.Execute() | Should -BeTrue
            $rng.Font.StrikeThrough | Should -Be -1
            $d.Close(0)
        }
        finally { $w.Quit() }
    }

    It 'puts the table of contents after the document header, not inside it' {
        $out = Join-Path $script:OutDir 'toc.docx'
        $null = Convert-Markdown (Join-Path $script:ExampleDir 'nextdocs-page.md') -OutputPath $out `
            -Format Docx -Force -TableOfContents -DiagramService none -WarningAction SilentlyContinue
        $w = New-Object -ComObject Word.Application
        try {
            $w.Visible = $false; $w.DisplayAlerts = 0
            $d = $w.Documents.Open($out, $false, $true)
            $d.TablesOfContents.Count | Should -Be 1
            $styles = @(); $i = 0
            foreach ($p in $d.Paragraphs) { $i++; if ($i -gt 12) { break }; $styles += [string]$p.Style.NameLocal }
            $styles[0] | Should -Be 'Heading 1'
            $styles[1] | Should -Be 'excerpt' -Because 'the standfirst must stay with the title'
            # Every heading in the body must reach the field.
            $toc = $d.TablesOfContents.Item(1).Range.Text
            $toc | Should -Match 'Before you start'
            $toc | Should -Match 'Rollback'
            $d.Close(0)
        }
        finally { $w.Quit() }
    }

    It 'refuses to convert an unterminated restricted section with -ExcludeRestricted' {
        $src = Join-Path $script:FixtureDir 'unclosed-variant.md'
        $out = Join-Path $script:OutDir 'refused.pdf'
        { Convert-Markdown $src -OutputPath $out -Format Pdf -Force -ExcludeRestricted -DiagramService none } |
            Should -Throw -ExpectedMessage '*!endvariant!*'
        Test-Path $out | Should -BeFalse
    }

    It 'keeps restricted content out of the .docx when excluded' {
        $out = Join-Path $script:OutDir 'excluded.docx'
        $null = Convert-Markdown (Join-Path $script:ExampleDir 'nextdocs-page.md') -OutputPath $out `
            -Format Docx -Force -ExcludeRestricted -DiagramService none -WarningAction SilentlyContinue
        $w = New-Object -ComObject Word.Application
        try {
            $w.Visible = $false; $w.DisplayAlerts = 0
            $d = $w.Documents.Open($out, $false, $true)
            $text = $d.Content.Text
            $text | Should -Not -Match 'Rotation is quarterly'
            $text | Should -Match 'service desk'
            $d.Close(0)
        }
        finally { $w.Quit() }
    }

    It 'adds a title page as its own section, in landscape when asked' {
        $src = Join-Path $script:OutDir 'cover.md'
        Set-Content -LiteralPath $src -Encoding UTF8 -Value "---`ntitle: Cover test`nauthor: Tester`n---`n`n## Body`n`nText."
        $out = Join-Path $script:OutDir 'cover.docx'
        $r = Convert-Markdown $src -OutputPath $out -Format Docx -Force -TitlePage -Orientation Landscape -DiagramService none
        $r.TitlePage | Should -BeTrue
        $r.Orientation | Should -Be 'Landscape'
        $r.Pages | Should -Be 2
        $w = New-Object -ComObject Word.Application
        try {
            $w.Visible = $false; $w.DisplayAlerts = 0
            $d = $w.Documents.Open($out, $false, $true)
            $d.Sections.Count | Should -Be 2
            $d.Sections.Item(1).PageSetup.LeftMargin | Should -Be 0
            $d.Sections.Item(2).PageSetup.Orientation | Should -Be 1      # wdOrientLandscape
            $d.Sections.Item(1).Headers.Item(1).Range.Text.Trim() | Should -BeNullOrEmpty
            $d.Sections.Item(2).Headers.Item(1).Range.Text | Should -Match 'Cover test'
            $d.Close(0)
        }
        finally { $w.Quit() }
    }

    It 'stays portrait with no title page by default' {
        $out = Join-Path $script:OutDir 'portrait.pdf'
        $r = Convert-Markdown (Join-Path $script:ExampleDir 'styled.md') -OutputPath $out -Format Pdf -Force -DiagramService none `
            -NoTitlePage -WarningAction SilentlyContinue
        $r.Orientation | Should -Be 'Portrait'
        $r.TitlePage | Should -BeFalse
    }

    It 'writes nothing under -WhatIf' {
        $out = Join-Path $script:OutDir 'whatif.pdf'
        Convert-Markdown (Join-Path $script:ExampleDir 'styled.md') -OutputPath $out -Format Pdf -WhatIf | Out-Null
        Test-Path $out | Should -BeFalse
    }

    It 'accepts pipeline input without leaking state between documents' {
        $results = Get-ChildItem (Join-Path $script:ExampleDir '*.md') -Exclude 'README.md' |
            Convert-Markdown -Format Pdf -Force -DiagramService none -WarningAction SilentlyContinue |
            ForEach-Object { $_ }
        $results.Count | Should -Be $script:Examples.Count
        ($results | Where-Object { $_.Source -match 'core-features' }).ExtendedBlocks | Should -Be 0
        ($results | Where-Object { $_.Source -match 'extended-markdown' }).ExtendedBlocks | Should -BeGreaterThan 0
    }
}
