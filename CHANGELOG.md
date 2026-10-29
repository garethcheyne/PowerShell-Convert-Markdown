# Changelog — ConvertMarkdown

Versioned `yyyy.MM.dd`.

---

## 2026.10.29

### Added

- **Title page.** `-TitlePage` (or `cover.enabled` in a style or theme) starts the
  document with a full-bleed title page: logo or accent bar, kicker, title and
  standfirst on a tinted panel, and the frontmatter fields along the bottom. It is
  its own section with no running header; the body's header and footer start on
  page two. The title, standfirst and fields come off the top of the body rather
  than being printed twice. `-NoTitlePage` overrides a style that enables it.
- **Landscape.** `-Orientation Landscape` (or `page.orientation`) turns every page,
  title page included. Portrait remains the default.
- **Themes.** A theme is a folder holding `theme.json` and the assets it names (a
  logo). It layers between the built-in defaults and the style file, and is chosen
  with `-Theme`, a style file's `"theme"` key, or `$env:CONVERTMARKDOWN_THEME`.
  Themes are found in `$env:CONVERTMARKDOWN_THEMES`, `~/.convertmarkdown/themes` and
  the module's own `theming/` folder. Relative paths in a theme or style file
  resolve against that file, so a theme can live anywhere.
- `Install.ps1` copies the repo's `theming/` folder (gitignored, for private brand
  themes) to `~/.convertmarkdown/themes`, and `-DefaultTheme <name>` makes one the
  machine-wide default, keeping the previous style file as `style.json.bak`.
- `New-ConvertMarkdownStyle -Theme <name>` writes a style file that only selects the
  theme. A file listing every default would hide the whole theme.
- Style keys `fonts.title`, `colours.heading`, `colours.muted`,
  `tables.headerStyle` (`filled` or the lighter `rule`), and the `cover` section.
  All default to the previous output, so existing styles are unchanged.
- `Orientation`, `TitlePage` and `Theme` on the result object.

---

## 2026.09.03

### Added

- **NextDocs extended Markdown is rendered instead of mangled.** `:::` blocks,
  inline directives and role-gated sections previously reached Word as broken
  markup: Markdig's custom-container extension threw the `:::api GET /path`
  header away, read the `---` separator inside a block as a setext underline,
  split a diff into two bullet lists, and its generic-attributes extension
  swallowed `:badge[x]{warning}` into an HTML attribute on the paragraph. Every
  construct is now resolved ahead of Markdig and re-imagined for a page that
  cannot be clicked:
  - `> [!NOTE|TIP|IMPORTANT|WARNING|CAUTION]` — tinted panel with a coloured
    left bar, glyph and label
  - `:::steps` — numbered disc in the gutter, bold step title, body beneath
  - `:::tabs` — every panel printed under its own label bar, so no alternative
    is lost to whichever tab happened to be open
  - `:::details` — shown open, summary kept as the heading
  - `:::filetree` — box-drawing tree, folders in bold
  - `:::api` — coloured method chip and monospace path
  - `:::diff` — tinted lines with the `+`/`-` kept for grey-scale printing
  - `:::gallery` — picture grid, alt text as the caption
  - `:::release` — accent title bar with version and teams
  - `:kbd[Ctrl+S]` — boxed key caps; `:badge[text]{variant}` — coloured pill
- **YAML frontmatter becomes the document header.** `title` is the H1 (and so the
  running header and the `.docx` Title property), `excerpt` the standfirst, and
  the remaining fields join the key/value box. Previously Markdig welded the whole
  block into one enormous H2.
- `-ExcludeRestricted` drops `!variant!# ROLE ... !endvariant!` sections. Without
  it they are printed behind a dashed "Restricted" banner and the conversion warns
  that the file carries them; they used to pass through with their markers intact.
  The removal runs before the diagram pass, so a mermaid block inside a restricted
  section is never sent to a rendering service either.
- `ExtendedBlocks`, `Callouts` and `Restricted` on the result object.
- An `examples/` folder: a plain-Markdown baseline, a full extended-Markdown
  reference, a realistic NextDocs page, a styling walkthrough and an edge-case
  document, with the sample images the gallery examples need.
- **A Pester test suite** (`tests/`, run with `tests/Invoke-Tests.ps1`). Split by
  tag: the default set is pure string-to-string checks that need neither Word nor
  network and runs in about a second; the `Word`-tagged set does end-to-end
  conversions. The corpus is `examples/*.md`, and the suite fails if a block type,
  callout type or inline directive is missing from it — so a feature nobody
  documented is a feature the build rejects. Malformed input lives in
  `tests/fixtures/`.

### Fixed

- Pictures inside a `:::gallery` are sized to their column rather than the page,
  which used to push the table off the right margin.
- Markdig's inline `<svg>` icon for a GitHub alert is stripped; Word imported it
  as literal path data.
- **`~~strikethrough~~` text no longer disappears.** Markdig emits `<del>`, which
  Word imports as a *tracked deletion* rather than as formatting — so the text
  became a hidden revision and was absent from the PDF entirely. `<del>` and
  `<ins>` are now mapped to `<s>` and `<u>`, which Word treats as ordinary
  character formatting.
- **Code blocks, blockquotes and placeholder boxes sit on the margin.** Word turns
  CSS padding on a bordered or shaded block into a border distance and draws that
  box *outward*, so a code block's grey panel hung 8pt past the text column on both
  sides while tables and rules sat flush. Each such block now carries a horizontal
  margin equal to its horizontal padding, which cancels the overhang exactly.
- Inline markup in a `**Key:** value` box or a frontmatter field is converted
  rather than printed as literal backticks and asterisks. Those cells are raw HTML
  by the time Markdig runs, and CommonMark does not look inside an HTML block.

- **An unterminated `!variant!` no longer defeats `-ExcludeRestricted` silently.**
  The strip is regex-paired, so a missing or misspelt `!endvariant!` removed
  nothing and said nothing — the restricted section shipped. The markers are now
  counted (ignoring fenced and inline code, so a document *about* the syntax is not
  a false positive); an imbalance refuses the conversion under `-ExcludeRestricted`
  and warns otherwise.
- `-TableOfContents` puts the TOC after the whole document header rather than
  after the title, so a frontmatter standfirst and key/value box are no longer
  split from the H1 by the contents field.
- A frontmatter `title` is emitted as encoded text rather than as markdown, so
  markup in a title is shown rather than interpreted.
- A block nested inside another block — a callout inside a `:::tabs` panel, say —
  no longer leaves its `@@NDBLOCK…@@` placeholder in the document. The inner block
  is tokenised first, so its token travelled inside the outer block's stashed HTML
  and one substitution pass never reached it.

### Changed

- **Tables are set more tightly.** Word gives each cell paragraph the body's 8pt
  space-after, which was more vertical room than the cell padding itself; data
  tables now have it zeroed and cell padding is 2.5pt rather than 4pt. Extended
  blocks keep normal paragraph spacing — they hold prose, not single values.
- The `-DiagramServiceUrl` example in the help text no longer names a private
  internal host.

---

## 2026.09.02

### Added

- **Markdown images are embedded as real pictures.** `![alt](screenshot.png)` now
  arrives in the document as an embedded picture. Word silently discards
  `<img src="file:///...">` from imported HTML, so previously these landed as empty
  placeholder boxes and the only way round it was to post-process the `.docx`
  yourself. Uses the same token-and-`AddPicture` mechanism as mermaid diagrams.
  - Paths resolve relative to the **document**, not the working directory
  - Absolute paths, `file://` URIs, percent-encoding, `<angle brackets>` and a
    trailing `"title"` are all handled
  - `http(s)` URLs are downloaded to a temp file and embedded
  - An image alone on its line is centred and kept off a page break; an image
    inside a sentence stays inline
  - Alt text is written to the picture's `AlternativeText`, so screen readers
    announce it and Word's accessibility checker is satisfied
  - Images inside fenced code blocks and inline code are left as text, so a
    document about Markdown keeps its examples
  - A missing file leaves a visible `[image not found: ...]` placeholder and a
    warning rather than failing the conversion
- `-ImageMaxWidth` (default 640px) and the `images.maxWidthPx` style key.
- `-NoRemoteImages` and the `images.allowRemote` style key. With
  `-DiagramService none`, gives a conversion that makes no network call at all.
- `Images`, `ImagesEmbedded` and `ImagesMissing` on the result object.

### Changed

- `New-ConvertMarkdownStyle` and `style.example.json` now include the `images`
  section. Existing style files stay valid — the section falls back to defaults.

---

## 2026.08.29

First public release. Consolidates a script that had been copied into several
repos and drifted into four different versions; the newest of those (v1.1.1) is
the basis for this module.

### Added

- **PDF output.** `-Format Docx` (default), `Pdf`, or `Both`. The PDF is exported
  from the same Word document, so the two formats can never disagree.
  Uses `ExportAsFixedFormat` with heading bookmarks and document structure tags,
  so the PDF is navigable and accessible rather than a flat print.
- **Installable module.** `Install.ps1` places it on `PSModulePath`, so
  `Convert-Markdown` autoloads in any directory or repo. No more per-repo copies.
  - `-Scope CurrentUser` (default, no elevation) or `AllUsers`
  - `-Link` symlinks back to source for development
- **Style files.** Colours, fonts, margin and diagram defaults now live in a JSON
  (or YAML, or `.psd1`) file rather than in the module source. The module ships
  neutral defaults; branding belongs to the user.
  - `New-ConvertMarkdownStyle` writes a starter file with every key at its
    default, so it doubles as the reference
  - `-StyleSheet <path>` for a one-off
  - Resolution order: `-StyleSheet` → `$env:CONVERTMARKDOWN_STYLE` →
    `.convertmarkdown.json` beside the document → `~/.convertmarkdown/style.json`
    → built-in
  - `Install.ps1` offers to create one
- **`-Margin`** — page margin in millimetres, all four sides. Default **10**.
- **`-DiagramServiceUrl`** — point at a self-hosted Kroki so diagram source never
  leaves the network. The privacy answer for internal architecture documents.
- **`-Force`** — overwrite existing output. Without it an existing file is an
  error rather than a silent overwrite.
- **Pipeline input** — `Get-ChildItem *.md | Convert-Markdown -Format Pdf`.
- **Aliases** `md2docx`, `md2pdf`, and `Convert-MarkdownToWord` for back-compat.
- **`Update-ConvertMarkdown`** — fetches the latest published version over HTTPS,
  compares it with the installed one, and replaces the module only when newer.
  No git client needed. `-CheckOnly` reports without changing anything, and a
  symlinked install is detected and left alone so uncommitted work is not lost.
- **`Get-ConvertMarkdownVersion`**.
- **Table header rows repeat across page breaks**, and rows no longer split
  mid-cell. Long reference tables were close to unreadable before.
- **Widow/orphan control**, and Title / generator stamped into document
  properties so files stay identifiable in a document library.
- **`-WhatIf` / `-Confirm`** support.

### Changed

- **Renamed** `Convert-MarkdownToWord` → **`Convert-Markdown`**. The old name is
  no longer accurate now that Word is one of two output formats. It survives as
  an alias.
- **Page margins are now 10mm on all four sides.** Previously the stylesheet said
  10mm and the Word automation then overrode it to 15mm — see Fixed.
- Header and footer distance follows the margin (half of it, floored at 8pt), so
  page furniture stays clear of the body at any margin.

### Fixed

- **Margins did not match the stylesheet.** The CSS `@page` rule set 10mm but the
  Word `PageSetup` block immediately overrode all four sides to 15mm, so the
  stylesheet value was dead code and every document came out at 15mm. Both now
  derive from `-Margin`.

---

## Inherited history

From the `Convert-MarkdownToWord.ps1` script this module replaces, retained for
traceability.

| Version | Date | Change |
| --- | --- | --- |
| 1.1.1 | 2026-08-21 | **Fix:** mermaid fences were never detected in CRLF files, so every diagram reached Word as raw source and the run reported `Diagrams = 0`. Source is normalised to LF before parsing, and an unrecognised fence now warns instead of failing silently. |
| 1.1.0 | 2026-08-02 | Render ` ```mermaid ` blocks to embedded PNG pictures. |
| 1.0.0 | 2026-07-31 | Initial: Markdig → styled HTML → Word .docx, with A4 page setup, header and footer. |

### Why a module

The script was copied per-repo and run in place. Seven copies had drifted into
four different versions, and the oldest predated mermaid support entirely — so
whether diagrams rendered depended on which folder you happened to run it from.
Installing once removes that class of problem.

---

## Sample style files

`style.example.json` (every option at its default) and `style.schema.json` (JSON
Schema, for editor completion and validation) are in the repo root. Generated
style files carry a `$schema` reference so they get the same editor support.

---

## Diagram styling

`diagrams.theme` and `diagrams.matchDocumentPalette` map the document palette
onto mermaid's theme variables, so rendered diagrams match the rest of the
document instead of arriving in mermaid's own colours. Only the `base` theme
honours theme variables; the named themes ignore them. A diagram carrying its
own `%%{init: ...}%%` directive is sent untouched.

Kroki now receives a JSON body (`diagram_source` + `diagram_options`) rather than
raw text, so the theme reaches it too.

---

## Page count fix

`Pages` on the result object was unreliable. Word lays out lazily and
`ComputeStatistics(wdStatisticPages)` reports only what has been paginated so
far, so a long document returned a short but plausible number — 3 or 4 — rather
than an error. Reading `Panes.Pages.Count` forces a real layout pass first.
