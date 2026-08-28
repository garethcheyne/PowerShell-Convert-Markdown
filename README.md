# ConvertMarkdown

Converts Markdown into a styled Word `.docx`, a PDF, or both — with mermaid
diagrams rendered as real embedded pictures.

**Version:** 2026.08.29 · **Requires:** PowerShell 7+ and Microsoft Word (desktop, Windows)

---

## Install

```powershell
git clone https://github.com/garethcheyne/PowerShell-Convert-Markdown.git
cd PowerShell-Convert-Markdown
.\Install.ps1
```

That places the module on `PSModulePath`, so `Convert-Markdown` autoloads from
**any directory**. No `Import-Module`, no per-repo copies.

The installer checks for Word, probes the diagram service, and offers to create a
style file.

While developing the module itself, install a symlink so edits take effect
immediately:

```powershell
.\Install.ps1 -Link -Force
```

### Updating

`Install.ps1` copies from whatever folder you run it in — it does **not** fetch
anything. Re-running it reinstalls that same local copy.

To pull the latest published version:

```powershell
Update-ConvertMarkdown -CheckOnly   # what's available, changes nothing
Update-ConvertMarkdown              # fetch and replace if newer
```

It downloads the branch archive over HTTPS, compares versions, and only replaces
the module when the available one is newer (`-Force` overrides). No git client
required. If you installed with `-Link` it will tell you to `git pull` in your
working copy instead, rather than overwriting uncommitted work.

---

## Use

```powershell
# Word document next to the source
Convert-Markdown .\CHANGELOG.md

# PDF, with a table of contents
Convert-Markdown .\design.md -Format Pdf -TableOfContents

# Both formats at once
Convert-Markdown .\design.md -Format Both -Force

# Every markdown file in the folder
Get-ChildItem *.md | Convert-Markdown -Format Pdf -Force

# Leave it open in Word to check before sending
Convert-Markdown .\proposal.md -Open
```

Aliases: `md2docx`, `md2pdf`, and `Convert-MarkdownToWord`.

---

## Styling

The module ships deliberately neutral defaults. Colours, fonts and margins live
in a **style file**, so you never edit the module to rebrand output.

```powershell
# Starter file with every setting at its default - doubles as the reference
New-ConvertMarkdownStyle

# Per-project style, picked up automatically for documents in that folder
New-ConvertMarkdownStyle -Path .\.convertmarkdown.json -Name 'Project docs'
```

```jsonc
{
  "name": "My brand",
  "page":     { "marginMm": 10 },
  "fonts":    { "body": "Segoe UI", "heading": "Segoe UI Semibold",
                "mono": "Consolas", "bodySizePt": 10.5 },
  "colours":  { "accent": "#1F3864", "accentSoft": "#2E74B5", "rule": "#C8102E",
                "border": "#B4C6E7", "band": "#F2F5FA", "text": "#1F1F1F" },
  "diagrams": { "service": "mermaid.ink", "serviceUrl": "", "maxWidthPx": 640 }
}
```

A style file only needs the keys it changes — everything else falls back to the
default. `colors` is accepted as well as `colours`.

**Resolution order**, first match wins:

| | Source |
| --- | --- |
| 1 | `-StyleSheet <path>` |
| 2 | `$env:CONVERTMARKDOWN_STYLE` |
| 3 | `.convertmarkdown.json` beside the document |
| 4 | `~/.convertmarkdown/style.json` |
| 5 | Built-in defaults |

Explicit parameters (`-Margin`, `-Accent`, …) always win over the style file.

**Formats:** `.json` works out of the box. `.psd1` uses
`Import-PowerShellDataFile`. `.yaml` / `.yml` need the `powershell-yaml` module.

---

## Parameters

| Parameter | Default | Notes |
| --- | --- | --- |
| `-Path` | — | Source `.md`. Accepts pipeline input. |
| `-Format` | `Docx` | `Docx`, `Pdf`, or `Both`. |
| `-OutputPath` | *alongside source* | With `-Format Both`, names the `.docx`; the PDF sits beside it. |
| `-StyleSheet` | *resolved* | Style file for this run. |
| `-Margin` | `10` | Page margin in millimetres, all four sides. |
| `-Accent` | `#1F3864` | Heading and table-header colour. |
| `-DiagramService` | `mermaid.ink` | `mermaid.ink`, `kroki`, or `none`. |
| `-DiagramServiceUrl` | — | Base URL of a self-hosted renderer. |
| `-DiagramMaxWidth` | `640` | Max diagram width in pixels. |
| `-TableOfContents` | off | Insert a TOC after the title. |
| `-Open` | off | Leave open in Word instead of closing. |
| `-Force` | off | Overwrite existing output. |

Supports `-WhatIf`, `-Confirm` and `-Verbose`.

---

## Mermaid diagrams

` ```mermaid ` fenced blocks are rendered to PNG and embedded as pictures. Without
this they land in the document as a grey box full of `flowchart TD` source.

> ### Privacy — read this before using it on confidential documents
>
> `mermaid.ink` and the public `kroki.io` are **third-party services**. The source
> of every mermaid block is sent over the internet to be rendered, and may be
> cached or logged by the provider. Nothing else in the document is transmitted.
>
> For internal architecture or governance documents, choose one of:
>
> ```powershell
> # Nothing leaves the machine - a labelled placeholder box is inserted instead
> Convert-Markdown .\architecture.md -DiagramService none
>
> # Self-hosted renderer - diagram source stays inside your network
> Convert-Markdown .\architecture.md -DiagramService kroki `
>                  -DiagramServiceUrl https://kroki.internal.example
> ```
>
> Or set `diagrams.service` in your style file so the safe choice is the default.

**On the public Kroki instance:** it returns HTTP 500 or times out for mermaid
diagrams, and has done since at least 2026-08-02. `mermaid.ink` is the default for
that reason. Use `kroki` with `-DiagramServiceUrl` pointing at your own instance.

### If diagrams don't render

| Symptom | Cause |
| --- | --- |
| Raw `flowchart TD` in the document | The fence was not recognised. Check it is exactly ` ```mermaid ` with no trailing text. A warning is emitted when this happens. |
| `Diagrams = 0` on a file that has them | Older versions failed to match fences in **CRLF** files — `(?m)$` does not match before `\r\n`. Fixed: source is normalised to LF before parsing. |
| `failed to render via mermaid.ink` | Network, proxy, or a syntax error in the diagram. The source block is kept so nothing is lost. |

---

## What it produces

- A4, **10mm margins** on all four sides (configurable)
- Running header with the document title and a hairline rule
- Footer with date, and `Page X of Y`
- Styled headings, banded tables, code blocks, blockquotes and task-list glyphs
- `**Key:** value` lines above the first `---` collapse into a metadata table
- Table header rows repeat across page breaks; rows do not split mid-cell
- PDF output carries heading bookmarks and structure tags, so it is navigable
  and accessible

---

## Notes

Word automation means this is **Windows + Word desktop only**. There is no
LibreOffice or cross-platform path, and adding one would mean giving up the
fidelity that makes the output worth having.

Markdown is parsed by Markdig, built into PowerShell 7's `ConvertFrom-Markdown` —
no external parser to install.

---

## Licence

MIT. See [LICENSE](LICENSE).
