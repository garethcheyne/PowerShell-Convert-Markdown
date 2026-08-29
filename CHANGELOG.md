# Changelog — ConvertMarkdown

Versioned `yyyy.MM.dd`.

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
