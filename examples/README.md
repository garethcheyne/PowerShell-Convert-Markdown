# Examples

Sample documents for trying the converter out and for checking a change did not
break anything. Convert any of them from this folder:

```powershell
Import-Module ConvertMarkdown

# One file
Convert-Markdown .\core-features.md -Format Both -Force

# All of them, no network calls at all
Get-ChildItem *.md -Exclude README.md |
    Convert-Markdown -Format Pdf -Force -DiagramService none -NoRemoteImages
```

Generated `.docx` and `.pdf` files are gitignored, so converting in place is safe.

| File | What it covers |
| --- | --- |
| [core-features.md](core-features.md) | Plain Markdown: headings, tables, lists, code, images, task lists, a mermaid diagram, the `**Key:** value` metadata box |
| [extended-markdown.md](extended-markdown.md) | Every NextDocs construct, with a note on what each becomes on the page |
| [nextdocs-page.md](nextdocs-page.md) | A realistic NextDocs documentation page — frontmatter, callouts, steps, tabs, a diff, a restricted section |
| [styled.md](styled.md) | The same content under a custom style file, showing what the palette changes |
| [edge-cases.md](edge-cases.md) | Awkward but valid input — steps past ten, every HTTP method, gallery variants, a deliberately missing image, unicode, nested blocks |

These files are also the test corpus. `..\tests\Invoke-Tests.ps1` asserts that every
block type, callout type and inline directive appears somewhere in this folder, so
adding a feature means adding it to an example here.

> `edge-cases.md` references one image that does not exist, on purpose. It converts
> with a warning and an `[image not found: ...]` placeholder, and the suite asserts
> `ImagesMissing` is exactly 1.

## Things worth looking at

**Frontmatter.** `nextdocs-page.md` and `extended-markdown.md` both open with YAML
frontmatter. `title` becomes the H1 — and so the running header and the `.docx`
Title property — `excerpt` becomes the standfirst, and the rest becomes the
key/value box.

**Restricted sections.** `nextdocs-page.md` contains a `!variant!# Platform-Ops`
block. Convert it twice to see both behaviours:

```powershell
Convert-Markdown .\nextdocs-page.md -Format Pdf -Force
Convert-Markdown .\nextdocs-page.md -OutputPath .\nextdocs-page-public.pdf -Format Pdf -Force -ExcludeRestricted
```

The first prints the section behind a dashed banner and warns you that it did.
The second leaves it out entirely — and strips it before diagrams are rendered, so
nothing inside it reaches a rendering service either.

**Fence protection.** `extended-markdown.md` ends with a section showing the raw
syntax inside fenced code. None of it is converted. That is deliberate: a document
that *documents* this syntax has to keep its examples as text.

**Offline conversion.** `core-features.md` has a mermaid diagram, which is sent to
`mermaid.ink` to be rendered. Use `-DiagramService none` to skip it and leave a
placeholder, and `-NoRemoteImages` to skip images referenced by URL. Together they
give a conversion that makes no network call at all.

## Styling

`styled.md` is converted against [style.example.json](../style.example.json) in the
repository root:

```powershell
Convert-Markdown .\styled.md -StyleSheet ..\style.example.json -Format Pdf -Force
```

Or try a one-off accent without a style file:

```powershell
Convert-Markdown .\core-features.md -Accent '#7A1F3D' -Format Pdf -Force
```

Diagrams follow the document palette by default, so a mermaid block converted with
a custom accent comes out in matching colours rather than mermaid's own.
