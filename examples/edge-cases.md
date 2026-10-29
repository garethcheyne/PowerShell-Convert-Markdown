---
title: Edge Cases
description: Awkward but valid input — the code paths the other examples never reach.
author: gareth-cheyne
publishedAt: 2026-09-03
tags: [testing, reference]
draft: false
restricted: false
---

Everything here is **valid** input that exercises a branch the other samples miss.
The test suite converts this file and asserts on the result, so it is deliberately
awkward rather than pretty.

Two things about the frontmatter above:

- It uses the aliases `description` (for `excerpt`) and `publishedAt` (for `date`).
- `draft: false` and `restricted: false` are present but must **not** appear as rows
  in the key/value box — only the true case is worth printing.

## Steps past ten

The numbering glyph changes at eleven: 1–10 are the filled discs (U+2776–277F),
11–20 the outline circles (U+2460–2473), and anything past twenty falls back to a
plain number. All three ranges appear here.

:::steps
### Step one
Filled disc.

### Step two
Filled disc.

### Step three
Filled disc.

### Step four
Filled disc.

### Step five
Filled disc.

### Step six
Filled disc.

### Step seven
Filled disc.

### Step eight
Filled disc.

### Step nine
Filled disc.

### Step ten
Last of the filled discs.

### Step eleven
First of the outline circles.

### Step twelve
Outline circle.
:::

## Every HTTP method

Each method has its own chip colour; anything unrecognised falls back to grey.

:::api GET /things
Read.
:::

:::api POST /things
Create.
:::

:::api PUT /things/:id
Replace.
:::

:::api PATCH /things/:id
Update in place.
:::

:::api DELETE /things/:id
Remove.
:::

:::api HEAD /things/:id
Headers only.
:::

:::api OPTIONS /things
Allowed methods.
:::

:::api TRACE /things
An unrecognised method, which should still render with a neutral chip.
:::

## Gallery variations

Fixed grid, `3x2`:

:::gallery 3x2
![One](./_img/sample-1.png)
![Two](./_img/sample-2.png)
![Three](./_img/sample-3.png)
:::

Single column:

:::gallery 1
![Full width](./_img/sample-1.png)
:::

A gallery holding a link-wrapped image and an image with no alt text. The picture
survives; the link around it does not.

:::gallery 2
[![Linked image](./_img/sample-2.png)](https://example.com)
![](./_img/sample-3.png)
:::

## A missing image

The next image does not exist. It must leave a visible placeholder and a warning
rather than failing the conversion — `ImagesMissing` on the result object reports
it.

![This file is deliberately absent](./_img/no-such-file.png)

## Inline directives, awkwardly

Single key :kbd[Esc], a long chord :kbd[Ctrl+Alt+Shift+F12], and a key that is
itself a plus sign :kbd[Ctrl++].

Every badge variant: :badge[default]{default}, :badge[success]{success},
:badge[warning]{warning}, :badge[error]{error}, :badge[info]{info},
:badge[purple]{purple}, and an unknown one :badge[unknown]{not-a-variant} which
falls back to the default.

A Fluent icon :#fluentui settings: is removed. A Lucide icon `:Rocket:` is
deliberately **left as written** — that pattern is too close to ordinary text such
as 10:30:00 or a 3:2:1 ratio to strip safely.

## A YouTube link

[Watch the overview](https://www.youtube.com/watch?v=dQw4w9WgXcQ) stays an ordinary
hyperlink. No thumbnail is fetched and no player is embedded, because neither means
anything on paper.

## Blocks in awkward places

A block nested inside a details panel:

:::details Open me for a nested block
The inner content is converted as markdown, so a list works here:

- one
- two
:::

A tab panel containing a table:

:::tabs
@tab With a table
| Column | Value |
| --- | --- |
| a | 1 |
| b | 2 |

@tab With a callout
> [!TIP]
> A callout inside a tab panel.
:::

## Deep file tree

:::filetree
- a/
  - b/
    - c/
      - d/
        - e/
          - deep.txt
        - sibling.txt
      - other.txt
- root.txt
:::

## Unicode and long content

:::steps
### Ünïcödé — ステップ — 中文
Body text with an em dash —, a checkmark ✓, and a non-breaking sequence.

### A deliberately long step title that will certainly wrap onto a second line in the rendered document
Wrapping in the title cell must not push the number disc out of alignment.
:::

## Empty and minimal blocks

A details block with no summary:

:::details
Falls back to the label "Details".
:::

A filetree with a single entry:

:::filetree
- lonely.txt
:::

## Restricted section

!variant!# Docs-Admins
This section is role-gated and balanced. `-ExcludeRestricted` removes it; without
that switch it prints behind a dashed banner and the conversion warns.
!endvariant!

Public text after the restricted block.

## Syntax that must stay literal

````markdown
:::steps
### Not a real step
:::

> [!NOTE]
> Not a real callout.

!variant!# Not-A-Real-Role
Not really restricted.
!endvariant!

Press :kbd[Ctrl+S] and check :badge[status]{info}.
````

And inline: `:::api GET /x`, `:kbd[Ctrl+S]`, `:badge[beta]{warning}`, `!variant!`.
