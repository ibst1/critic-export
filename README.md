# critic-export

Export Markdown manuscripts with [CriticMarkup](https://github.com/CriticMarkup/CriticMarkup-toolkit) and Zotero citations to **Word with real tracked changes and comments**, and to **PDF**, **HTML** and any other format [Pandoc](https://pandoc.org) can write. Obsidian syntax is converted along the way, so notes export cleanly straight from your vault.

Write and revise in plain text (for example in [Obsidian](https://obsidian.md)), cite with Zotero, and hand co-authors or editors a Word file where every `{++insertion++}`, `{--deletion--}` and `{>>comment<<}` shows up as a tracked change or a Word comment they can accept, reject or answer.

## Features

- **Word**: CriticMarkup becomes tracked changes and Word comments, attributed to you or to the reviewer named in the markup.
- **Round trip**: import a reviewer's tracked changes and comments from Word back into your Markdown as CriticMarkup, with their name.
- **PDF and HTML**: changes shown in colour with comments as footnotes, or a clean copy with all changes accepted.
- **Citations**: resolves `[@citekey]` citations against a Better BibTeX export, in any citation style installed in Zotero.
- **Zotero Integration fixes**: locators such as `{p. 13}` are labelled in the document's language (`s. 13`, `S. 13`, `p. 13` …) instead of always printing "p.".
- **Obsidian-aware**: embedded notes and images (`![[…]]`), wikilinks, `%%comments%%`, `==highlights==`, callouts, tags and block IDs are converted or removed. Private `%%comments%%` never leave the vault.
- **Numbered figures, tables and equations** with cross-references (`@fig:id`), in the document's language. No pandoc-crossref needed; both pandoc-crossref and Quarto syntax work.
- **Page breaks** (`\newpage` or Quarto's `{{< pagebreak >}}`) that work in Word, ODT, PDF, HTML and Typst, not only in PDF.
- **Private notes stay private**: `<!-- HTML comments -->` are removed, so they cannot travel hidden inside an HTML or EPUB file.
- **Any format**: ODT, EPUB, LaTeX, Typst, RTF, Markdown and every other Pandoc output format.
- **Check before sending**: unknown citation keys and labels, missing images and unbalanced CriticMarkup, without exporting.
- **Word and character count** of the clean text, leaving out comments, citations and the reference list.
- **Per-project settings** in a `critic-export.toml` next to the manuscript.
- **Timestamped exports** in a subfolder, so earlier versions are never overwritten.
- Runs from the command line or from inside Obsidian. The source file is never modified.
- A single Python file with no dependencies beyond the standard library.

## What happens to CriticMarkup

| Markup | Word (`track`) | HTML/PDF (`markup`) | Any format (`accept`) |
|---|---|---|---|
| `{++added++}` | tracked insertion | blue, underlined | kept |
| `{--removed--}` | tracked deletion | red, struck through | removed |
| `{~~old~>new~~}` | deletion + insertion | both, coloured | "new" kept |
| `{>>note<<}` | Word comment | footnote "Comment: …" | removed |
| `{==text==}{>>note<<}` | comment anchored to "text" | footnote after "text" | "text" kept |

### Who made a change

By default, changes and comments are attributed to the `author` in your settings. To attribute them to someone else, name them with `@`:

| Markup | Meaning |
|---|---|
| `{>>@Anna: Is this right?<<}` | a comment by Anna |
| `{==word==}{>>@Anna: Check<<}` | Anna's comment on "word" |
| `{++new text++}{>>@Anna<<}` | an insertion by Anna (the same works after `{--…--}` and `{~~…~>…~~}`) |

In Word these become Anna's tracked changes and comments; in HTML and PDF her comments are labelled "Anna:". This is also how [imported reviews](#bringing-a-review-back-from-word) are written. Names without `@` (`{>>Anna: …<<}`) are recognised only for people listed in `reviewers` in the `[critic]` settings, so that a comment such as `{>>Source: …<<}` is not mistaken for a name.

Defaults: `track` for Word, `markup` for HTML, `accept` for everything else. Change them in the settings file or per run with `--critic`. `markup` works for HTML, EPUB, PDF and LaTeX. Other formats are always exported with changes accepted.

## Requirements

- **Python 3.11** or later
- **Pandoc 3.x**, on PATH or configured in the settings file
- **For PDF**: a LaTeX distribution with `xelatex`, such as [TinyTeX](https://yihui.org/tinytex/) or [MiKTeX](https://miktex.org). R users can install TinyTeX with `install.packages("tinytex"); tinytex::install_tinytex()`. Missing LaTeX packages are fetched automatically on the first PDF export.
- **For citations** (optional): [Zotero](https://www.zotero.org) with the [Better BibTeX](https://retorque.re/zotero-better-bibtex/) plugin
- **For Obsidian** (optional): the [Zotero Integration](https://github.com/mgmeyers/obsidian-zotero-integration) and [Shell commands](https://github.com/Taitava/obsidian-shellcommands) plugins

## Installation

1. Download or clone this repository to a permanent location, for example `C:\Tools\critic-export` or `~/tools/critic-export`.
2. **Windows: double-click `critic-export.cmd`.** It is the only file you start; the program itself is in the `app` folder. The first time, it creates `settings.toml` from `settings.example.toml`, starts the [Obsidian hotkeys](#with-autohotkey-windows), and offers to start them at every login and to add the [Explorer right-click menu](#explorer-context-menu-windows). Edit `settings.toml` afterwards (see [Configuration](#configuration)).

   **macOS and Linux:** run `python3 app/critic_export.py --show-settings` once to create `settings.toml`, and use `python3 app/critic_export.py` wherever this README says `critic-export`.

3. **Windows: type `critic-export` from any folder.** Add the critic-export folder to your PATH, so `critic-export manuscript.md` works in any terminal. To add it for your user:

   ```powershell
   $p = "C:\Tools\critic-export"
   [Environment]::SetEnvironmentVariable("Path", [Environment]::GetEnvironmentVariable("Path","User") + ";$p", "User")
   ```

   Open a new terminal afterwards, and restart programs that should find it (Obsidian). Check with `critic-export --version`.

## Setting up Zotero

**Export your library automatically.** In Zotero, right-click *My Library*, choose *Export Library…*, select **Better CSL JSON** and tick **Keep updated**. Save the file somewhere permanent and enter its path as `bibliography` in `settings.toml`. Better BibTeX now updates the file whenever your library changes.

**Insert citations in Obsidian.** In the Zotero Integration settings, add a *Citation Format* with output format **Pandoc**. Its command opens Zotero's citation picker and inserts citations like `[@dawkinsBlindWatchmaker1986, {p. 13}]`. Assign it a hotkey for convenience.

**Several bibliographies.** If you cite from several Zotero libraries, such as your own and a group library, export each one with *Keep updated* and list them all:

```toml
[citations]
bibliography = ["~/Zotero/my-library.json", "~/Zotero/group-*.json"]
```

Glob patterns pick up every matching file. If the same citation key occurs in more than one file, the last file listed wins and critic-export warns you, since Pandoc would otherwise pick one silently. A document can also list its own files in its YAML header (`bibliography: [a.json, b.bib]`), which replaces the settings for that document.

You can also type citations by hand: `[@key]`, `[@key, 13]`, `[@key1; @key2, 13–15]`. See Pandoc's [citation syntax](https://pandoc.org/MANUAL.html#citation-syntax).

To place the reference list somewhere other than the end of the document:

```markdown
## References

::: {#refs}
:::
```

## Usage

```
critic-export manuscript.md                       # Word, as set in settings
critic-export manuscript.md --to pdf html         # PDF and HTML
critic-export manuscript.md --to docx pdf html    # all three
critic-export manuscript.md --to pdf --critic markup   # review copy as PDF
critic-export manuscript.md --to odt epub         # any Pandoc format
critic-export manuscript.md --style apa           # another citation style
critic-export manuscript.md --open                # open the result when done
critic-export --list-styles                       # styles installed in Zotero
```

By default, files go to a `critic-exports` folder next to the manuscript and are named like it (`manuscript.docx`), so each export replaces the previous one. To keep every export, set `filename = "{stem}@{timestamp}"` in the settings: files are then named like `manuscript@261001T142605.docx`, and all files from one run share the same timestamp.

| Option | Meaning |
|---|---|
| `--to FORMAT …` | one or more formats: `docx`, `pdf`, `html`, `odt`, `epub`, `latex`, `typst`, `rtf`, `md`, or any other Pandoc writer |
| `--critic MODE` | `track`, `markup` or `accept` for all formats in this run |
| `--style STYLE` | Zotero style name (see `--list-styles`) or path to a `.csl` file |
| `--bibliography FILE …` | one or more bibliography files or glob patterns for this run |
| `-o FILE` | exact output file (only with a single format, or with `--import`) |
| `--author NAME` | name on tracked changes and comments |
| `--title TEXT` | document title, when the YAML header has none |
| `--doc-author NAME` | document author, when the YAML header has none (repeat for several) |
| `--save-metadata` | also write `--title` / `--doc-author` into the manuscript's YAML header |
| `--no-prompt` | never ask for a missing title or author |
| `--reference-doc FILE` | Word or ODT template for fonts and heading styles |
| `--zotero-fields` | make citations in Word Zotero fields (see [Zotero citations in Word](#zotero-citations-in-word)); `--no-zotero-fields` turns the setting off for one run |
| `--obsidian MODE` | `auto` (default), `on` or `off`: convert Obsidian syntax |
| `--settings FILE` | use another settings file for this run |
| `--show-settings` | show the settings file in use and its values |
| `--update-settings` | add new settings from `settings.example.toml` to `settings.toml` |
| `--open` | open the exported files afterwards |
| `--import DOCX` | bring the tracked changes and comments in a reviewed Word file back as CriticMarkup |
| `--in-place` | with `--import`: update the manuscript itself (the previous version is kept) |
| `--check` | report problems without exporting |
| `--count` | count words and characters in the clean text |
| `--install-context-menu` | add the Explorer right-click menu (Windows) |
| `--uninstall-context-menu` | remove it |
| `--pause-on-issues` | keep the console open on errors or warnings |
| `--keep` | keep the intermediate Markdown file, for debugging |

## Zotero citations in Word

By default, citations in Word are plain text: they look right, but the Zotero Word plugin cannot change them. With `--zotero-fields` (or `zotero_fields = true` in the `[docx]` section of the settings), they become Zotero citation fields instead, with the reference list as a Zotero bibliography:

```
critic-export manuscript.md --to docx --zotero-fields
```

Click *Refresh* on Zotero's tab in Word, and Zotero takes over: you and your co-authors can change the citation style, edit citations and add new ones with the plugin. The document preferences (citation style and language) are set to match the export.

- Zotero must be running when you export, since the item data comes from Better BibTeX. If it is not, the export still works, with plain-text citations and a warning.
- Citation keys that are not in your Zotero library stay plain text, and are listed.
- An in-text citation (`@dawkins1986 says`) keeps the author's name as text, and the field shows only the year, as Zotero does it.
- Co-authors who edit the citations need Zotero and its Word plugin. Others can read and edit the text as usual.

## Bringing a review back from Word

Send the Word file to a co-author or editor. When it comes back with tracked changes and comments, import them into your manuscript:

```
critic-export manuscript.md --import manuscript-reviewed.docx
```

This writes `manuscript@import-261005T091500.md` next to the manuscript, with every change and comment from other people turned into CriticMarkup, with their name:

```markdown
Evolution finds {--the--}{>>@Anna Lund<<} solutions{++ over++}{>>@Anna Lund<<} again …
{==Convergence==}{>>@Anna Lund: Is this the right word?<<} is common …
```

Review it in Obsidian and accept or reject each change by editing the markup. Use `--in-place` to update the manuscript itself instead; the previous version is then saved in the export folder.

- Changes are matched word by word against your manuscript, so your Markdown, citations, links and Obsidian syntax are kept. A change inside a word covers the whole word: `{~~design~>designs~~}`.
- Your own tracked changes and comments (by the `author` in your settings, or any name in `[import] own_authors`) are already in the manuscript and are skipped. So are changes and comments from an earlier import, so a file can go back and forth for several rounds without anything being added twice.
- A change that cannot be placed exactly, for example inside a rendered citation or a figure number, is added as a comment at the nearest position, and listed when the import finishes. Nothing is lost.
- Changes in footnotes and in the generated reference list are not imported; their number is reported.
- A citation the reviewer inserted with the Zotero Word plugin becomes a citation key, `{++ See [@gould1989, 5].++}`, if the item is in your bibliography. It is matched by citation key, DOI, Zotero item (when Zotero is running), or title and year, so this also works when the reviewer cites from their own Zotero library. Citations that cannot be matched are imported as text and listed. Citations typed by hand stay text.

## Checking a manuscript

```
critic-export manuscript.md --check
```

reports, without exporting anything:

- citation keys that are not in your bibliography
- references to figure, table, equation and section labels that do not exist, and labels used twice
- images and embedded notes that cannot be found
- CriticMarkup or `%%comments%%` that are opened but not closed

and notes how much review markup is still in the text, which is useful before sending a clean copy. Each problem is given with its line number where possible. The exit code is 1 when there are problems, so the check can be used in scripts.

## Counting words and characters

```
critic-export manuscript.md --count
```

counts words and characters (with and without spaces) in the text as it will be published: changes accepted, comments removed, citations left out. Headings, figure and table captions, tables and footnotes are listed separately. Sections headed "References", "Bibliography", "Referenser" or "Litteratur" are left out, so a reference list written by hand is not counted; change the list with `exclude_sections` in the `[count]` settings.

```
                Words   Characters   without spaces
Main text       5 844       38 186           32 388
Headings           82          511              444
Total           5 926       38 697           32 832
```

## Running from Obsidian

Install the **Shell commands** community plugin and add a command such as:

```
critic-export {{file_path:absolute}} --to docx --open
```

This uses `critic-export.cmd` on PATH (see [Installation](#installation)). Without it, give the full path: `"C:/Tools/critic-export/critic-export.cmd" {{file_path:absolute}} --to docx --open`.

`{{file_path:absolute}}` is replaced with the note that is currently open. Give the command an alias like "Export to Word", then assign it a hotkey under *Settings → Hotkeys*, or run it from the command palette. Add more commands for other formats, for example `--to pdf --critic markup` for a PDF review copy.

If Obsidian cannot find Python, use its full path instead of `python`. Find it with `where python` (Windows) or `which python3` (macOS/Linux).

Obsidian's own export and plugins such as Enhancing Export are not needed.

### With AutoHotkey (Windows)

If you use [AutoHotkey v2](https://www.autohotkey.com), `app/critic-export-obsidian.ahk` exports the note that is open in Obsidian with a hotkey, without any Obsidian plugin:

| Hotkey | Export |
|---|---|
| `Ctrl+Alt+W` | Word |
| `Ctrl+Alt+P` | PDF, review copy with changes and comments |
| `Ctrl+Alt+Shift+P` | PDF, clean |
| `Ctrl+Alt+H` | HTML |

The hotkeys only work while Obsidian is the active window. The script saves the note, runs the export in the background and shows a notification, or a message box with Pandoc's output if there were errors or warnings. Double-clicking `critic-export.cmd` starts it (and can add it to the Startup folder). Change the hotkeys and options at the top of the file.

It runs `critic-export.cmd` in the folder above it, so PATH is not needed. It finds the note from Obsidian's window title ("Note - Vault - Obsidian 1.x") and the vault's location in Obsidian's own settings, so two limitations follow: if two notes in the vault share a name, it cannot tell which one is open and says so, and a plugin that changes the window title stops it from working. The script's messages are in Swedish; edit the strings to translate them.

## Figures, tables and cross-references

Give a figure, table, equation or section a label, and refer to it with `@`:

```markdown
![Convergent evolution of the camera eye.](figures/eye.png){#fig:eye}

| Species | Eye    |
|---------|--------|
| Octopus | camera |

: Comparison of eye types {#tbl:eyes}

$$ p = 1 - (1 - \mu)^n $$ {#eq:prob}

# Methods {#sec:methods}

The eyes are strikingly similar (@fig:eye; see also @tbl:eyes and @eq:prob).
```

This gives "Figure 1: Convergent evolution…" as the caption and "figure 1", "table 1", "equation 1" and "section 1" in the text, each linked to its target. Numbers update automatically when you move things around, and they are the same in every format.

| You write | You get |
|---|---|
| `@fig:eye` | figure 1 |
| `@Fig:eye` | Figure 1 (capitalised, for the start of a sentence) |
| `[-@fig:eye]` | 1 (number only) |
| `[@fig:eye; @fig:ear]` | figures 1 and 2 |
| `[@fig:eye; @tbl:eyes]` | figure 1, table 1 |

Labels start with `fig`, `tbl`, `eq` or `sec`, followed by `:` (pandoc-crossref style) or `-` (Quarto style): `{#fig:eye}` and `{#fig-eye}` both work. Only labelled items are numbered. A reference to a missing label prints "??" and a warning.

The words follow the document's language (`lang` in the YAML header), with English, Swedish, Danish, Norwegian, German, French and Spanish built in. Change them in the `[crossref]` section of the settings file, for example to "Fig." and "figs.".

In Obsidian, an embedded image becomes a numbered figure when it has a caption and a label: `![[eye.png|Convergent evolution of the camera eye|400]]{#fig:eye}`. Obsidian shows the `{#fig:eye}` as text after the image, but the export handles it.

## Title and author

The title and author on the title page, and in the document properties of the PDF, Word or HTML file, come from the YAML header:

```yaml
---
title: My chapter
author: Your Name        # or a list: [Anna Andersson, Bo Berg]
---
```

When the header has no `author:`, `document_author` in the `[general]` section of the settings is used, so on your own computer you set your name once. When a PDF, Word or HTML export still lacks a title or author and runs in a console, critic-export asks for them and offers to save the answers in the YAML header. The suggested title is the H1 heading the note opens with, or else the file name; press Enter to accept it. An opening H1 that repeats the title is left out of the export, so the title is not printed twice. `--title` and `--doc-author` give them on the command line, and `--no-prompt` turns the question off (a note is printed instead). The Obsidian hotkeys ask for a missing title in a dialog of their own, with the same suggestion filled in.

## Page breaks and comments

A page break on a line of its own becomes a real page break in every format that has pages: Word, ODT, PDF and Typst, and HTML/EPUB when printed. Pandoc on its own keeps `\newpage` only in PDF output.

```markdown
Title page

\newpage

Chapter 1 …
```

`\newpage`, `\pagebreak`, `\clearpage` and Quarto's `{{< pagebreak >}}` all work.

`<!-- HTML comments -->` are removed before export, inline or spanning several lines. Word would drop them anyway, but HTML and EPUB would otherwise carry them along, hidden in the page source. A comment used to separate two lists still keeps them apart, and comments inside code are left alone. Both features can be switched off in the `[syntax]` section of the settings file.

## Explorer context menu (Windows)

Right-click a Markdown file in Windows Explorer and export it from a *critic-export* submenu. Double-clicking `critic-export.cmd` offers to add it, or run:

```
critic-export --install-context-menu
```

The menu is added for `.md`, `.qmd`, `.rmd` and `.markdown` files, for the current user only, so no administrator rights are needed. On Windows 11 it appears under *Show more options* (or hold Shift while right-clicking).

Menu items and file types are set in the `[context_menu]` section of `settings.toml`, for example in your own language:

```toml
[context_menu]
title = "critic-export"
extensions = [".md", ".qmd", ".rmd"]
items = [
  { label = "Word", args = "--to docx --open" },
  { label = "PDF (review copy)", args = "--to pdf --critic markup --open" },
]
```

Run `--install-context-menu` again after changing the section or moving the script folder. `--uninstall-context-menu` removes the menu.

Each item runs the export in a console window. The window closes by itself when everything went well, and stays open if there were errors or warnings (missing citations, duplicate keys and so on) until you press Enter.

## Obsidian syntax

When the document is inside an Obsidian vault (a folder containing `.obsidian`), Obsidian syntax is converted before export. Links are resolved the way Obsidian resolves them: by path, relative to the note, or by file name anywhere in the vault.

| Obsidian | Exported as |
|---|---|
| `![[image.png]]` | image, found anywhere in the vault |
| `![[image.png\|400]]`, `![[image.png\|400x300]]` | image with width (and height) |
| `![[image.png\|Caption\|400]]{#fig:id}` | numbered figure with caption (see above) |
| `![[Note]]` | the note's content, inserted in place |
| `![[Note#Heading]]`, `![[Note#^block-id]]` | that section or block only |
| `[[Note]]`, `[[Note\|Alias]]`, `[[Note#Heading]]` | plain text: "Note", "Alias", "Note > Heading" |
| `%%comment%%` | removed |
| `==highlight==` | highlighted text |
| `> [!note] Title` | block quote with "Title" in bold |
| a list right after a line of text | a list (Pandoc on its own needs a blank line before and after) |
| `#tag` | removed (lines with only tags disappear) |
| `^block-id` | removed |

Embedded notes are processed like the note itself, so their citations, CriticMarkup and embeds work too. This makes it possible to write a chapter or thesis as separate notes and assemble them in one "master" note of `![[…]]` embeds. Code blocks and inline code are never changed. Turn conversion off with `--obsidian off`, or always on with `--obsidian on`.

## Citation styles

Any style installed in Zotero can be used by its file name without `.csl`:

```
critic-export --list-styles
critic-export manuscript.md --style vancouver
```

Set the default with `style` in `settings.toml`. A single document can choose its own style in its YAML header, which takes precedence over the settings file. `--style` takes precedence over both.

```yaml
---
title: My chapter
lang: sv-SE
csl: apa
---
```

Install more styles in Zotero (*Settings → Cite → Styles → +*) or download them from the [Zotero Style Repository](https://www.zotero.org/styles). A `.csl` file anywhere on disk also works: `--style path/to/style.csl`.

**Language.** Set `lang` in the document's YAML header (`sv-SE`, `de-DE`, `en-GB` …). The citation style then uses that language's terms ("och" for "and", "s." for "p." and so on).

**Page labels.** Some styles, including Chicago author-date, never print a page label: they give `(Dawkins 1986, 13)` rather than `(Dawkins 1986, p. 13)`. That is the style's rule, not a bug. Choose another style if you want labels.

## Locators from Zotero Integration

Zotero Integration writes locators in English and in braces, e.g. `[@key, {p. 13}]`. The braces stop Pandoc from interpreting them, so "p." would be printed as is, whatever the document's language. critic-export handles this before exporting:

- **Page numbers** (`{p. 13}`, `{pp. 13–15}`) become `13` and `13–15`. Pandoc reads a bare number as a page, and the citation style labels it in the document's language.
- **Other locators** (`{chap. 3}`, `{para. 4}` …) lose their braces. In English documents this is all that is needed. For other languages, map the English abbreviations to your language's terms in the `[locators]` section of `settings.toml`.

## Configuration

All settings live in `settings.toml` in the critic-export folder, next to `critic-export.cmd`. Settings are not written elsewhere, so the tool and its settings can be moved, backed up or synced as one folder. If `settings.toml` is missing, it is created from `settings.example.toml` on the first run. `settings.toml` is listed in `.gitignore`, so your personal paths are never committed, and updating the tool never overwrites it.

`--settings FILE` reads another file for a single run, and `--show-settings` prints which file is in use and the resulting values.

New versions of critic-export may add settings. They work with their defaults right away, but are not listed in your existing `settings.toml`. `--show-settings` tells you when that is the case, and `--update-settings` adds them, with their documentation, from `settings.example.toml`. Your own values and comments are never changed, and the previous file is kept as `settings.toml.bak`.

Every setting is optional and documented in [`settings.example.toml`](settings.example.toml). The most important ones:

```toml
[general]
author = "Your Name"

[citations]
bibliography = "~/Zotero/library.json"
style = "chicago-author-date"

[critic]
comment_label = "Comment"
```

### Project settings

Different manuscripts often need different settings: another citation style, a publisher's Word template, a co-author's bibliography. Put a `critic-export.toml` in the manuscript's folder, or in any folder above it, with just the settings that differ:

```toml
# critic-export.toml for the anthology chapter
[citations]
style = "chicago-author-date"
bibliography = ["~/Zotero/library.json", "anthology-group.json"]

[docx]
reference_doc = "publisher-template.docx"
```

It applies to every document in that folder and below. Relative paths are relative to the project file. `--show-settings` with a document name shows which project file is used:

```
critic-export manuscript.md --show-settings
```

Options are applied in this order, later ones winning: built-in defaults, `settings.toml`, `critic-export.toml`, the document's YAML header (`bibliography`, `csl`), command-line options.

Versions before 1.5 kept settings in the user config folder (`%APPDATA%\critic-export` on Windows, `~/.config/critic-export` elsewhere). An existing file there is moved to the script folder automatically.

## Limitations

- Comments become a single paragraph in Word. Paragraph breaks become line breaks, and block quotes become plain text.
- Links inside comments become plain text in Word ("name (url)"), because Word cannot display links in comments written by Pandoc. In HTML and PDF they work normally.
- CriticMarkup must not span block structures such as a heading and the following paragraph. Changes spanning several plain paragraphs are fine.
- A comment standing alone between paragraphs gets an empty paragraph of its own in Word.
- Embedded PDFs, audio and video become links, not embedded content. Dataview queries and other plugin syntax are exported as written.
- `--import` matches words, so it works best when the reviewer edits the text rather than rewriting whole paragraphs. A paragraph that was rewritten from scratch becomes one large replacement or a comment.
- `--import` works on the manuscript file itself; changes in text that came from an embedded note (`![[…]]`) are added as comments, not to the embedded note.
- Quarto code execution, callout blocks and shortcodes are not supported. `.qmd` files are read as plain Pandoc Markdown (Quarto-style cross-references do work).
- Cross-references cover figures, tables, equations and sections. Subfigures, listings and theorem environments are not numbered.

## How it works

critic-export first inserts embedded notes, then removes Obsidian comments, fixes locators, converts CriticMarkup and finally converts the remaining Obsidian syntax, with code protected throughout. CriticMarkup becomes Pandoc's own markup for tracked changes and comments (spans with the classes `insertion`, `deletion`, `comment-start` and `comment-end`), which Pandoc's Word writer turns into real revisions and comments. A second embedded Lua filter numbers labelled items and resolves `@fig:` references before Pandoc's citation processor runs. For HTML, EPUB and PDF, a small embedded Lua filter renders the same spans in colour and turns comments into footnotes. Citations are resolved by Pandoc's built-in citeproc. With `--zotero-fields`, another embedded Lua filter then wraps each rendered citation in a Zotero field, with item data from Better BibTeX's JSON-RPC interface.

## Development

The tests use [pytest](https://pytest.org). Tests that go through Pandoc are skipped when it is not installed.

```
python -m pip install pytest
python -m pytest
```

They run automatically on Windows, macOS and Linux for every push and pull request (see `.github/workflows/tests.yml`).

## License

[MIT](LICENSE)
