#!/usr/bin/env python3
"""
critic-export: export Markdown with CriticMarkup and Zotero citations to
Word (tracked changes and comments), PDF, HTML and any other Pandoc format.
Understands Obsidian syntax (embeds, transclusion, wikilinks, %%comments%%,
==highlights==, callouts, tags, block IDs) and numbers figures, tables,
equations and sections with cross-references (@fig:id, @tbl:id ...).

    python critic_export.py manuscript.md
    python critic_export.py manuscript.md --to docx pdf html
    python critic_export.py manuscript.md --to pdf --critic markup --style apa
    python critic_export.py manuscript.md --to odt epub latex
    python critic_export.py --list-styles

See README.md for setup and settings.toml for configuration.
The source file is never modified.
"""
from __future__ import annotations

import argparse
import difflib
import glob
import json
import os
import random
import re
import shutil
import subprocess
import sys
import tempfile
import tomllib
import zipfile
from datetime import datetime, timezone
from itertools import count
from pathlib import Path
from urllib.parse import unquote

__version__ = "2.0.0"

CRITIC_MODES = ("track", "markup", "accept")
LATEX_ENGINES = {"xelatex", "lualatex", "pdflatex", "latexmk", "tectonic"}

# Output format name -> (Pandoc writer, file extension). Anything not listed
# is passed to Pandoc as is, with the format name as extension.
FORMAT_ALIASES = {
    "docx": ("docx", "docx"),
    "pdf": (None, "pdf"),          # writer depends on the PDF engine
    "html": ("html5", "html"),
    "odt": ("odt", "odt"),
    "epub": ("epub3", "epub"),
    "latex": ("latex", "tex"),
    "tex": ("latex", "tex"),
    "typst": ("typst", "typ"),
    "rtf": ("rtf", "rtf"),
    "pptx": ("pptx", "pptx"),
    "md": ("markdown-citations", "md"),
    "markdown": ("markdown-citations", "md"),
    "gfm": ("gfm", "md"),
    "plain": ("plain", "txt"),
    "txt": ("plain", "txt"),
    "rst": ("rst", "rst"),
    "asciidoc": ("asciidoc", "adoc"),
    "org": ("org", "org"),
    "mediawiki": ("mediawiki", "wiki"),
    "jats": ("jats", "xml"),
    "icml": ("icml", "icml"),
}

# Formats that can show CriticMarkup. Word shows real tracked changes; the
# others show coloured insertions/deletions and comments as footnotes.
TRACK_FORMATS = {"docx"}
MARKUP_FORMATS = {"html", "epub", "pdf", "latex", "tex"}

# Formats whose export asks for a missing title or author.
ASK_METADATA = {"pdf", "docx", "html"}

# PDF engine -> Pandoc writer used to produce the PDF.
PDF_ENGINE_WRITERS = {
    "typst": "typst", "context": "context", "weasyprint": "html5",
    "wkhtmltopdf": "html5", "prince": "html5", "pagedjs-cli": "html5",
    "pdfroff": "ms", "groff": "ms",
}

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

DEFAULTS: dict = {
    "general": {
        "author": "",
        "document_author": "",
        "formats": ["docx"],
        "output_dir": "critic-exports",
        "filename": "{stem}",
        "timestamp": "%y%m%dT%H%M%S",
        "open_after_export": False,
    },
    "tools": {
        "pandoc": "",
        "pdf_engine": "xelatex",
    },
    "citations": {
        "bibliography": "",
        "style": "",
        "zotero_data_dir": "",
        "link_citations": True,
    },
    "critic": {
        "docx": "track",
        "html": "markup",
        "pdf": "accept",
        "epub": "accept",
        "comment_label": "Comment",
        "other": "accept",
        "reviewers": [],
    },
    "obsidian": {
        "enabled": "auto",
        "remove_tags": True,
        "max_embed_depth": 5,
    },
    "crossref": {
        "enabled": True,
        "figure": "", "figure_ref": "", "figure_refs": "",
        "table": "", "table_ref": "", "table_refs": "",
        "equation": "", "equation_ref": "", "equation_refs": "",
        "section": "", "section_ref": "", "section_refs": "",
        "and": "", "separator": "",
    },
    "syntax": {
        "page_breaks": True,
        "remove_html_comments": True,
    },
    "context_menu": {
        "title": "critic-export",
        "extensions": [".md", ".qmd", ".rmd", ".markdown"],
        "items": [
            {"label": "Word", "args": "--to docx --open"},
            {"label": "PDF (review copy)", "args": "--to pdf --critic markup --open"},
            {"label": "PDF (clean)", "args": "--to pdf --critic accept --open"},
            {"label": "HTML", "args": "--to html --open"},
            {"label": "Word, PDF and HTML", "args": "--to docx pdf html"},
        ],
    },
    "import": {"own_authors": []},
    "count": {"exclude_sections": ["References", "Bibliography", "Referenser", "Litteratur",
                                   "Litteraturförteckning", "Källor", "Literatur", "Références"]},
    "docx": {"reference_doc": "", "zotero_fields": False},
    "html": {"css": ""},
    "locators": {},
}


APP_DIR = Path(__file__).resolve().parent          # this script and the hotkey script
HOME_DIR = APP_DIR.parent                          # critic-export.cmd and the settings
START_CMD = HOME_DIR / "critic-export.cmd"
HOTKEY_SCRIPT = APP_DIR / "critic-export-obsidian.ahk"
LOCAL_SETTINGS = HOME_DIR / "settings.toml"
EXAMPLE_SETTINGS = HOME_DIR / "settings.example.toml"


def legacy_settings() -> Path:
    """Where versions before 1.5 kept settings (user config folder)."""
    if sys.platform == "win32":
        return Path(os.environ.get("APPDATA", Path.home())) / "critic-export" / "settings.toml"
    return Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / "critic-export" / "settings.toml"


def ensure_local_settings() -> Path | None:
    """Settings live in the critic-export folder, so they travel and sync with it.
    If settings.toml is missing, move an old one from the user config folder
    (versions before 1.5), otherwise create it from settings.example.toml."""
    if LOCAL_SETTINGS.exists():
        return LOCAL_SETTINGS
    legacy = legacy_settings()
    try:
        if legacy.exists():
            shutil.move(str(legacy), LOCAL_SETTINGS)
            try:
                legacy.parent.rmdir()                # only if now empty
            except OSError:
                pass
            print(f"Moved settings from {legacy} to {LOCAL_SETTINGS}", file=sys.stderr)
            return LOCAL_SETTINGS
        if EXAMPLE_SETTINGS.exists():
            shutil.copyfile(EXAMPLE_SETTINGS, LOCAL_SETTINGS)
            print(f"Created {LOCAL_SETTINGS} from {EXAMPLE_SETTINGS.name}. "
                  "Edit it to set your bibliography and citation style.", file=sys.stderr)
            return LOCAL_SETTINGS
    except OSError as e:
        print(f"Warning: could not create {LOCAL_SETTINGS}: {e}. Using built-in defaults.",
              file=sys.stderr)
        return None
    print(f"Warning: neither {LOCAL_SETTINGS.name} nor {EXAMPLE_SETTINGS.name} found in "
          f"{HOME_DIR}. Using built-in defaults.", file=sys.stderr)
    return None


TOML_HEADER = re.compile(r"^\s*\[([^\[\]]+)\]\s*(?:#.*)?$")
TOML_KEY = re.compile(r'^\s*("[^"]+"|[A-Za-z0-9_-]+)\s*=')


def _toml_sections(text: str) -> list[tuple[str, int, int]]:
    """(name, start line, end line) of each [section] in a TOML text."""
    lines = text.splitlines()
    heads = [(i, m.group(1).strip()) for i, line in enumerate(lines)
             if (m := TOML_HEADER.match(line))]
    return [(name, i, heads[n + 1][0] if n + 1 < len(heads) else len(lines))
            for n, (i, name) in enumerate(heads)]


def _toml_key_blocks(lines: list[str]) -> list[tuple[str, list[str]]]:
    """Each key in a section body with the comment lines just above it and
    the full value (multi-line arrays included)."""
    blocks, i = [], 0
    while i < len(lines):
        m = TOML_KEY.match(lines[i])
        if not m:
            i += 1
            continue
        start = i
        while start > 0 and lines[start - 1].lstrip().startswith("#"):
            start -= 1
        end, depth = i, 0
        while True:
            code = re.sub(r'"(?:[^"\\]|\\.)*"|#.*', "", lines[end])
            depth += code.count("[") - code.count("]") + code.count("{") - code.count("}")
            if depth <= 0 or end + 1 >= len(lines):
                break
            end += 1
        blocks.append((m.group(1).strip('"'), lines[start:end + 1]))
        i = end + 1
    return blocks


def missing_settings(user_text: str, example_text: str) -> list[tuple[str, str | None, list[str]]]:
    """Sections and keys documented in the example but absent from the user's
    file: (section, key or None for a whole section, lines to add)."""
    user = tomllib.loads(user_text)
    ex_lines = example_text.splitlines()
    missing = []
    for name, start, end in _toml_sections(example_text):
        body = ex_lines[start + 1:end]
        if name not in user:
            while body and not body[-1].strip():
                body = body[:-1]
            missing.append((name, None, [ex_lines[start], *body]))
            continue
        for key, block in _toml_key_blocks(body):
            if key not in user[name]:
                missing.append((name, key, block))
    return missing


def update_settings(path: Path) -> None:
    """Add settings documented in settings.example.toml that settings.toml
    lacks, with their comments. Existing lines are never changed."""
    if not EXAMPLE_SETTINGS.exists():
        sys.exit(f"{EXAMPLE_SETTINGS} not found.")
    text = path.read_text(encoding="utf-8")
    missing = missing_settings(text, EXAMPLE_SETTINGS.read_text(encoding="utf-8"))
    if not missing:
        print(f"{path.name} already has every setting in {EXAMPLE_SETTINGS.name}.")
        return

    lines = text.splitlines()
    sections = {name: (start, end) for name, start, end in _toml_sections(text)}
    inserts: dict[int, list[str]] = {}
    appended: list[str] = []
    for section, key, block in missing:
        if key is None:
            appended += ["", *block]
        else:
            end = sections[section][1]
            while end > sections[section][0] + 1 and not lines[end - 1].strip():
                end -= 1                              # before trailing blank lines
            inserts.setdefault(end, []).extend(["", *block])
    for pos in sorted(inserts, reverse=True):
        lines[pos:pos] = inserts[pos]
    new_text = "\n".join(lines).rstrip() + "\n" + ("\n".join(appended) + "\n" if appended else "")

    try:
        tomllib.loads(new_text)
    except tomllib.TOMLDecodeError as e:
        sys.exit(f"Could not update {path.name} safely ({e}); nothing was changed.")
    backup = path.with_name(path.name + ".bak")
    shutil.copyfile(path, backup)
    path.write_text(new_text, encoding="utf-8")
    print(f"Updated {path} (previous version saved as {backup.name}). Added:")
    for section, key, _ in missing:
        print(f"  [{section}]" + (f" {key}" if key else "  (whole section)"))


def find_settings(explicit: Path | None) -> Path | None:
    """--settings if given, otherwise settings.toml in the critic-export folder."""
    if explicit:
        if not explicit.exists():
            sys.exit(f"Settings file not found: {explicit}")
        return explicit.resolve()
    return ensure_local_settings()


def load_settings(path: Path | None) -> dict:
    settings = {section: dict(values) for section, values in DEFAULTS.items()}
    if path:
        try:
            user = tomllib.loads(path.read_text(encoding="utf-8"))
        except tomllib.TOMLDecodeError as e:
            sys.exit(f"Error in {path}: {e}")
        for section, values in user.items():
            if section not in settings:
                print(f"Warning: unknown section [{section}] in {path.name}", file=sys.stderr)
                continue
            settings[section].update(values)
    return settings


def expand(p: str, base: Path | None = None) -> Path:
    """Expand ~ and environment variables; resolve relative to base."""
    path = Path(os.path.expandvars(os.path.expanduser(p)))
    if base and not path.is_absolute():
        path = base / path
    return path


# ---------------------------------------------------------------------------
# CriticMarkup conversion
# ---------------------------------------------------------------------------

# One pattern for every CriticMarkup variant. Order matters:
# highlight+comment must be tried before a lone comment.
# A change may be followed by {>>@Name<<} to say who made it.
def _by(group: str) -> str:
    return rf"(?:\{{>>\s*@(?P<{group}>[^:<>\n]{{1,60}}?)\s*<<\}})?"


CRITIC = re.compile(
    r"\{~~(?P<sub_old>.*?)~>(?P<sub_new>.*?)~~\}" + _by("sub_by") +
    r"|\{==(?P<hl>.*?)==\}\s*\{>>(?P<hl_comment>.*?)<<\}"
    r"|\{\+\+(?P<ins>.*?)\+\+\}" + _by("ins_by") +
    r"|\{--(?P<dele>.*?)--\}" + _by("del_by") +
    r"|\{>>(?P<comment>.*?)<<\}"
    r"|\{==(?P<hl_only>.*?)==\}",
    re.DOTALL,
)

# {>>@Anna: text<<} is a comment by Anna. Without @, "Anna: text" is only
# read as a name when Anna is listed in [critic] reviewers.
COMMENT_AUTHOR = re.compile(r"^\s*@([^:\n]{1,60}?)\s*:\s*(.*)$", re.DOTALL)
COMMENT_NAME = re.compile(r"^\s*([^:\n]{1,60}?)\s*:\s*(.*)$", re.DOTALL)


def comment_author(text: str, reviewers: list[str]) -> tuple[str | None, str]:
    bare = re.match(r"^\s*@([^:<>\n]{1,60}?)\s*$", text)
    if bare:                                    # a lone {>>@Name<<} carries no comment
        return bare.group(1).strip(), ""
    m = COMMENT_AUTHOR.match(text)
    if m:
        return m.group(1).strip(), m.group(2)
    m = COMMENT_NAME.match(text)
    if m and m.group(1).strip().lower() in {r.lower() for r in reviewers}:
        return m.group(1).strip(), m.group(2)
    return None, text

PARA_BREAK = re.compile(r"\n[ \t]*\n")

# A page break on a line of its own: LaTeX style or Quarto's shortcode.
PAGE_BREAK = re.compile(
    r"^[ \t]*(?:\\newpage|\\pagebreak|\\clearpage|\{\{<\s*pagebreak\s*>\}\})[ \t]*$",
    re.MULTILINE)


def mark_page_breaks(text: str) -> str:
    """Turn page breaks into an empty div that the page-break filter renders
    for each format (Pandoc would otherwise drop \\newpage outside PDF)."""
    return PAGE_BREAK.sub("\n::: pagebreak\n:::\n", text)
LIST_ITEM = re.compile(r"(?:[-*+]|\d{1,9}[.)])(?:[ \t]|$)")
LIST_START = re.compile(r"(?:[-*+]|1[.)])[ \t]+\S")
YAML_HEAD = re.compile(r"\A---\r?\n.*?\r?\n(?:---|\.\.\.)\r?\n", re.DOTALL)
MD_LINK = re.compile(r"\[([^\]]*)\]\(([^)\s]+)[^)]*\)")

# Zotero Integration writes locators in braces, e.g. [@key, {p. 13}].
# Braces stop Pandoc from parsing the locator, so "p." is printed literally
# whatever the document language.
BRACED_PAGE = re.compile(r",\s*\{pp?\.\s*(\d[^}]*)\}")
BRACED_LOCATOR = re.compile(r",\s*\{([A-Za-z]+\.?\s+[^}]*)\}")


def fix_locators(text: str, terms: dict) -> str:
    """Page locators become bare numbers, which Pandoc always reads as pages
    and the citation style labels in the document's language (p., s., S. ...).
    Other braced locators are unwrapped, and their English abbreviation is
    replaced via the [locators] table in settings, e.g. "chap." = "kap."."""
    text = BRACED_PAGE.sub(r", \1", text)

    def unwrap(m: re.Match) -> str:
        locator = m.group(1).strip()
        for english, local in terms.items():
            if locator.startswith(english + " "):
                locator = local + locator[len(english):]
                break
        return f", {locator}"

    return BRACED_LOCATOR.sub(unwrap, text)


def wrap_change(text: str, kind: str, attrs: str) -> str:
    """Wrap text in an insertion/deletion span. Spans cannot cross
    paragraphs, so multi-paragraph changes are split per paragraph.
    Pandoc drops spaces at the edges of a span, which would glue words
    together in Word ("som{++ länge++}" -> "somlänge"), so an edge space is
    kept as an escaped (non-breaking) space inside the change."""
    if not text.strip():
        return text if kind == "insertion" else ""

    def span(p: str) -> str:
        if not p.strip():
            return p
        lead = "\\ " if p[0].isspace() else ""
        trail = "\\ " if p[-1].isspace() else ""
        return f"[{lead}{p.strip()}{trail}]{{.{kind} {attrs}}}"

    return "\n\n".join(span(p) for p in PARA_BREAK.split(text))


def flatten_comment(text: str, plain_links: bool) -> str:
    """Turn a comment into a single paragraph: block quote markers are
    removed and paragraph breaks become hard line breaks.
    For Word, links become plain text "name (url)": Pandoc writes the link
    relationships of comments to the wrong part of the .docx, which makes
    Word report "Missing content" and the links do not work anyway."""
    if plain_links:
        text = MD_LINK.sub(r"\1 (\2)", text)
    lines = [re.sub(r"^\s*>\s?", "", ln) for ln in text.strip().splitlines()]
    return PARA_BREAK.sub("\\\\\n", "\n".join(lines))


def split_yaml(source: str) -> tuple[str, str]:
    m = YAML_HEAD.match(source)
    head = m.group(0) if m else ""
    return head, source[len(head):]


def has_yaml_key(head: str, key: str) -> bool:
    """True if the YAML header gives `key` a value, inline or as an indented list."""
    return re.search(rf"^{key}\s*:[ \t]*(?:[^\s#]|\r?\n[ \t]+\S)", head, re.MULTILINE) is not None


LEADING_H1 = re.compile(r"\A\s*#[ \t]+(.+?)(?:[ \t]+#+)?[ \t]*(?:\r?\n|\Z)")


def leading_h1(body: str) -> str | None:
    """The text of an H1 heading that opens the document, if there is one."""
    m = LEADING_H1.match(body)
    return m.group(1).strip() if m else None


def yaml_title(head: str) -> str | None:
    m = re.search(r"^title\s*:[ \t]*(.+?)[ \t]*$", head, re.MULTILINE)
    return m.group(1).strip("'\"").strip() if m else None


def drop_title_heading(head: str, body: str) -> str:
    """An opening H1 that repeats the title would be printed twice, once on the
    title block and once as a heading, so it is left out of the export."""
    h1, title = leading_h1(body), yaml_title(head)
    if h1 and title and h1.casefold() == title.casefold():
        return body[LEADING_H1.match(body).end():].lstrip("\r\n")
    return body


def yaml_metadata_lines(title: str | None, authors: list[str]) -> list[str]:
    def q(s: str) -> str:
        return json.dumps(s, ensure_ascii=False)        # a JSON string is a YAML string
    lines = [f"title: {q(title)}"] if title else []
    if len(authors) == 1:
        lines.append(f"author: {q(authors[0])}")
    elif authors:
        lines += ["author:", *(f"  - {q(a)}" for a in authors)]
    return lines


def add_to_yaml(head: str, lines: list[str]) -> str:
    """Add lines to a YAML header, creating the header if there is none."""
    if not lines:
        return head
    if not head:
        return "---\n" + "".join(f"{l}\n" for l in lines) + "---\n"
    nl = "\r\n" if head.endswith("\r\n") else "\n"
    body, close = head[:-len(nl)].rsplit(nl, 1)
    return body + nl + "".join(f"{l}{nl}" for l in lines) + close + nl


def save_metadata(src: Path, lines: list[str]) -> None:
    """Write title/author into the manuscript's own YAML header, keeping its
    line endings and byte order mark."""
    raw = src.read_bytes()
    bom = raw.startswith(b"\xef\xbb\xbf")
    text = raw.decode("utf-8-sig")
    head, body = split_yaml(text)
    if not head and "\r\n" in text:
        lines_nl = add_to_yaml("", lines).replace("\n", "\r\n")
    else:
        lines_nl = add_to_yaml(head, lines)
    src.write_bytes((b"\xef\xbb\xbf" if bom else b"") + (lines_nl + body).encode("utf-8"))


def ask_metadata(head: str, body: str, args, src: Path, ask: bool,
                 default_author=None) -> list[str]:
    """Title and author when the YAML header lacks them: from the command
    line, the default author in the settings, or, if `ask` (PDF, Word or HTML
    export), asked for when someone is there to answer. The suggested title is
    an opening H1, otherwise the file name. Returns YAML lines to add to the header."""
    has_title, has_author = has_yaml_key(head, "title"), has_yaml_key(head, "author")
    title = None if has_title else args.title
    authors = [] if has_author else list(args.doc_author or [])
    if not has_author and not authors and default_author:
        authors = [default_author] if isinstance(default_author, str) else list(default_author)
    need_title, need_author = not has_title and not title, not has_author and not authors
    if not ask or not (need_title or need_author):
        lines = yaml_metadata_lines(title, authors)
        if lines and args.save_metadata:
            save_metadata(src, lines)
            print(f"Saved in the YAML header of {src.name}.")
        return lines

    missing = " and ".join(k for k, n in (("title", need_title), ("author", need_author)) if n)
    interactive = not args.no_prompt and sys.stdin is not None and sys.stdin.isatty()
    if not interactive:
        print(f"Note: the export has no {missing}. Add it to the YAML header "
              f"(title: ..., author: ...), use --title / --doc-author, "
              f"or set document_author in settings.toml.", file=sys.stderr)
        return yaml_metadata_lines(title, authors)

    print(f"{src.name} has no {missing}.")
    try:
        if need_title:
            suggestion = leading_h1(body) or src.stem
            answer = input(f"Title [{suggestion}] (Enter = use it, - = no title): ").strip()
            title = None if answer == "-" else answer or suggestion
        if need_author:
            authors = [a.strip() for a in input("Author(s), separated by ';' "
                                                "(Enter = none): ").split(";") if a.strip()]
        lines = yaml_metadata_lines(title, authors)
        if lines and (args.save_metadata or input(
                f"Save in the YAML header of {src.name}? [Y/n] ").strip().lower()
                in ("", "y", "yes", "j", "ja")):
            save_metadata(src, lines)
            print(f"Saved in {src.name}.")
    except EOFError:
        lines = yaml_metadata_lines(title, authors)
    return lines


def convert(body: str, mode: str, fmt: str, author: str, date: str,
            critic: dict | None = None) -> str:
    reviewers = list((critic or {}).get("reviewers", []))
    ids = count()

    def attrs_for(name: str | None, extra: str = "") -> str:
        who = name or author
        a = f'author="{who}" date="{date}"' if who else f'date="{date}"'
        return a + extra

    def accept(g: dict) -> str:
        if g["sub_old"] is not None:
            return g["sub_new"]
        if g["ins"] is not None:
            return g["ins"]
        if g["dele"] is not None:
            return ""
        if g["hl_only"] is not None:
            return g["hl_only"]
        return g["hl"] or ""                    # comments go, highlighted text stays

    def mark(g: dict) -> str:
        if g["sub_old"] is not None:
            a = attrs_for(g["sub_by"])
            return wrap_change(g["sub_old"], "deletion", a) + wrap_change(g["sub_new"], "insertion", a)
        if g["ins"] is not None:
            return wrap_change(g["ins"], "insertion", attrs_for(g["ins_by"]))
        if g["dele"] is not None:
            return wrap_change(g["dele"], "deletion", attrs_for(g["del_by"]))
        if g["hl_only"] is not None:
            return g["hl_only"]

        anchor = g["hl"] or ""
        comment = g["hl_comment"] if g["hl_comment"] is not None else g["comment"]
        name, comment = comment_author(comment, reviewers)
        if not comment.strip():                 # empty {>><<} or a stray {>>@Name<<}
            return anchor
        i = next(ids)
        text = flatten_comment(comment, plain_links=(fmt == "docx"))
        label = f' critic-name="{name}"' if name else ""
        return (f'[{text}]{{.comment-start id="{i}" {attrs_for(name, label)}}}'
                f'{anchor}[]{{.comment-end id="{i}"}}')

    handler = accept if mode == "accept" else mark
    return CRITIC.sub(lambda m: handler(m.groupdict()), body)


# ---------------------------------------------------------------------------
# Code protection
# ---------------------------------------------------------------------------

FENCED_CODE = re.compile(r"^(?P<fence>`{3,}|~{3,}).*?^(?P=fence)[ \t]*$", re.DOTALL | re.MULTILINE)
INLINE_CODE = re.compile(r"(`+)(?!`).+?(?<!`)\1(?!`)", re.DOTALL)


class CodeGuard:
    """Swap code blocks and inline code for placeholders, so CriticMarkup
    and Obsidian syntax inside code is left untouched."""

    def __init__(self) -> None:
        self.saved: list[str] = []

    def _stash(self, m: re.Match) -> str:
        self.saved.append(m.group(0))
        return f"\uE000{len(self.saved) - 1}\uE001"

    def protect(self, text: str) -> str:
        return INLINE_CODE.sub(self._stash, FENCED_CODE.sub(self._stash, text))

    def restore(self, text: str) -> str:
        pattern = re.compile("\uE000(\\d+)\uE001")
        while pattern.search(text):           # code may sit inside restored code
            text = pattern.sub(lambda m: self.saved[int(m.group(1))], text)
        return text


# ---------------------------------------------------------------------------
# Obsidian syntax
# ---------------------------------------------------------------------------

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".bmp", ".tif", ".tiff", ".avif"}
EMBED = re.compile(r"!\[\[([^\]|#^]+?)(#[^\]|]*)?(?:\|([^\]]*))?\]\](\{[^}\n]*\})?")
WIKILINK = re.compile(r"(?<!!)\[\[([^\]|#]*?)(#[^\]|]*)?(?:\|([^\]]*))?\]\]")
OBS_COMMENT = re.compile(r"%%.*?%%", re.DOTALL)
HIGHLIGHT = re.compile(r"(?<![={])==(?=\S)(.+?)(?<=\S)==(?![=}])")
CALLOUT = re.compile(r"^([ \t]*(?:>[ \t]?)+)\[!(\w[\w-]*)\][+-]?[ \t]*(.*)$", re.MULTILINE)
BLOCK_ID = re.compile(r"[ \t]+\^[A-Za-z0-9-]+[ \t]*$", re.MULTILINE)
TAG = re.compile(r"(?<![\w\S])#(?!\d+\b)[\w/-]*[^\W\d_][\w/-]*", re.UNICODE)
SIZE = re.compile(r"^\s*(\d+)(?:\s*x\s*(\d+))?\s*$")
HEADING = re.compile(r"^(#{1,6})[ \t]+(.+?)[ \t]*#*[ \t]*$", re.MULTILINE)


def find_vault(start: Path) -> Path | None:
    for folder in (start, *start.parents):
        if (folder / ".obsidian").is_dir():
            return folder
    return None


class Vault:
    """Resolve Obsidian links the way Obsidian does: by path, relative to the
    note, or by file name anywhere in the vault (shortest path wins)."""

    def __init__(self, root: Path) -> None:
        self.root = root
        self._index: dict[str, list[Path]] | None = None

    def index(self) -> dict[str, list[Path]]:
        if self._index is None:
            self._index = {}
            for folder, dirs, files in os.walk(self.root):
                dirs[:] = [d for d in dirs if not d.startswith(".")
                           and d not in ("node_modules", "critic-exports", "document_exports")]
                for f in files:
                    self._index.setdefault(f.lower(), []).append(Path(folder) / f)
            for paths in self._index.values():
                paths.sort(key=lambda p: (len(p.parts), str(p)))
        return self._index

    def resolve(self, target: str, note_dir: Path) -> Path | None:
        target = target.strip().replace("\\", "/")
        names = [target] if Path(target).suffix else [target + ".md", target]
        for name in names:
            for base in (self.root, note_dir):
                candidate = base / name
                if candidate.is_file():
                    return candidate
            matches = self.index().get(Path(name).name.lower(), [])
            if "/" in name:
                matches = [m for m in matches if m.as_posix().lower().endswith(name.lower())]
            if matches:
                return matches[0]
        return None


def extract_section(text: str, anchor: str) -> str | None:
    """The part of a note a link points to: '#Heading' or '#^block-id'."""
    anchor = anchor.lstrip("#")
    if anchor.startswith("^"):
        block_id = re.escape(anchor[1:])
        for para in PARA_BREAK.split(text):
            if re.search(rf"\^{block_id}[ \t]*$", para, re.MULTILINE):
                return para
        return None
    wanted = anchor.split("#")[-1].strip().lower()
    for m in HEADING.finditer(text):
        if m.group(2).strip().lower() == wanted:
            level = len(m.group(1))
            rest = text[m.start():]
            end = re.compile(rf"^#{{1,{level}}}[ \t]", re.MULTILINE).search(rest, 1)
            return rest[: end.start()] if end else rest
    return None


def markdown_path(path: Path, relative_to: Path) -> str:
    try:
        p = os.path.relpath(path, relative_to)
    except ValueError:                       # another drive on Windows
        p = str(path)
    p = p.replace("\\", "/")
    return f"<{p}>" if re.search(r"[\s()<>]", p) else p


class ObsidianConverter:
    def __init__(self, vault: Vault, out_dir: Path, settings: dict) -> None:
        self.vault = vault
        self.base = out_dir                  # where Pandoc runs (the exported note's folder)
        self.remove_tags = settings["remove_tags"]
        self.max_depth = int(settings["max_embed_depth"])
        self.warnings: list[str] = []

    def warn(self, msg: str) -> None:
        self.warnings.append(msg)
        print(f"Warning: {msg}", file=sys.stderr)

    # Transclusion runs first, on raw text, so embedded notes are processed
    # exactly like the note itself.
    def expand_embeds(self, text: str, note_dir: Path, stack: tuple[Path, ...] = ()) -> str:
        guard = CodeGuard()
        text = guard.protect(text)

        def embed(m: re.Match) -> str:
            target, anchor, label, attrs = m.group(1), m.group(2) or "", m.group(3), m.group(4)
            path = self.vault.resolve(target, note_dir)
            if path is None:
                self.warn(f"embedded file not found: {target}")
                return label or target
            if path.suffix.lower() in IMAGE_EXTS:
                return self.image(path, label, attrs)
            if path.suffix.lower() != ".md":
                return f"[{label or path.name}]({markdown_path(path, self.base)})"
            if path in stack:
                self.warn(f"circular embed skipped: {path.name}")
                return ""
            if len(stack) >= self.max_depth:
                self.warn(f"embed depth limit reached at {path.name}")
                return ""
            _, content = split_yaml(path.read_text(encoding="utf-8-sig"))
            if anchor:
                section = extract_section(content, anchor)
                if section is None:
                    self.warn(f"section {anchor} not found in {path.name}")
                    return ""
                content = section
            content = self.expand_embeds(content, path.parent, stack + (path,))
            content = self.rebase_images(content, path.parent)
            return f"\n\n{content.strip()}\n\n"

        return guard.restore(EMBED.sub(embed, text))

    def image(self, path: Path, label: str | None, attrs: str | None = None) -> str:
        """![[img.png|Caption|400]]{#fig:id} -> ![Caption](img.png){#fig:id width=400px}.
        A caption makes a standalone image a numbered figure (with an id)."""
        parts: list[str] = []
        alt = []
        for piece in (label or "").split("|"):
            size = SIZE.match(piece)
            if size:
                parts.append(f"width={size.group(1)}px")
                if size.group(2):
                    parts.append(f"height={size.group(2)}px")
            elif piece.strip():
                alt.append(piece.strip())
        if attrs:
            parts.insert(0, attrs.strip("{} "))
        attr_text = "{" + " ".join(parts) + "}" if parts else ""
        return f"![{' | '.join(alt)}]({markdown_path(path, self.base)}){attr_text}"

    def rebase_images(self, text: str, note_dir: Path) -> str:
        """Standard images in an embedded note are relative to that note."""
        if note_dir == self.base:
            return text

        def fix(m: re.Match) -> str:
            url = m.group(2).strip("<>")
            if re.match(r"^[a-z]+:|^/|^[A-Za-z]:", url):
                return m.group(0)
            return f"![{m.group(1)}]({markdown_path(note_dir / url, self.base)})"

        return re.sub(r"!\[([^\]]*)\]\(([^)\s]+|<[^>]+>)\)", fix, text)

    @staticmethod
    def strip_comments(text: str) -> str:
        """%% Obsidian comments %% never leave the vault."""
        return OBS_COMMENT.sub("", text)

    def convert(self, text: str) -> str:
        """Inline Obsidian syntax. Runs after CriticMarkup, so {==...==}
        has already been consumed. Code must already be protected."""
        text = WIKILINK.sub(self.wikilink, text)
        text = CALLOUT.sub(self.callout, text)
        text = HIGHLIGHT.sub(r"[\1]{.mark}", text)
        text = BLOCK_ID.sub("", text)
        text = self.separate_lists(text)
        if self.remove_tags:
            lines = []
            for line in text.split("\n"):
                stripped = TAG.sub("", line)
                if line.strip() and not stripped.strip("# \t"):       # tag-only line
                    continue
                if stripped != line:
                    lead = re.match(r"[ \t]*", stripped).group(0)
                    stripped = lead + re.sub(r"[ \t]{2,}", " ", stripped[len(lead):]).rstrip()
                lines.append(stripped)
            text = "\n".join(lines)
        return text

    @staticmethod
    def separate_lists(text: str) -> str:
        """Obsidian starts a list right after a line of text, and ends it at
        the next unindented line; Pandoc needs blank lines around it. As in
        CommonMark, a numbered list can only interrupt text if it starts at 1."""
        out: list[str] = []
        in_list = False
        for line in text.split("\n"):
            prev = out[-1] if out else ""
            if not line.strip():
                pass
            elif LIST_ITEM.match(line) and (in_list or not prev.strip() or LIST_START.match(line)):
                if prev.strip() and not in_list:
                    out.append("")
                in_list = True
            elif line[0] not in " \t":
                if in_list and prev.strip():
                    out.append("")
                in_list = False
            out.append(line)
        return "\n".join(out)

    @staticmethod
    def wikilink(m: re.Match) -> str:
        name, anchor, alias = m.group(1).strip(), (m.group(2) or "").lstrip("#"), m.group(3)
        if alias:
            return alias
        name = Path(name).stem if name.lower().endswith(".md") else name
        anchor = anchor.lstrip("^")
        if name and anchor:
            return f"{name} > {anchor}"
        return name or anchor

    @staticmethod
    def callout(m: re.Match) -> str:
        prefix, kind, title = m.group(1).rstrip(), m.group(2), m.group(3).strip()
        heading = title or kind.replace("-", " ").capitalize()
        return f"{prefix} **{heading}**\n{prefix}"


# Renders changes and comments in HTML and LaTeX/PDF. In Word the spans are
# left alone and become tracked changes and comments.
LUA_FILTER = r"""
local function latex_wrap(content, pre, post)
  local out = pandoc.List({pandoc.RawInline('latex', pre)})
  out:extend(content)
  out:insert(pandoc.RawInline('latex', post))
  return out
end

function Pandoc(doc)
  if FORMAT == 'docx' then return nil end
  local is_latex = FORMAT:match('latex') ~= nil
  local label = 'Comment'
  if doc.meta['critic-comment-label'] then
    label = pandoc.utils.stringify(doc.meta['critic-comment-label'])
  end

  doc = doc:walk({
    Span = function(el)
      if el.classes:includes('comment-start') then
        local who = el.attributes['critic-name'] or label
        local c = pandoc.List({pandoc.Strong({pandoc.Str(who .. ':')}), pandoc.Space()})
        c:extend(el.content)
        return pandoc.Note({pandoc.Para(c)})
      elseif el.classes:includes('comment-end') then
        return {}
      elseif is_latex and el.classes:includes('insertion') then
        return latex_wrap(el.content, '\\textcolor{blue}{', '}')
      elseif is_latex and el.classes:includes('deletion') then
        return latex_wrap(el.content, '\\textcolor{red}{\\sout{', '}}')
      end
    end
  })

  doc.meta['critic-comment-label'] = nil
  if is_latex then
    local pkgs = pandoc.MetaBlocks({
      pandoc.RawBlock('latex', '\\usepackage{xcolor}\n\\usepackage[normalem]{ulem}')
    })
    local hi = doc.meta['header-includes']
    if hi == nil then
      doc.meta['header-includes'] = pkgs
    elseif pandoc.utils.type(hi) == 'List' then
      hi:insert(pkgs)
    else
      doc.meta['header-includes'] = pandoc.List({hi, pkgs})
    end
  end
  return doc
end
"""

CROSSREF_FILTER = r"""
-- Numbered figures, tables, equations and sections with cross-references.
-- Labels:  {#fig:id} {#tbl:id} {#eq:id} {#sec:id}  (pandoc-crossref style)
--          {#fig-id} {#tbl-id} {#eq-id} {#sec-id}  (Quarto style)
-- Refs:    @fig:id  [@fig:a; @fig:b]  [-@fig:id] (number only)  @Fig:id (capitalised)
-- Must run before citeproc, which would otherwise read @fig:id as a citation.

local KINDS = { fig = true, tbl = true, eq = true, sec = true }

local L10N = {
  en = { fig = {"Figure", "figure", "figures"}, tbl = {"Table", "table", "tables"},
         eq = {"Equation", "equation", "equations"}, sec = {"Section", "section", "sections"}, ["and"] = "and" },
  sv = { fig = {"Figur", "figur", "figur"}, tbl = {"Tabell", "tabell", "tabell"},
         eq = {"Ekvation", "ekvation", "ekvation"}, sec = {"Avsnitt", "avsnitt", "avsnitt"}, ["and"] = "och" },
  da = { fig = {"Figur", "figur", "figur"}, tbl = {"Tabel", "tabel", "tabel"},
         eq = {"Ligning", "ligning", "ligning"}, sec = {"Afsnit", "afsnit", "afsnit"}, ["and"] = "og" },
  nb = { fig = {"Figur", "figur", "figur"}, tbl = {"Tabell", "tabell", "tabell"},
         eq = {"Likning", "likning", "likning"}, sec = {"Avsnitt", "avsnitt", "avsnitt"}, ["and"] = "og" },
  de = { fig = {"Abbildung", "Abbildung", "Abbildungen"}, tbl = {"Tabelle", "Tabelle", "Tabellen"},
         eq = {"Gleichung", "Gleichung", "Gleichungen"}, sec = {"Abschnitt", "Abschnitt", "Abschnitte"}, ["and"] = "und" },
  fr = { fig = {"Figure", "figure", "figures"}, tbl = {"Tableau", "tableau", "tableaux"},
         eq = {"Équation", "équation", "équations"}, sec = {"Section", "section", "sections"}, ["and"] = "et" },
  es = { fig = {"Figura", "figura", "figuras"}, tbl = {"Tabla", "tabla", "tablas"},
         eq = {"Ecuación", "ecuación", "ecuaciones"}, sec = {"Sección", "sección", "secciones"}, ["and"] = "y" },
}
L10N.no = L10N.nb
L10N.nn = L10N.nb

local NAMES = { fig = "figure", tbl = "table", eq = "equation", sec = "section" }

local function kind_of(id)
  local k = id:match("^(%a+)[:%-]")
  if k and KINDS[k:lower()] then return k:lower() end
  return nil
end

local function capitalise(s)
  return pandoc.text.upper(pandoc.text.sub(s, 1, 1)) .. pandoc.text.sub(s, 2)
end

local function stringify(v) return v and pandoc.utils.stringify(v) or nil end

function Pandoc(doc)
  -- Labels: built-in table for the document language, overridden by
  -- metadata set from settings.toml (crossref-figure, crossref-figure-ref ...).
  local lang = (stringify(doc.meta.lang) or "en"):lower():match("^(%a+)")
  local base = L10N[lang] or L10N.en
  local labels, conj = {}, stringify(doc.meta["crossref-and"]) or base["and"]
  for k, words in pairs(base) do
    if KINDS[k] then
      local n = NAMES[k]
      labels[k] = {
        stringify(doc.meta["crossref-" .. n]) or words[1],
        stringify(doc.meta["crossref-" .. n .. "-ref"]) or words[2],
        stringify(doc.meta["crossref-" .. n .. "-refs"]) or words[3],
      }
    end
  end
  local sep = stringify(doc.meta["crossref-separator"]) or ": "

  local refs, counters, sections = {}, { fig = 0, tbl = 0, eq = 0 }, {}

  local function register(id, kind, number)
    if refs[id] then io.stderr:write("Warning: duplicate label " .. id .. "\n") end
    refs[id] = { kind = kind, number = number }
  end

  local function prefix_caption(caption, kind, number)
    local lead = pandoc.Inlines({ pandoc.Str(labels[kind][1] .. " " .. number .. sep) })
    if #caption.long == 0 then
      caption.long = pandoc.Blocks({ pandoc.Plain(lead) })
    else
      local first = caption.long[1]
      if first.content then
        lead:extend(first.content)
        first.content = lead
      end
    end
    return caption
  end

  -- Pass 1: number everything in document order.
  doc = doc:walk({
    traverse = "topdown",
    Header = function(h)
      if h.classes:includes("unnumbered") then return nil end
      sections[h.level] = (sections[h.level] or 0) + 1
      for l = h.level + 1, 6 do sections[l] = nil end
      if kind_of(h.identifier) == "sec" then
        local parts, started = {}, false
        for l = 1, h.level do
          local n = sections[l] or 0
          if n > 0 or started then started = true; parts[#parts + 1] = tostring(n) end
        end
        register(h.identifier, "sec", table.concat(parts, "."))
      end
    end,
    Figure = function(f)
      if kind_of(f.identifier) ~= "fig" then return nil end
      counters.fig = counters.fig + 1
      register(f.identifier, "fig", tostring(counters.fig))
      f.caption = prefix_caption(f.caption, "fig", counters.fig)
      return f
    end,
    Table = function(t)
      local id = t.identifier
      local plain = t.caption.long[1]
      if id == "" and plain and plain.content and #plain.content > 0 then
        local last = plain.content[#plain.content]
        local tag = last.t == "Str" and last.text:match("^{#([%w_:%-%.]+)}$")
        if tag and kind_of(tag) == "tbl" then
          id = tag
          plain.content:remove(#plain.content)
          if #plain.content > 0 and plain.content[#plain.content].t == "Space" then
            plain.content:remove(#plain.content)
          end
          t.identifier = id
        end
      end
      if kind_of(id) ~= "tbl" then return t end
      counters.tbl = counters.tbl + 1
      register(id, "tbl", tostring(counters.tbl))
      t.caption = prefix_caption(t.caption, "tbl", counters.tbl)
      return t
    end,
    Inlines = function(inlines)
      -- $$ ... $$ {#eq:id}
      local i, changed = 1, false
      while i <= #inlines do
        local el = inlines[i]
        if el.t == "Math" and el.mathtype == "DisplayMath" then
          local j = i + 1
          if inlines[j] and inlines[j].t == "Space" then j = j + 1 end
          local tag = inlines[j] and inlines[j].t == "Str" and inlines[j].text:match("^{#([%w_:%-%.]+)}$")
          if tag and kind_of(tag) == "eq" then
            counters.eq = counters.eq + 1
            register(tag, "eq", tostring(counters.eq))
            el.text = el.text .. "\\qquad(" .. counters.eq .. ")"
            inlines[i] = pandoc.Span({ el }, pandoc.Attr(tag))
            for _ = i + 1, j do inlines:remove(i + 1) end
            changed = true
          end
        end
        i = i + 1
      end
      if changed then return inlines end
    end,
  })

  -- LaTeX numbers figures and tables itself; switch that off so every
  -- format shows the same numbers (unlabelled figures stay unnumbered).
  if FORMAT:match("latex") and (counters.fig > 0 or counters.tbl > 0) then
    local inc = pandoc.MetaBlocks({ pandoc.RawBlock("latex",
      "\\usepackage{caption}\n\\captionsetup{labelformat=empty}") })
    local hi = doc.meta["header-includes"]
    if hi == nil then
      doc.meta["header-includes"] = inc
    elseif pandoc.utils.type(hi) == "List" then
      hi:insert(inc)
    else
      doc.meta["header-includes"] = pandoc.List({ hi, inc })
    end
  end

  -- Pass 2: replace references.
  local function number_list(items)
    local out = pandoc.Inlines({})
    for n, item in ipairs(items) do
      if n > 1 then
        if n == #items then out:extend({ pandoc.Space(), pandoc.Str(conj), pandoc.Space() })
        else out:extend({ pandoc.Str(","), pandoc.Space() }) end
      end
      out:insert(pandoc.Link(item.number, "#" .. item.id))
    end
    return out
  end

  doc = doc:walk({
    Cite = function(cite)
      for _, c in ipairs(cite.citations) do
        if not kind_of(c.id) then return nil end   -- an ordinary citation
      end
      local out, group = pandoc.Inlines({}), {}
      local suppress = cite.citations[1].mode == "SuppressAuthor"
      local capital = cite.citations[1].id:match("^%u") ~= nil

      local function flush()
        if #group == 0 then return end
        if #out > 0 then out:extend({ pandoc.Str(","), pandoc.Space() }) end
        if not suppress then
          local word = labels[group[1].kind][#group > 1 and 3 or 2]
          if capital and #out == 0 then word = capitalise(word) end
          out:extend({ pandoc.Str(word), pandoc.Space() })
        end
        out:extend(number_list(group))
        group = {}
      end

      for _, c in ipairs(cite.citations) do
        local key = kind_of(c.id) .. c.id:sub(#c.id:match("^%a+") + 1)
        local ref = refs[key]
        local item = { id = key, kind = kind_of(c.id), number = ref and ref.number or "??" }
        if not ref then io.stderr:write("Warning: reference to unknown label " .. c.id .. "\n") end
        if #group > 0 and group[1].kind ~= item.kind then flush() end
        group[#group + 1] = item
        if c.prefix and #c.prefix > 0 then out:extend(c.prefix); out:insert(pandoc.Space()) end
      end
      flush()
      return out
    end,
  })
  for key in pairs(doc.meta) do        -- settings passed as metadata are not content
    if key:match("^crossref%-") then doc.meta[key] = nil end
  end
  return doc
end
"""

# Renders ::: pagebreak divs as a real page break in each format.
PAGEBREAK_FILTER = r"""
local BREAKS = {
  docx = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
  latex = '\\newpage{}',
  beamer = '\\newpage{}',
  context = '\\page',
  typst = '#pagebreak()',
  ms = '.bp',
  odt = '<text:p text:style-name="Pagebreak"/>',
  opendocument = '<text:p text:style-name="Pagebreak"/>',
}
local RAW = { docx = 'openxml', odt = 'opendocument', opendocument = 'opendocument',
              latex = 'latex', beamer = 'latex', context = 'context', typst = 'typst', ms = 'ms' }

function Div(el)
  if not el.classes:includes('pagebreak') then return nil end
  local fmt = FORMAT
  if BREAKS[fmt] then
    return pandoc.RawBlock(RAW[fmt], BREAKS[fmt])
  elseif fmt:match('html') or fmt:match('epub') then
    return pandoc.RawBlock('html', '<div style="break-after: page; page-break-after: always;"></div>')
  end
  return {}
end
"""

ZOTERO_FILTER = r"""
-- Turns citations into Zotero fields in Word, after citeproc has rendered
-- them, so the Zotero Word plugin can refresh and edit them. The item data
-- comes from Better BibTeX (a JSON file named in CRITIC_EXPORT_ZOTERO_ITEMS,
-- not in metadata, which Word would keep as a document property).
-- The rendered text is shown until the document is refreshed in Word.
local items = {}

local LABELS = {
  ["p."] = "page", ["pp."] = "page", ["s."] = "page", ["S."] = "page",
  ["chap."] = "chapter", ["ch."] = "chapter", ["kap."] = "chapter",
  ["sec."] = "section", ["para."] = "paragraph", ["fig."] = "figure",
  ["vol."] = "volume", ["n."] = "note", ["l."] = "line", ["ll."] = "line",
  ["no."] = "issue", ["col."] = "column", ["bk."] = "book", ["v."] = "verse",
}

local function esc(s)
  s = s:gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;")
  return s
end

local function run(text)
  return '<w:r><w:t xml:space="preserve">' .. esc(text) .. '</w:t></w:r>'
end

local BEGIN = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
local SEPARATE = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
local END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

local function instr(text)
  return '<w:r><w:instrText xml:space="preserve">' .. esc(text) .. '</w:instrText></w:r>'
end

local function random_id()
  local chars, out = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", {}
  for _ = 1, 8 do
    local i = math.random(#chars)
    out[#out + 1] = chars:sub(i, i)
  end
  return table.concat(out)
end

-- ", chap. 3, emphasis added" -> "3", "chapter", ", emphasis added"
local function locator(suffix)
  local s = suffix:gsub("^%s*,?%s*", "")
  local label = "page"
  local word, rest = s:match("^(%S+%.)%s*(.*)$")
  if word and LABELS[word] then
    label, s = LABELS[word], rest
  end
  local loc, tail = s:match("^(%w*%d[%w%-–,]*)(.*)$")
  if loc then
    loc = loc:gsub(",$", "")
    return loc, label, (tail:match("^%s*$") and "" or tail)
  end
  return nil, nil, suffix
end

local function item_entry(c)
  local data = items[c.id]
  if not data then return nil end
  local custom = data.custom or {}
  local item = {}
  for k, v in pairs(data) do
    if k ~= "custom" then item[k] = v end
  end
  item.id = custom.itemID or c.id
  local entry = {id = custom.itemID or c.id, uris = {custom.uri}, itemData = item}
  local loc, label, suffix = locator(pandoc.utils.stringify(c.suffix))
  if loc then entry.locator, entry.label = loc, label end
  if suffix ~= "" then entry.suffix = suffix end
  local prefix = pandoc.utils.stringify(c.prefix)
  if prefix ~= "" then entry.prefix = prefix end
  if c.mode == "SuppressAuthor" then entry["suppress-author"] = true end
  return entry
end

local function read_items()
  local path = os.getenv("CRITIC_EXPORT_ZOTERO_ITEMS")
  if not path then return end
  local f = io.open(path, "r")
  if not f then return end
  items = pandoc.json.decode(f:read("a"), false)
  f:close()
  math.randomseed(os.time())
end

local function cite(c)
  local entries = {}
  for _, citation in ipairs(c.citations) do
    local e = item_entry(citation)
    if not e then return nil end                 -- not in Zotero: leave as text
    entries[#entries + 1] = e
  end
  local text = pandoc.utils.stringify(c.content)
  local lead = ""
  if c.citations[1].mode == "AuthorInText" then
    -- "Dawkins (1986, 13)": the name stays text, Zotero gets "(1986, 13)"
    local author, rest = text:match("^(.-)%s*(%b())$")
    if author then
      lead, text = author .. " ", rest
      entries[1]["suppress-author"] = true
    end
  end
  local json = pandoc.json.encode({
    citationID = random_id(),
    properties = {formattedCitation = text, plainCitation = text, noteIndex = 0},
    citationItems = entries,
    schema = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",
  })
  return pandoc.RawInline("openxml", (lead ~= "" and run(lead) or "") .. BEGIN
    .. instr(" ADDIN ZOTERO_ITEM CSL_CITATION " .. json .. " ") .. SEPARATE .. run(text) .. END)
end

local function bibliography(div)
  if div.identifier ~= "refs" then return nil end
  local entries = {}
  for _, block in ipairs(div.content) do
    entries[#entries + 1] = pandoc.utils.stringify(block)
  end
  if #entries == 0 then return nil end
  local ppr = '<w:pPr><w:pStyle w:val="Bibliography"/></w:pPr>'
  local out = {}
  for i, entry in ipairs(entries) do
    local runs = run(entry)
    if i == 1 then
      runs = BEGIN .. instr(' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} '
        .. 'CSL_BIBLIOGRAPHY ') .. SEPARATE .. runs
    end
    if i == #entries then runs = runs .. END end
    out[#out + 1] = '<w:p>' .. ppr .. runs .. '</w:p>'
  end
  return pandoc.RawBlock("openxml", table.concat(out))
end

return {{Meta = read_items}, {Cite = cite, Div = bibliography}}
"""

CRITIC_CSS = """.insertion { color: #1a6fb5; text-decoration: underline; }
.deletion  { color: #b3261e; text-decoration: line-through; }
"""

# ---------------------------------------------------------------------------
# Tools and citation styles
# ---------------------------------------------------------------------------


def find_tool(configured: str, default: str, hint: str) -> str:
    name = configured or default
    path = expand(name)
    if path.exists():
        return str(path)
    found = shutil.which(name)
    if found:
        return found
    sys.exit(f"Cannot find {name}. {hint}")


BIB_KEY = re.compile(r"^\s*@\w+\s*\{\s*([^,\s]+)\s*,", re.MULTILINE)


def resolve_bibliographies(items, base: Path) -> list[Path]:
    """A bibliography setting is one path or a list of paths; each may be a
    glob pattern such as "~/Zotero/exports/*.json"."""
    if isinstance(items, str):
        items = [items] if items.strip() else []
    paths: list[Path] = []
    for item in items:
        pattern = expand(item, base)
        if glob.has_magic(str(pattern)):
            matches = sorted(Path(p) for p in glob.glob(str(pattern)))
            if not matches:
                print(f"Warning: no bibliography files match {pattern}", file=sys.stderr)
            paths += matches
        elif pattern.exists():
            paths.append(pattern)
        else:
            sys.exit(f"Bibliography not found: {pattern}\n"
                     "Set up a Better BibTeX auto-export (see README).")
    return list(dict.fromkeys(paths))


def bibliography_keys(path: Path) -> list[str]:
    """Citation keys in a CSL JSON or BibTeX/BibLaTeX file (others: none)."""
    try:
        text = path.read_text(encoding="utf-8-sig", errors="replace")
        if path.suffix.lower() == ".json":
            data = json.loads(text)
            items = data.get("items", []) if isinstance(data, dict) else data
            return [str(i["id"]) for i in items if isinstance(i, dict) and "id" in i]
        if path.suffix.lower() in (".bib", ".bibtex"):
            return BIB_KEY.findall(text)
    except (OSError, ValueError):
        pass
    return []


def warn_duplicate_keys(paths: list[Path]) -> None:
    """Pandoc silently uses the last file when a key occurs in several."""
    if len(paths) < 2:
        return
    seen: dict[str, Path] = {}
    dupes: dict[str, list[Path]] = {}
    for p in paths:
        for key in bibliography_keys(p):
            if key in seen and seen[key] != p:
                dupes.setdefault(key, [seen[key]]).append(p)
            seen[key] = p
    if dupes:
        print(f"Warning: {len(dupes)} citation key(s) occur in more than one bibliography; "
              "the last file listed wins:", file=sys.stderr)
        for key, files in list(dupes.items())[:10]:
            print(f"  {key}: {', '.join(f.name for f in files)}", file=sys.stderr)
        if len(dupes) > 10:
            print(f"  ... and {len(dupes) - 10} more", file=sys.stderr)


def zotero_styles_dir(settings: dict) -> Path:
    configured = settings["citations"]["zotero_data_dir"]
    base = settings.get("_base")
    return (expand(configured, base) if configured else Path.home() / "Zotero") / "styles"


def style_title(csl: Path) -> str:
    try:
        head = csl.read_text(encoding="utf-8", errors="ignore")[:4000]
    except OSError:
        return ""
    m = re.search(r"<title>(.*?)</title>", head, re.DOTALL)
    return m.group(1).strip() if m else ""


def resolve_style(style: str, settings: dict, base: Path) -> Path:
    """A style is the name of a style installed in Zotero (the file name
    without .csl, e.g. "apa", "vancouver") or a path to a .csl file.
    A path that does not exist falls back to a Zotero style with the same
    file name, so styles can be moved into Zotero without editing settings."""
    styles = zotero_styles_dir(settings)
    candidate = expand(style, base)
    is_path = candidate.suffix.lower() == ".csl" or any(sep in style for sep in "/\\")
    if is_path:
        if candidate.exists():
            return candidate
        fallback = styles / f"{candidate.stem}.csl"
        if fallback.exists():
            print(f"Note: {candidate} not found; using {candidate.stem} from Zotero.", file=sys.stderr)
            return fallback
        sys.exit(f"Citation style not found: {candidate}\n"
                 f"(also looked for {candidate.stem}.csl in {styles})")
    match = styles / f"{style}.csl"
    if match.exists():
        return match
    installed = [f.stem for f in styles.glob("*.csl")] if styles.is_dir() else []
    close = difflib.get_close_matches(style, installed, n=3, cutoff=0.5)
    hint = f" Did you mean: {', '.join(close)}?" if close else ""
    sys.exit(f'Citation style "{style}" is not installed in {styles}.{hint}\n'
             "Run with --list-styles to see installed styles, or give a path to a .csl file.")


def list_styles(settings: dict) -> None:
    styles = zotero_styles_dir(settings)
    files = sorted(styles.glob("*.csl")) if styles.is_dir() else []
    if not files:
        print(f"No styles found in {styles}")
        return
    width = max(len(f.stem) for f in files)
    for f in files:
        print(f"{f.stem:<{width}}  {style_title(f)}")


ODT_PAGEBREAK_STYLE = ('<style:style style:name="Pagebreak" style:family="paragraph">'
                       '<style:paragraph-properties fo:break-after="page"/></style:style>')


def add_odt_pagebreak_style(path: Path) -> None:
    """Pandoc's default ODT template has no "Pagebreak" paragraph style, so
    page breaks would be ignored. Add it to the document's automatic styles
    unless the template (--reference-doc) already defines it."""
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info.filename)) for info in z.infolist()]
    content = next(data for info, data in items if info.filename == "content.xml").decode("utf-8")
    styles = next((data for info, data in items if info.filename == "styles.xml"), b"").decode("utf-8")
    if "Pagebreak" not in content or 'style:name="Pagebreak"' in content + styles:
        return
    if "<office:automatic-styles/>" in content:
        content = content.replace("<office:automatic-styles/>",
                                  f"<office:automatic-styles>{ODT_PAGEBREAK_STYLE}</office:automatic-styles>", 1)
    else:
        content = content.replace("</office:automatic-styles>",
                                  f"{ODT_PAGEBREAK_STYLE}</office:automatic-styles>", 1)
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w") as z:
        for info, data in items:                 # mimetype stays first and uncompressed
            if info.filename == "content.xml":
                data = content.encode("utf-8")
            z.writestr(info, data, compress_type=info.compress_type)
    tmp.replace(path)


def open_file(path: Path) -> None:
    try:
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except OSError as e:
        print(f"Could not open {path.name}: {e}", file=sys.stderr)


def ensure_writable(path: Path) -> None:
    """A file open in Word or a PDF viewer is locked on Windows."""
    if path.exists():
        try:
            with open(path, "a"):
                pass
        except PermissionError:
            sys.exit(f"{path.name} is open in another program. Close it and try again.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        prog="critic-export",
        description="Export Markdown with CriticMarkup and Zotero citations to Word, "
                    "PDF, HTML and other formats via Pandoc. Understands Obsidian syntax.")
    ap.add_argument("input", nargs="?", type=Path, help="Markdown file (.md or .qmd)")
    ap.add_argument("--to", nargs="+", metavar="FORMAT",
                    help="one or more output formats: docx, pdf, html, odt, epub, latex, "
                         "typst, rtf, md, ... (default from settings: docx)")
    ap.add_argument("--critic", choices=CRITIC_MODES,
                    help="CriticMarkup handling for all formats in this run "
                         "(default from settings: track for docx, markup for html, accept otherwise)")
    ap.add_argument("--style", help="citation style: Zotero style name or path to a .csl file")
    ap.add_argument("--bibliography", nargs="+", metavar="FILE",
                    help="one or more bibliography files (.json, .bib, .yaml); "
                         "glob patterns allowed")
    ap.add_argument("-o", "--output", type=Path,
                    help="output file (only with a single format, or with --import)")
    ap.add_argument("--author", help="author name for tracked changes and comments")
    ap.add_argument("--title", help="document title, when the YAML header has none")
    ap.add_argument("--doc-author", action="append", metavar="NAME",
                    help="document author, when the YAML header has none (repeat for several)")
    ap.add_argument("--save-metadata", action="store_true",
                    help="write --title / --doc-author, or the answers to the prompt, "
                         "into the manuscript's YAML header without asking")
    ap.add_argument("--no-prompt", action="store_true",
                    help="never ask for a missing title or author")
    ap.add_argument("--reference-doc", help="Word or ODT template for styles")
    ap.add_argument("--zotero-fields", action=argparse.BooleanOptionalAction,
                    help="make citations in Word Zotero fields that the Zotero Word plugin can "
                         "refresh and edit (needs Zotero running; default from settings: off)")
    ap.add_argument("--obsidian", choices=("auto", "on", "off"),
                    help="convert Obsidian syntax (default from settings: auto = when the "
                         "file is inside an Obsidian vault)")
    ap.add_argument("--import", dest="import_docx", metavar="DOCX",
                    help="bring the tracked changes and comments in a reviewed Word file "
                         "back into the manuscript as CriticMarkup")
    ap.add_argument("--in-place", action="store_true",
                    help="with --import: update the manuscript itself (a copy is kept)")
    ap.add_argument("--check", action="store_true",
                    help="report problems (unknown citations and labels, missing images, "
                         "unbalanced CriticMarkup) without exporting")
    ap.add_argument("--count", action="store_true",
                    help="count words and characters in the clean text")
    ap.add_argument("--settings", type=Path, help="settings file to use")
    ap.add_argument("--open", action="store_true", help="open the exported file(s) afterwards")
    ap.add_argument("--keep", action="store_true", help="keep intermediate files for debugging")
    ap.add_argument("--list-styles", action="store_true", help="list citation styles installed in Zotero")
    ap.add_argument("--show-settings", action="store_true",
                    help="show which settings file is used and the resulting settings")
    ap.add_argument("--pause-on-issues", action="store_true",
                    help="wait for Enter before closing if there were errors or warnings "
                         "(used by the Explorer context menu)")
    ap.add_argument("--update-settings", action="store_true",
                    help="add settings from settings.example.toml that settings.toml lacks")
    ap.add_argument("--install-context-menu", action="store_true",
                    help="add a critic-export submenu to the Windows Explorer context menu "
                         "for the file types in [context_menu] (current user, no admin rights)")
    ap.add_argument("--uninstall-context-menu", action="store_true",
                    help="remove the Explorer context menu")
    ap.add_argument("--start", action="store_true",
                    help="start the Obsidian hotkeys and offer to set up autostart and the "
                         "Explorer context menu (what double-clicking critic-export.cmd does)")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return ap.parse_args()


def obsidian_enabled(choice, vault: Path | None) -> bool:
    if isinstance(choice, bool):
        return choice
    choice = str(choice).lower()
    if choice in ("on", "true", "yes"):
        return True
    if choice in ("off", "false", "no"):
        return False
    return vault is not None                               # auto


def critic_mode_for(fmt: str, requested: str, settings: dict) -> tuple[str, str | None]:
    """Return the mode to use and an optional note if it had to change."""
    crit = settings["critic"]
    mode = requested or crit.get(fmt, crit["other"])
    if mode not in CRITIC_MODES:
        sys.exit(f'Invalid critic mode "{mode}" for {fmt} in settings.')
    if fmt in TRACK_FORMATS:
        return ("track" if mode == "markup" else mode), None
    if mode == "track":
        mode = "markup"                                    # tracked changes exist only in Word
    if mode == "markup" and fmt not in MARKUP_FORMATS:
        return "accept", f"{fmt} cannot show CriticMarkup; exporting with changes accepted."
    return mode, None


# ---------------------------------------------------------------------------
# Windows Explorer context menu
# ---------------------------------------------------------------------------

MENU_KEY = "critic-export"


def _menu_root(ext: str) -> str:
    # SystemFileAssociations applies per extension, whatever program opens the file.
    return rf"Software\Classes\SystemFileAssociations\{ext}\shell\{MENU_KEY}"


def _delete_tree(winreg, root, path: str) -> None:
    try:
        with winreg.OpenKey(root, path, 0, winreg.KEY_ALL_ACCESS) as key:
            previous = None
            while True:
                try:
                    child = winreg.EnumKey(key, 0)
                except OSError:
                    break
                if child == previous:            # could not be deleted; avoid looping
                    break
                previous = child
                _delete_tree(winreg, root, f"{path}\\{child}")
        winreg.DeleteKey(root, path)
    except FileNotFoundError:
        pass


def uninstall_context_menu(settings: dict, quiet: bool = False) -> None:
    if sys.platform != "win32":
        sys.exit("The context menu is only available on Windows.")
    import winreg
    exts = set(settings["context_menu"]["extensions"]) | set(DEFAULTS["context_menu"]["extensions"])
    for ext in exts:
        _delete_tree(winreg, winreg.HKEY_CURRENT_USER, _menu_root(ext))
    if not quiet:
        print("Removed the critic-export context menu.")


def install_context_menu(settings: dict) -> None:
    """A cascading "critic-export" submenu on the chosen file types. Written to
    HKEY_CURRENT_USER, so no administrator rights are needed. Each command runs
    critic-export.cmd, so the menu keeps working when Python is updated."""
    if sys.platform != "win32":
        sys.exit("The context menu is only available on Windows.")
    import winreg
    menu = settings["context_menu"]
    items = menu["items"]
    if not items:
        sys.exit("No [context_menu] items in settings.toml.")
    python = Path(sys.executable)
    uninstall_context_menu(settings, quiet=True)          # start from a clean slate

    for ext in menu["extensions"]:
        ext = ext if ext.startswith(".") else "." + ext
        root = _menu_root(ext.lower())
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, root) as key:
            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, menu["title"])
            winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ, "")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, f"{python},0")
        for n, item in enumerate(items, 1):
            sub = f"{root}\\shell\\{n:02d}"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, sub) as key:
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, item["label"])
            command = f'"{START_CMD}" "%1" {item["args"]} --pause-on-issues'
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, sub + "\\command") as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)

    print(f"Added the \"{menu['title']}\" context menu for {', '.join(menu['extensions'])}.")
    print("On Windows 11 it is under \"Show more options\" (or Shift + right-click).")
    print("Run --install-context-menu again after moving the critic-export folder or changing [context_menu].")


STARTUP_SHORTCUT = "critic-export hotkeys.lnk"


def _yes(question: str) -> bool:
    try:
        return input(f"{question} [Y/n] ").strip().lower() in ("", "y", "yes", "j", "ja")
    except EOFError:
        return False


def context_menu_installed(settings: dict) -> bool:
    import winreg
    ext = (settings["context_menu"]["extensions"] or [".md"])[0]
    try:
        winreg.OpenKey(winreg.HKEY_CURRENT_USER, _menu_root(ext if ext.startswith(".") else "." + ext)).Close()
        return True
    except OSError:
        return False


def run_start(settings: dict) -> None:
    """Double-clicking critic-export.cmd: start the Obsidian hotkeys, offer
    the one-time setup that is still missing, and show how to export."""
    print(f"critic-export {__version__}\n")
    if sys.platform == "win32":
        try:
            os.startfile(HOTKEY_SCRIPT)
            print("Obsidian hotkeys are running:\n"
                  "  Ctrl+Alt+W        Word\n"
                  "  Ctrl+Alt+P        PDF with changes and comments shown\n"
                  "  Ctrl+Alt+Shift+P  PDF with changes accepted\n"
                  "  Ctrl+Alt+H        HTML\n")
        except OSError:
            print("The Obsidian hotkeys need AutoHotkey v2: https://www.autohotkey.com\n")
        startup = Path(os.environ.get("APPDATA", "")) / "Microsoft/Windows/Start Menu/Programs/Startup"
        link = startup / STARTUP_SHORTCUT
        if startup.is_dir() and not link.exists() \
                and _yes("Start the hotkeys automatically when you log in?"):
            q = lambda p: str(p).replace("'", "''")
            subprocess.run(["powershell", "-NoProfile", "-Command",
                            f"$s = (New-Object -ComObject WScript.Shell).CreateShortcut('{q(link)}'); "
                            f"$s.TargetPath = '{q(HOTKEY_SCRIPT)}'; $s.Save()"], check=False)
            print(f"Added {link.name} to the Startup folder.\n" if link.exists()
                  else "Could not add the shortcut to the Startup folder.\n")
        if not context_menu_installed(settings) \
                and _yes("Add critic-export to the Explorer right-click menu for Markdown files?"):
            install_context_menu(settings)
            print()
    print("From a terminal:  critic-export note.md --to pdf   (critic-export --help for all options)")
    print(f"Settings:         {LOCAL_SETTINGS}")
    try:
        input("\nPress Enter to close...")
    except EOFError:
        pass


def show_settings(path: Path | None, settings: dict) -> None:
    print(f"Settings file in use: {path or 'none (built-in defaults)'}")
    if settings.get("_project"):
        print(f"Project settings:     {settings['_project']}")
    if path != LOCAL_SETTINGS:
        print(f"  (default location: {LOCAL_SETTINGS})")
    print()
    for section, values in settings.items():
        if section.startswith("_") or not values:
            continue
        print(f"[{section}]")
        for key, value in values.items():
            print(f"  {key} = {value!r}")
    print(f"\nZotero styles folder: {zotero_styles_dir(settings)}")
    if path and EXAMPLE_SETTINGS.exists():
        try:
            missing = missing_settings(path.read_text(encoding="utf-8"),
                                       EXAMPLE_SETTINGS.read_text(encoding="utf-8"))
        except (OSError, tomllib.TOMLDecodeError):
            missing = []
        if missing:
            print(f"\n{len(missing)} setting(s) documented in {EXAMPLE_SETTINGS.name} are not in "
                  f"{path.name} (built-in defaults are used). Run --update-settings to add them.")


# ---------------------------------------------------------------------------
# Per-project settings
# ---------------------------------------------------------------------------

PROJECT_SETTINGS = "critic-export.toml"
PROJECT_PATH_KEYS = {("citations", "bibliography"), ("citations", "zotero_data_dir"),
                     ("docx", "reference_doc"), ("html", "css")}


def find_project_settings(start: Path) -> Path | None:
    """The nearest critic-export.toml in the document's folder or above."""
    for folder in (start, *start.parents):
        candidate = folder / PROJECT_SETTINGS
        if candidate.is_file():
            return candidate
    return None


def _project_path(value, base: Path):
    if isinstance(value, list):
        return [_project_path(v, base) for v in value]
    if isinstance(value, str) and value.strip():
        return str(expand(value, base))
    return value


def merge_project_settings(settings: dict, path: Path) -> None:
    """Values in critic-export.toml override settings.toml for documents in
    that folder and below. Relative paths are relative to the project file."""
    try:
        data = tomllib.loads(path.read_text(encoding="utf-8"))
    except tomllib.TOMLDecodeError as e:
        sys.exit(f"Error in {path}: {e}")
    for section, values in data.items():
        if section not in DEFAULTS or not isinstance(values, dict):
            print(f"Warning: unknown section [{section}] in {path}", file=sys.stderr)
            continue
        for key, value in values.items():
            if (section, key) in PROJECT_PATH_KEYS:
                value = _project_path(value, path.parent)
            elif section == "citations" and key == "style" and isinstance(value, str) \
                    and (value.lower().endswith(".csl") or "/" in value or "\\" in value):
                value = str(expand(value, path.parent))
            elif section == "tools" and isinstance(value, str) and ("/" in value or "\\" in value):
                value = str(expand(value, path.parent))
            settings[section][key] = value
    settings["_project"] = path


# ---------------------------------------------------------------------------
# --check
# ---------------------------------------------------------------------------

CITEKEY = re.compile(r"(?<![\w.@])@([A-Za-z0-9_](?:\w|[:.#$%&+?<>~/-](?=\w))*)")
CROSSREF_KEY = re.compile(r"^(?i:fig|tbl|eq|sec)[:-]")
LABEL = re.compile(r"\{#((?i:fig|tbl|eq|sec)[:-][\w:.-]*)")
IMAGE = re.compile(r"!\[[^\]]*\]\((<[^>]+>|[^)\s]+)")
DELIMITER = re.compile(r"\{\+\+|\+\+\}|\{--|--\}|\{~~|~~\}|\{>>|<<\}|\{==|==\}")


def yaml_bibliographies(head: str, base: Path) -> list[Path]:
    """bibliography: entries from a YAML header (single value or list)."""
    m = re.search(r"^bibliography\s*:\s*(.*?)\s*$((?:\n\s*-\s.*)*)", head, re.MULTILINE)
    if not m:
        return []
    items = [x.strip(" '\"") for x in m.group(1).strip("[]").split(",") if x.strip(" '\"")]
    items += [x.strip()[1:].strip(" '\"") for x in m.group(2).splitlines() if x.strip()]
    return [expand(i, base) for i in items]


def run_check(args, settings: dict, src: Path) -> int:
    """Report problems without exporting. Exit code 1 if there are any."""
    raw = src.read_text(encoding="utf-8-sig")
    doc = Document(src, settings)
    problems: list[str] = []
    notes: list[str] = []

    def where(snippet: str) -> str:
        i = raw.find(snippet)
        return f"line {raw.count(chr(10), 0, i) + 1}" if i >= 0 else "in an embedded note"

    if doc.obsidian:
        problems += doc.obsidian.warnings
    guard = CodeGuard()
    text = guard.protect(doc.body)
    if text.count("%%") % 2:
        problems.append("an Obsidian %%comment%% is not closed (odd number of %%)")
    if doc.obsidian:
        text = doc.obsidian.strip_comments(text)
    text = re.sub(r"<!--.*?-->", "", text, flags=re.DOTALL)

    # CriticMarkup: anything left after removing complete markup is unbalanced.
    kinds = {"ins": "insertion", "dele": "deletion", "sub_old": "replacement",
             "comment": "comment", "hl_comment": "comment", "hl_only": "highlight"}
    found: dict[str, int] = {}
    for m in CRITIC.finditer(text):
        g = m.groupdict()
        kind = next(kinds[k] for k in kinds if g[k] is not None)
        if kind == "comment" and not comment_author(g["comment"] or g["hl_comment"] or "",
                                                    settings["critic"]["reviewers"])[1].strip():
            continue
        found[kind] = found.get(kind, 0) + 1
    for m in DELIMITER.finditer(CRITIC.sub("", text)):
        problems.append(f"{where(m.group(0))}: unmatched CriticMarkup {m.group(0)}")
    if found:
        listed = ", ".join(f"{n} {k}{'s' if n > 1 else ''}" for k, n in sorted(found.items()))
        notes.append(f"review markup still in the text: {listed} "
                     "(export with --critic accept for a clean copy)")

    # Everything below looks at the text as exported: changes accepted,
    # comments (and their @Name tags) gone.
    text = convert(text, "accept", "check", "", "", settings["critic"])

    # Citation keys against the bibliographies.
    try:
        _, bibs = citation_setup(args, settings, doc.head)
    except SystemExit as e:
        problems.append(str(e.code))
        bibs = []
    if not bibs:
        bibs = yaml_bibliographies(doc.head, src.parent)
    known: set[str] = set()
    checkable = True
    for b in bibs:
        if not b.exists():
            problems.append(f"bibliography not found: {b}")
            checkable = False
        elif b.suffix.lower() not in (".json", ".bib", ".bibtex"):
            checkable = False
        known.update(bibliography_keys(b))
    refs = [k for k in CITEKEY.findall(text) if not CROSSREF_KEY.match(k)]
    if refs and not bibs:
        notes.append(f"{len(set(refs))} citation key(s) found, but no bibliography is set")
    elif refs and not checkable:
        notes.append("citation keys not checked: only CSL JSON and BibTeX files can be read")
    elif refs:
        for key in dict.fromkeys(refs):
            if key not in known:
                problems.append(f"{where('@' + key)}: unknown citation key @{key}")
    warn_duplicate_keys([b for b in bibs if b.exists()])

    # Cross-reference labels and references.
    labels = [l[0].lower() + l[1:] for l in LABEL.findall(text)]
    for label in {l for l in labels if labels.count(l) > 1}:
        problems.append(f"{where('{#' + label)}: label {label} is used more than once")
    if settings["crossref"]["enabled"]:
        for key in dict.fromkeys(k for k in CITEKEY.findall(text) if CROSSREF_KEY.match(k)):
            norm = key[0].lower() + key[1:]
            if norm not in labels:
                problems.append(f"{where('@' + key)}: reference to unknown label @{key}")

    # Images.
    for m in IMAGE.finditer(text):
        url = m.group(1).strip("<>")
        if re.match(r"^[a-z][a-z0-9+.-]*:", url, re.IGNORECASE) and not re.match(r"^[A-Za-z]:", url):
            continue                                    # http:, data: ...
        if not (src.parent / unquote(url)).exists():
            problems.append(f"{where(m.group(1))}: image not found: {url}")

    print(f"Checked {src.name}: {len(problems)} problem(s)")
    for p in problems:
        print(f"  - {p}")
    for n in notes:
        print(f"  Note: {n}")
    return 1 if problems else 0


# ---------------------------------------------------------------------------
# --count
# ---------------------------------------------------------------------------

COUNT_FILTER = r"""
local cats = { "body", "headings", "captions", "tables", "footnotes" }
local counts = {}
for _, c in ipairs(cats) do counts[c] = { 0, 0, 0 } end
local cites = 0

local function add(cat, text)
  local c = counts[cat]
  for w in text:gmatch("%S+") do
    if w:match("[%w\128-\255]") then c[1] = c[1] + 1 end
  end
  c[2] = c[2] + (utf8.len(text) or #text)
  local compact = text:gsub("%s", "")
  c[3] = c[3] + (utf8.len(compact) or #compact)
end

local count_blocks

local function clean(inlines)
  return pandoc.Inlines(inlines):walk({
    Cite = function(_) cites = cites + 1; return {} end,
    Note = function(n) count_blocks(n.content, "footnotes"); return {} end,
  })
end

local function caption_text(caption)
  local parts = {}
  for _, b in ipairs(caption.long or {}) do
    if b.content then parts[#parts + 1] = pandoc.utils.stringify(clean(b.content)) end
  end
  return table.concat(parts, " ")
end

local excluded = {}
local skip_level = nil

count_blocks = function(blocks, cat)
  for _, b in ipairs(blocks) do
    local t = b.t
    if t == "Header" and skip_level and b.level <= skip_level then skip_level = nil end
    if t == "Header" and cat == "body" and excluded[pandoc.text.lower(pandoc.utils.stringify(b.content))] then
      skip_level = b.level
    end
    if skip_level then
      -- inside an excluded section (reference list written by hand)
    elseif t == "Header" then
      add(cat == "body" and "headings" or cat, pandoc.utils.stringify(clean(b.content)))
    elseif t == "Para" or t == "Plain" then
      add(cat, pandoc.utils.stringify(clean(b.content)))
    elseif t == "LineBlock" then
      for _, line in ipairs(b.content) do add(cat, pandoc.utils.stringify(clean(line))) end
    elseif t == "BulletList" or t == "OrderedList" then
      for _, item in ipairs(b.content) do count_blocks(item, cat) end
    elseif t == "DefinitionList" then
      for _, item in ipairs(b.content) do
        add(cat, pandoc.utils.stringify(clean(item[1])))
        for _, def in ipairs(item[2]) do count_blocks(def, cat) end
      end
    elseif t == "BlockQuote" or t == "Div" then
      if b.identifier ~= "refs" then count_blocks(b.content, cat) end
    elseif t == "Figure" then
      add("captions", caption_text(b.caption))
    elseif t == "Table" then
      local cap = caption_text(b.caption)
      add("captions", cap)
      local all = pandoc.utils.stringify(b)
      add("tables", all:sub(#pandoc.utils.stringify(b.caption.long or {}) + 1))
    end
  end
end

function Pandoc(doc)
  local list = doc.meta["count-exclude"] and pandoc.utils.stringify(doc.meta["count-exclude"]) or ""
  for name in list:gmatch("[^|]+") do excluded[pandoc.text.lower(name)] = true end
  count_blocks(doc.blocks, "body")
  local out = {}
  for _, c in ipairs(cats) do
    out[#out + 1] = c .. "=" .. counts[c][1] .. "," .. counts[c][2] .. "," .. counts[c][3]
  end
  out[#out + 1] = "citations=" .. cites
  return pandoc.Pandoc({ pandoc.Plain({ pandoc.Str(table.concat(out, ";")) }) })
end
"""


def run_count(args, settings: dict, src: Path) -> None:
    """Words and characters in the clean text: changes accepted, comments
    removed, citations and the reference list left out."""
    pandoc = find_tool(settings["tools"]["pandoc"], "pandoc", "Install Pandoc.")
    doc = Document(src, settings)
    text = doc.render("accept", "count", "", "")
    with tempfile.TemporaryDirectory() as td:
        counter = Path(td) / "count.lua"
        counter.write_text(COUNT_FILTER, encoding="utf-8")
        crossref = Path(td) / "crossref.lua"
        crossref.write_text(CROSSREF_FILTER, encoding="utf-8")
        exclude = "|".join(settings["count"]["exclude_sections"])
        cmd = [pandoc, "--from", "markdown+mark", "--strip-comments", "--to", "plain",
               "--metadata", f"count-exclude={exclude}"]
        if settings["crossref"]["enabled"]:
            cmd += ["--lua-filter", str(crossref)]
        cmd += ["--lua-filter", str(counter)]
        result = subprocess.run(cmd, input=text.encode("utf-8"), cwd=src.parent,
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode:
        sys.stderr.write(result.stderr.decode("utf-8", errors="replace"))
        sys.exit("Counting failed.")
    data = {}
    for part in result.stdout.decode("utf-8").strip().split(";"):
        key, _, value = part.partition("=")
        data[key] = [int(v) for v in value.split(",")]

    labels = {"body": "Main text", "headings": "Headings", "captions": "Captions",
              "tables": "Tables", "footnotes": "Footnotes"}
    total = [0, 0, 0]
    print(f"{src.name}  (changes accepted; comments, citations and reference sections left out)\n")
    print(f"{'':<12}{'Words':>9}{'Characters':>13}{'without spaces':>17}")
    for key, label in labels.items():
        w, c, cs = data.get(key, [0, 0, 0])
        if w or c:
            print(f"{label:<12}{w:>9,}{c:>13,}{cs:>17,}".replace(",", " "))
            total = [total[0] + w, total[1] + c, total[2] + cs]
    print(f"{'Total':<12}{total[0]:>9,}{total[1]:>13,}{total[2]:>17,}".replace(",", " "))
    cites = data.get("citations", [0])[0]
    if cites:
        print(f"\n{cites} citation(s) not counted.")


# ---------------------------------------------------------------------------
# --import: bring a reviewer's Word changes back as CriticMarkup
# ---------------------------------------------------------------------------


def _norm(word: str) -> str:
    return re.sub(r"[\W_]+", "", word.lower())


class Review:
    """A reviewed Word file as two parallel texts: `old`, as the reviewer
    received it, and `new`, with the reviewer's changes applied. Changes
    and comments by the manuscript's own author are treated as plain text,
    since they are already in the manuscript as CriticMarkup."""

    def __init__(self, pandoc: str, docx: Path, own: set[str]) -> None:
        result = subprocess.run([pandoc, str(docx), "--from", "docx", "--track-changes=all",
                                 "--to", "json"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode:
            sys.stderr.write(result.stderr.decode("utf-8", errors="replace"))
            sys.exit(f"Could not read {docx.name}.")
        self.own = own
        self.old: list[str] = []
        self.new: list[str] = []
        self.old_len = 0
        self.new_len = 0
        self.changes: list[dict] = []        # {author, o1, o2, n1, n2}
        self.comments: dict[str, dict] = {}  # id -> {author, text, o1, o2}
        self.skipped = 0
        self._walk(json.loads(result.stdout)["blocks"])
        self.old_text = "".join(self.old)
        self.new_text = "".join(self.new)

    def _is_own(self, author: str) -> bool:
        return not author.strip() or author.strip().lower() in self.own

    def _emit(self, text: str, kind: str, author: str) -> None:
        if kind != "keep":
            last = self.changes[-1] if self.changes else None
            if last and last["author"] == author and last["o2"] == self.old_len \
                    and last["n2"] == self.new_len:
                change = last
            else:
                change = {"author": author, "o1": self.old_len, "o2": self.old_len,
                          "n1": self.new_len, "n2": self.new_len}
                self.changes.append(change)
        if kind != "ins":
            self.old.append(text)
            self.old_len += len(text)
        if kind != "del":
            self.new.append(text)
            self.new_len += len(text)
        if kind != "keep":
            change["o2"], change["n2"] = self.old_len, self.new_len

    @staticmethod
    def _plain(node) -> str:
        out: list[str] = []

        def walk(n):
            if isinstance(n, dict):
                t, c = n.get("t"), n.get("c")
                if t == "Str":
                    out.append(c)
                elif t in ("Space", "SoftBreak", "LineBreak"):
                    out.append(" ")
                elif t in ("Code", "Math"):
                    out.append(c[1])
                else:
                    walk(c)
            elif isinstance(n, list):
                for x in n:
                    walk(x)
        walk(node)
        return re.sub(r"\s+", " ", "".join(out)).strip()

    def _walk(self, node, kind: str = "keep", author: str = "", skip: bool = False) -> None:
        if isinstance(node, list):
            for x in node:
                self._walk(x, kind, author, skip)
            return
        if not isinstance(node, dict):
            return
        t, c = node.get("t"), node.get("c")
        if t in ("Str", "Code", "Math"):
            if not skip:
                self._emit(c if t == "Str" else c[1], kind, author)
        elif t in ("Space", "SoftBreak", "LineBreak"):
            if not skip:
                self._emit(" ", kind, author)
        elif t == "Span":
            (_, classes, attrs), content = c[0], c[1]
            attrs = dict(attrs)
            who = attrs.get("author", "")
            if "comment-start" in classes:
                if self._is_own(who):
                    return
                if skip:
                    self.skipped += 1
                    return
                self.comments[attrs.get("id", str(len(self.comments)))] = {
                    "author": who, "text": self._plain(content),
                    "o1": self.old_len, "o2": self.old_len}
            elif "comment-end" in classes:
                comment = self.comments.get(attrs.get("id", ""))
                if comment is not None and not skip:
                    comment["o2"] = self.old_len
            elif "insertion" in classes or "deletion" in classes:
                if self._is_own(who):
                    self._walk(content, kind, author, skip)
                elif skip:
                    self.skipped += 1
                else:
                    self._walk(content, "ins" if "insertion" in classes else "del", who, skip)
            else:
                self._walk(content, kind, author, skip)
        elif t == "Note":
            self._walk(c, kind, author, True)               # footnotes are not imported
        elif t == "Div" and c[0][0] == "refs":
            self._walk(c[1], kind, author, True)            # nor the reference list
        else:
            block = t in ("Para", "Plain", "Header", "CodeBlock", "Div", "BlockQuote",
                          "BulletList", "OrderedList", "Table", "Figure", "LineBlock")
            if block and not skip:
                self._emit("\n", "keep", "")
            self._walk(c, kind, author, skip)
            if block and not skip:
                self._emit("\n", "keep", "")


FIELD_TOKEN = re.compile(r'<w:fldChar\b[^>]*w:fldCharType="(begin|separate|end)"[^>]*>'
                         r'|<w:instrText\b[^>]*>(.*?)</w:instrText>'
                         r'|<w:t\b[^>]*>(.*?)</w:t>'
                         r'|<w:fldSimple\b[^>]*w:instr="([^"]*)"[^>]*>(.*?)</w:fldSimple>', re.DOTALL)


def zotero_fields_in_docx(docx: Path) -> list[tuple[str, dict]]:
    """The Zotero citations in a Word file: the text Word shows for each, and
    the citation data in its field code. Read from the XML, since Pandoc does
    not see fields inside tracked insertions."""
    import html
    import zipfile
    try:
        with zipfile.ZipFile(docx) as z:
            xml = "".join(z.read(n).decode("utf-8") for n in z.namelist()
                          if re.fullmatch(r"word/(document|footnotes|endnotes)\.xml", n))
    except (OSError, zipfile.BadZipFile):
        return []
    fields: list[dict] = []
    stack: list[dict] = []
    for m in FIELD_TOKEN.finditer(xml):
        kind, code, text, simple_code, simple_body = m.groups()
        if kind == "begin":
            stack.append({"code": "", "shown": "", "result": False})
        elif kind == "separate" and stack:
            stack[-1]["result"] = True
        elif kind == "end" and stack:
            fields.append(stack.pop())
            if stack and stack[-1]["result"]:
                stack[-1]["shown"] += fields[-1]["shown"]
        elif code is not None and stack and not stack[-1]["result"]:
            stack[-1]["code"] += code
        elif text is not None and stack and stack[-1]["result"]:
            stack[-1]["shown"] += text
        elif simple_code is not None:
            fields.append({"code": simple_code, "result": True,
                           "shown": "".join(re.findall(r"<w:t\b[^>]*>(.*?)</w:t>", simple_body))})
    out = []
    decoder = json.JSONDecoder()
    for f in fields:
        code = html.unescape(f["code"])
        m = re.search(r"ADDIN ZOTERO_ITEM CSL_CITATION\s*", code)
        if m:
            try:
                out.append((html.unescape(f["shown"]), decoder.raw_decode(code, m.end())[0]))
            except ValueError:
                pass
    return out


def zotero_item_id(ci: dict) -> str:
    d = ci.get("itemData", {})
    return str(d.get("citation-key") or d.get("id") or ci.get("id"))


def _title_year(item: dict) -> tuple[str, str]:
    year = str(((item.get("issued") or {}).get("date-parts") or [[""]])[0][0])
    return _norm(item.get("title", "")), year


def resolve_zotero_ids(fields: dict[str, dict], bibs: list[Path]) -> dict[str, str]:
    """Citation keys in your bibliography for the items in Zotero fields:
    by citation key, DOI, Zotero item key (via Better BibTeX, if Zotero is
    running) or title and year. Items that cannot be matched are left out."""
    entries: list[dict] = []
    for b in bibs:
        if b.suffix.lower() == ".json" and b.exists():
            try:
                entries += json.loads(b.read_text(encoding="utf-8-sig"))
            except ValueError:
                pass
    known = {e.get("id") for e in entries}
    for b in bibs:
        if b.exists():
            known.update(bibliography_keys(b))
    by_doi = {e["DOI"].lower(): e["id"] for e in entries if e.get("DOI")}
    by_title = {_title_year(e): e["id"] for e in entries if e.get("title")}

    out: dict[str, str] = {}
    pending: dict[str, str] = {}
    for pid, f in fields.items():
        d = f["itemData"]
        doi = (d.get("DOI") or "").lower()
        if pid in known:
            out[pid] = pid
        elif d.get("citation-key") in known:
            out[pid] = d["citation-key"]
        elif doi in by_doi:
            out[pid] = by_doi[doi]
        elif f["uris"]:
            pending[pid] = f["uris"][0].rstrip("/").rsplit("/", 1)[-1]
        if pid not in out and pid not in pending and d.get("title") and _title_year(d) in by_title:
            out[pid] = by_title[_title_year(d)]
    if pending:
        found = bbt_call("item.citationkey", sorted(set(pending.values()))) or {}
        for pid, key in pending.items():
            if found.get(key):
                out[pid] = found[key]
            elif fields[pid]["itemData"].get("title") and \
                    _title_year(fields[pid]["itemData"]) in by_title:
                out[pid] = by_title[_title_year(fields[pid]["itemData"])]
    return out


# CSL locator labels as written in a manuscript; pages need no label.
MD_LOCATOR_LABELS = {"page": "", "chapter": "chap.", "section": "sec.", "paragraph": "para.",
                     "figure": "fig.", "volume": "vol.", "note": "n.", "line": "l.",
                     "issue": "no.", "column": "col.", "book": "bk.", "verse": "v."}


def zotero_markdown(citation: dict, keys: dict[str, str]) -> str | None:
    """A Pandoc citation for a Zotero field, or None if an item is not in
    your bibliography."""
    parts = []
    for ci in citation.get("citationItems", []):
        key = keys.get(zotero_item_id(ci))
        if not key:
            return None
        label = MD_LOCATOR_LABELS.get(ci.get("label", "page"), ci.get("label", ""))
        locator = f"{label} {ci['locator']}".strip() if ci.get("locator") else ""
        tail = ", ".join(x for x in (locator, ci.get("suffix", "").strip().lstrip(", ")) if x)
        prefix = ci.get("prefix", "").strip()
        parts.append((f"{prefix} " if prefix else "") + ("-" if ci.get("suppress-author") else "")
                     + f"@{key}" + (f", {tail}" if tail else ""))
    return "[" + "; ".join(parts) + "]" if parts else None


def _word_bounds(text: str, a: int, b: int) -> tuple[int, int]:
    """Widen [a, b) to whole words (runs of non-space characters)."""
    while a > 0 and not text[a - 1].isspace():
        a -= 1
    while b < len(text) and not text[b].isspace():
        b += 1
    return a, b


def _boundary(s: str, i: int) -> bool:
    return i == 0 or i == len(s) or not (s[i - 1].isalnum() and s[i].isalnum())


def _trim(old: str, new: str) -> tuple[int, int]:
    """Common prefix and suffix lengths, cut back to word boundaries."""
    p = len(os.path.commonprefix([old, new]))
    while p and not (_boundary(old, p) and _boundary(new, p)):
        p -= 1
    q = len(os.path.commonprefix([old[p:][::-1], new[p:][::-1]]))
    while q and not (_boundary(old, len(old) - q) and _boundary(new, len(new) - q)):
        q -= 1
    return p, q


CITATION_AHEAD = re.compile(r"(?:[ \t]*(?:\[[^\[\]\n]*@[^\[\]\n]*\]|-?@[\w:.#$%&+?<>~/-]+))+[.,;:!?)]*")


def mask_source(raw: str) -> str:
    """The manuscript with markup blanked out (same length, same line
    breaks), so its words line up with the text Word shows."""
    chars = list(raw)

    def blank(a: int, b: int) -> None:
        for i in range(a, b):
            if chars[i] != "\n":
                chars[i] = " "

    head = YAML_HEAD.match(raw)
    if head:
        blank(0, head.end())
    for pattern in (r"\{>>.*?<<\}", r"%%.*?%%", r"<!--.*?-->",
                    r"\[[^\[\]]*@[^\[\]]*\]", r"(?<![\w.])@[\w:.#$%&+?<>~/-]+",
                    r"\]\([^)]*\)", r"\{[#.][^}\n]*\}", r"!\[\[[^\]]*\]\]",
                    r"\[\[[^\]|]*\|", r"\\newpage|\\pagebreak|\\clearpage|\{\{<[^>]*>\}\}",
                    r"\{\+\+|\+\+\}|\{--|--\}|\{~~|~~\}|~>|\{==|==\}", r"[*_#>|\[\]`^~=\\]"):
        for m in re.finditer(pattern, raw, re.DOTALL):
            blank(m.start(), m.end())
    return "".join(chars)


def run_import(args, settings: dict, src: Path) -> None:
    docx = Path(args.import_docx).resolve()
    if not docx.exists():
        sys.exit(f"File not found: {docx}")
    pandoc = find_tool(settings["tools"]["pandoc"], "pandoc", "Install Pandoc.")
    own = {n.strip().lower() for n in [settings["general"]["author"],
                                       *settings["import"]["own_authors"]] if n.strip()}
    review = Review(pandoc, docx, own)
    # Zotero citations a reviewer inserted show up as their text; they are
    # turned back into [@key] where the item is in your bibliography.
    zotero = zotero_fields_in_docx(docx)
    cite_md: list[tuple[re.Pattern, str | None, str]] = []
    if zotero:
        head = split_yaml(src.read_text(encoding="utf-8-sig"))[0]
        try:
            _, bibs = citation_setup(args, settings, head)
        except SystemExit:
            bibs = []
        items = {zotero_item_id(ci): {"itemData": ci.get("itemData", {}),
                                      "uris": ci.get("uris") or ci.get("uri") or []}
                 for _, c in zotero for ci in c.get("citationItems", [])}
        keys = resolve_zotero_ids(items, bibs or yaml_bibliographies(head, src.parent))
        for shown, c in sorted(zotero, key=lambda z: -len(z[0])):
            words = shown.replace("\xa0", " ").split()
            if words:
                cite_md.append((re.compile(r"\s+".join(map(re.escape, words))),
                                zotero_markdown(c, keys), " ".join(words)))
    unmatched: list[str] = []

    def with_citations(text: str) -> str:
        for pattern, md, shown in cite_md:
            if pattern.search(text):
                if md:
                    text = pattern.sub(lambda _: md, text)
                elif shown not in unmatched:
                    unmatched.append(shown)
        return text
    # Non-breaking spaces come from edge spaces of exported changes (and from
    # Word); same length, so the two texts stay aligned.
    old, new = review.old_text.replace("\xa0", " "), review.new_text.replace("\xa0", " ")

    raw = src.read_text(encoding="utf-8-sig")
    masked = mask_source(raw)
    src_tokens = [(m.start(), m.end(), _norm(m.group(0))) for m in re.finditer(r"\S+", masked)]
    src_tokens = [t for t in src_tokens if t[2]]
    base_tokens = [(m.start(), m.end(), _norm(m.group(0))) for m in re.finditer(r"\S+", old)]

    # Line up the words the reviewer saw with the words of the manuscript.
    usable = [i for i, t in enumerate(base_tokens) if t[2]]
    matcher = difflib.SequenceMatcher(None, [base_tokens[i][2] for i in usable],
                                      [t[2] for t in src_tokens])
    to_src: dict[int, int] = {}
    for a, b, n in matcher.get_matching_blocks():
        for k in range(n):
            to_src[usable[a + k]] = b + k

    def tokens_in(o1: int, o2: int) -> list[int]:
        return [i for i, (s, e, _) in enumerate(base_tokens) if s < o2 and e > o1]

    def src_after_token_before(o: int) -> int | None:
        """Source position after the base word that contains or precedes o."""
        before = [i for i, (s, e, _) in enumerate(base_tokens) if s < o]
        if before and before[-1] in to_src:
            return src_tokens[to_src[before[-1]]][1]
        # The words just before are a rendered citation, which is not in the
        # manuscript as text: go past the [@key] that follows the last known word.
        known_before = [i for i in before if i in to_src]
        if known_before:
            end = src_tokens[to_src[known_before[-1]]][1]
            m = CITATION_AHEAD.match(raw, end)
            if m:
                return m.end()
        return None

    def src_before_token_after(o: int) -> int | None:
        after = [i for i, (s, e, _) in enumerate(base_tokens) if s >= o]
        return src_tokens[to_src[after[0]]][0] if after and after[0] in to_src else None

    def nearest(o: int) -> int:
        for i in reversed([i for i, (s, e, _) in enumerate(base_tokens) if e <= o]):
            if i in to_src:
                return src_tokens[to_src[i]][1]
        return len(raw.rstrip())

    def mapped_span(o1: int, o2: int) -> tuple[int, int] | None:
        """Source span for the base text [o1, o2), if every word in it was
        found in the manuscript (punctuation-only tokens may be missing)."""
        idx = tokens_in(o1, o2)
        if not idx or any(base_tokens[i][2] and i not in to_src for i in idx):
            return None
        mapped = [to_src[i] for i in idx if i in to_src]
        if not mapped or mapped != list(range(mapped[0], mapped[-1] + 1)):
            return None
        s, e = src_tokens[mapped[0]][0], src_tokens[mapped[-1]][1]
        if re.search(r"\{\+\+|\{--|\{~~|\{>>|\{==|\n\s*\n|%%|<!--", raw[s:e]):
            return None
        return s, e

    # Changes and comments already in the manuscript (from an earlier import,
    # exported again and sent back) are recognised and not added twice.
    reviewers = settings["critic"]["reviewers"]
    known: set[tuple[str, str, str]] = set()
    for m in CRITIC.finditer(raw):
        g = m.groupdict()
        if g["ins"] is not None:
            known.add(("insertion", (g["ins_by"] or "").lower(), _norm(g["ins"])))
        elif g["dele"] is not None:
            known.add(("deletion", (g["del_by"] or "").lower(), _norm(g["dele"])))
        elif g["sub_old"] is not None:
            known.add(("replacement", (g["sub_by"] or "").lower(),
                       _norm(g["sub_old"]) + ">" + _norm(g["sub_new"])))
        else:
            who, text = comment_author(g["comment"] or g["hl_comment"] or "", reviewers)
            known.add(("comment", (who or "").lower(), _norm(text)))
    already = 0

    ops: list[tuple[int, int, str]] = []
    unplaced: list[str] = []
    stats: dict[str, dict[str, int]] = {}

    def count_stat(author: str, kind: str) -> None:
        stats.setdefault(author, {}).setdefault(kind, 0)
        stats[author][kind] += 1

    def fallback(author: str, what: str, o: int) -> None:
        pos = nearest(o)
        ops.append((pos, pos, f"{{>>@{author}: [critic-export could not place this] {what}<<}}"))
        unplaced.append(f"{author}: {what}")

    # Widen each change to whole words and merge changes that then touch.
    edits: list[dict] = []
    for ch in review.changes:
        a, b = ch["o1"], ch["o2"]
        inserted = new[ch["n1"]:ch["n2"]]
        # " more words" added after "end." starts a new word, so it stays an insertion.
        separate = b == a and ((inserted[:1].isspace() and (a == len(old) or old[a].isspace()))
                               or (inserted[-1:].isspace() and (a == 0 or old[a - 1].isspace())))
        if not separate and (b > a or (a > 0 and not old[a - 1].isspace())
                             or (a < len(old) and not old[a].isspace())):
            # A change inside or next to a word covers the whole word, so that
            # "design" -> "designs" becomes {~~design~>designs~~}.
            a, b = _word_bounds(old, a, b)
        n1, n2 = ch["n1"] - (ch["o1"] - a), ch["n2"] + (b - ch["o2"])
        if edits and edits[-1]["author"] == ch["author"] and a <= edits[-1]["o2"]:
            last = edits[-1]
            last["o2"], last["n2"] = max(last["o2"], b), max(last["n2"], n2)
        else:
            edits.append({"author": ch["author"], "o1": a, "o2": b, "n1": n1, "n2": n2})

    for ed in edits:
        author = ed["author"]
        old_part, new_part = old[ed["o1"]:ed["o2"]], with_citations(new[ed["n1"]:ed["n2"]])
        if not old_part.strip():
            # Insertion between words: attach after the previous word.
            text = new_part.strip()
            if not text:
                continue
            if ("insertion", author.lower(), _norm(text)) in known:
                already += 1
                continue
            after = src_after_token_before(ed["o1"])
            if after is not None:
                ops.append((after, after, f"{{++ {text}++}}{{>>@{author}<<}}"))
            else:
                before = src_before_token_after(ed["o2"])
                if before is None:
                    fallback(author, f'inserted "{text}"', ed["o1"])
                    continue
                ops.append((before, before, f"{{++{text} ++}}{{>>@{author}<<}}"))
            count_stat(author, "insertion")
            continue

        p0, q0 = _trim(old_part.strip(), new_part.strip())
        o0 = old_part.strip()[p0:len(old_part.strip()) - q0]
        n0 = new_part.strip()[p0:len(new_part.strip()) - q0]
        guess = ("replacement", _norm(o0) + ">" + _norm(n0)) if o0 and n0 else \
            ("deletion", _norm(o0)) if o0 else ("insertion", _norm(n0))
        if (guess[0], author.lower(), guess[1]) in known:
            already += 1
            continue
        span = mapped_span(ed["o1"], ed["o2"])
        if span is None:
            what = (f'replaced "{old_part.strip()}" with "{new_part.strip()}"' if new_part.strip()
                    else f'deleted "{old_part.strip()}"')
            fallback(author, what, ed["o1"])
            continue
        s, e = span
        source = raw[s:e]
        old_core, new_core = old_part.strip(), new_part.strip()
        if source == old_core:
            p, q = _trim(old_core, new_core)
            pre, suf = source[:p], source[len(source) - q:] if q else ""
            o_mid, n_mid = old_core[p:len(old_core) - q], new_core[p:len(new_core) - q]
        else:                                    # markup inside: replace the whole span
            pre = suf = ""
            o_mid, n_mid = source, new_core
        if o_mid and n_mid:
            markup, kind = f"{{~~{o_mid}~>{n_mid}~~}}", "replacement"
            key = _norm(o_mid) + ">" + _norm(n_mid)
        elif o_mid:
            markup, kind, key = f"{{--{o_mid}--}}", "deletion", _norm(o_mid)
        elif n_mid:
            markup, kind, key = f"{{++{n_mid}++}}", "insertion", _norm(n_mid)
        else:
            continue
        if (kind, author.lower(), key) in known:
            already += 1
            continue
        ops.append((s, e, f"{pre}{markup}{{>>@{author}<<}}{suf}"))
        count_stat(author, kind)

    for com in review.comments.values():
        author = com["author"] or "Reviewer"
        text = com["text"].replace("<<}", "<< }")
        note = f"{{>>@{author}: {text}<<}}"
        if ("comment", author.lower(), _norm(com["text"])) in known:
            already += 1
            continue
        count_stat(author, "comment")
        if com["o2"] > com["o1"]:
            span = mapped_span(*_word_bounds(old, com["o1"], com["o2"]))
            if span:
                s, e = span
                ops.append((s, e, f"{{=={raw[s:e]}==}}{note}"))
                continue
        at = src_after_token_before(com["o2"] if com["o2"] > com["o1"] else com["o1"])
        if at is None:
            at = src_before_token_after(com["o1"])
        if at is None:
            pos = nearest(com["o1"])
            ops.append((pos, pos, f"{{>>@{author}: [critic-export could not place this] {text}<<}}"))
            unplaced.append(f'{author}: comment "{text[:60]}"')
        else:
            ops.append((at, at, note))

    # Apply from the end. A replacement overlapping an earlier one is
    # turned into a comment, so nothing is lost.
    ops.sort(key=lambda o: (o[0], o[1]))
    accepted: list[tuple[int, int, str]] = []
    last_end = -1
    for s, e, text in ops:
        if e > s and s < last_end:
            accepted.append((e, e, text if text.startswith("{>>") else
                             "{>>@critic-export: overlapping change from the review: "
                             + text.replace("<<}", "<< }") + "<<}"))
            unplaced.append(f'overlapping change near "{raw[s:e][:40]}"')
            continue
        accepted.append((s, e, text))
        last_end = max(last_end, e)
    out = raw
    for s, e, text in sorted(accepted, key=lambda o: (o[0], o[1]), reverse=True):
        out = out[:s] + text + out[e:]

    stamp = datetime.now().strftime("%y%m%dT%H%M%S")
    backup = None
    if args.in_place:
        backup_dir = expand(settings["general"]["output_dir"], src.parent) \
            if settings["general"]["output_dir"] else src.parent
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup = backup_dir / f"{src.stem}@before-import-{stamp}{src.suffix}"
        shutil.copyfile(src, backup)
        target = src
    else:
        target = args.output.resolve() if args.output else \
            src.with_name(f"{src.stem}@import-{stamp}{src.suffix}")
    target.write_text(out, encoding="utf-8")

    print(f"Imported the review in {docx.name} into {target}")
    if backup:
        print(f"  (the previous version is saved as {backup})")
    if not stats:
        print("  No new changes or comments from other people were found.")
    if already:
        print(f"  {already} change(s) or comment(s) already in the manuscript were skipped.")
    for author, kinds in stats.items():
        listed = ", ".join(f"{n} {k}{'s' if n > 1 else ''}" for k, n in kinds.items())
        print(f"  {author}: {listed}")
    if unplaced:
        print(f"Warning: {len(unplaced)} item(s) could not be placed exactly; they were added as "
              "comments at the nearest position:", file=sys.stderr)
        for u in unplaced[:20]:
            print(f"  {u}", file=sys.stderr)
    if unmatched:
        print(f"Warning: {len(unmatched)} inserted Zotero citation(s) are not in your "
              f"bibliography and were imported as text: {'; '.join(unmatched)}",
              file=sys.stderr)
    if review.skipped:
        print(f"Warning: {review.skipped} change(s) or comment(s) in footnotes or the reference list "
              "were not imported.", file=sys.stderr)


# ---------------------------------------------------------------------------
# Shared document preparation
# ---------------------------------------------------------------------------


class Document:
    """A manuscript read from disk, with embedded Obsidian notes expanded,
    ready to be rendered for one format and CriticMarkup mode."""

    def __init__(self, src: Path, settings: dict) -> None:
        self.src = src
        self.settings = settings
        self.head, body = split_yaml(src.read_text(encoding="utf-8-sig"))
        vault_root = find_vault(src.parent)
        choice = settings["_obsidian_override"] or settings["obsidian"]["enabled"]
        self.obsidian = None
        if obsidian_enabled(choice, vault_root):
            self.obsidian = ObsidianConverter(Vault(vault_root or src.parent), src.parent,
                                              settings["obsidian"])
            body = self.obsidian.expand_embeds(body, src.parent, (src,))
        self.body = body

    def render(self, mode: str, fmt: str, author: str, date: str) -> str:
        """Order matters: code is protected first; Obsidian comments go before
        CriticMarkup (they may contain it); CriticMarkup goes before other
        Obsidian syntax ({==highlight==} must not be read as ==highlight==)."""
        guard = CodeGuard()
        text = guard.protect(self.body)
        if self.obsidian:
            text = self.obsidian.strip_comments(text)
        text = fix_locators(text, self.settings["locators"])
        if self.settings["syntax"]["page_breaks"]:
            text = mark_page_breaks(text)
        text = convert(text, mode, fmt, author, date, self.settings["critic"])
        if self.obsidian:
            text = self.obsidian.convert(text)
        return self.head + guard.restore(text)


BBT_RPC = "http://127.0.0.1:23119/better-bibtex/json-rpc"


def bbt_call(method: str, *params):
    """Call Better BibTeX in the running Zotero. None if Zotero is not running."""
    import urllib.request
    body = json.dumps({"jsonrpc": "2.0", "method": method, "params": list(params)}).encode()
    req = urllib.request.Request(BBT_RPC, body, {"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return json.loads(r.read().decode("utf-8")).get("result")
    except (OSError, ValueError):
        return None


def zotero_field_setup(text: str, head: str, citeproc_args: list[str], td: Path,
                       filter_path: Path) -> list[str] | None:
    """Pandoc options that make Word citations Zotero fields: the item data
    for the document's citation keys from Better BibTeX, and Zotero's
    document preferences (style and language). None if Zotero is not running."""
    keys = list(dict.fromkeys(k for k in CITEKEY.findall(text) if not CROSSREF_KEY.match(k)))
    if not keys:
        return []
    result = bbt_call("item.pandoc_filter", keys, True)
    if result is None:
        print("Warning: Zotero is not running, so the citations in Word are plain text. "
              "Start Zotero and export again for Zotero fields.", file=sys.stderr)
        return None
    items = result.get("items", {})
    missing = [k for k in keys if k not in items]
    if missing:
        print(f"Note: not found in Zotero, left as plain text: {', '.join('@' + k for k in missing)}",
              file=sys.stderr)
    data = td / "zotero-items.json"
    data.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")

    style_id = "http://www.zotero.org/styles/chicago-author-date"   # Pandoc's default style
    csl = citeproc_args[citeproc_args.index("--csl") + 1] if "--csl" in citeproc_args else None
    if csl and Path(csl).exists():
        m = re.search(r"<id>\s*(.*?)\s*</id>", Path(csl).read_text(encoding="utf-8", errors="replace"))
        style_id = m.group(1) if m else style_id
    lang = re.search(r"^lang\s*:\s*['\"]?([\w-]+)", head, re.MULTILINE)
    session = "".join(random.choices("abcdefghijklmnopqrstuvwxyz0123456789", k=8))
    prefs = (f'<data data-version="3" zotero-version="7.0"><session id="{session}"/>'
             f'<style id="{style_id}" locale="{lang.group(1) if lang else "en-US"}" '
             'hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
             '<prefs><pref name="fieldType" value="Field"/></prefs></data>')
    os.environ["CRITIC_EXPORT_ZOTERO_ITEMS"] = str(data)
    args = ["--lua-filter", str(filter_path)]
    for n, i in enumerate(range(0, len(prefs), 255), 1):            # Word's limit per property
        args += ["--metadata", f"ZOTERO_PREF_{n}={prefs[i:i + 255]}"]
    return args


def citation_setup(args, settings: dict, head: str) -> tuple[list[str], list[Path]]:
    """Pandoc citeproc options and the bibliography files in use.
    Document YAML wins over settings; command-line options win over both."""
    cit, base = settings["citations"], settings["_base"]
    doc_has_bib = re.search(r"^bibliography\s*:", head, re.MULTILINE) is not None
    doc_has_csl = re.search(r"^csl\s*:", head, re.MULTILINE) is not None

    citeproc_args = ["--citeproc"]
    if args.bibliography:
        bibs = resolve_bibliographies(args.bibliography, Path.cwd())
    elif not doc_has_bib:
        bibs = resolve_bibliographies(cit["bibliography"], base)
    else:
        bibs = []                                  # the document lists its own
    for b in bibs:
        citeproc_args += ["--bibliography", str(b)]
    style = args.style or ("" if doc_has_csl else cit["style"])
    if style:
        style_base = Path.cwd() if args.style else base
        citeproc_args += ["--csl", str(resolve_style(style, settings, style_base))]
    if cit["link_citations"]:
        citeproc_args += ["--metadata", "link-citations=true"]
    return citeproc_args, bibs


def main() -> None:
    args = parse_args()
    settings_path = find_settings(args.settings)
    settings = load_settings(settings_path)
    settings["_base"] = settings_path.parent if settings_path else Path.cwd()
    settings["_path"] = settings_path
    settings["_obsidian_override"] = args.obsidian

    src = args.input.resolve() if args.input else None
    if src:
        project = find_project_settings(src.parent)
        if project:
            merge_project_settings(settings, project)

    if args.list_styles:
        list_styles(settings)
        return
    if args.show_settings:
        show_settings(settings_path, settings)
        return
    if args.update_settings:
        if not settings_path:
            sys.exit("No settings file to update.")
        update_settings(settings_path)
        return
    if args.install_context_menu:
        install_context_menu(settings)
        return
    if args.uninstall_context_menu:
        uninstall_context_menu(settings)
        return
    if args.start:
        run_start(settings)
        return
    if not src:
        sys.exit("No input file given. Run with --help for usage.")
    if not src.exists():
        sys.exit(f"File not found: {src}")

    if args.import_docx:
        run_import(args, settings, src)
    elif args.check:
        sys.exit(run_check(args, settings, src))
    elif args.count:
        run_count(args, settings, src)
    else:
        sys.exit(run_export(args, settings, src))


def run_export(args, settings: dict, src: Path) -> int:
    g, crit = settings["general"], settings["critic"]
    settings_base = settings["_base"]
    formats = [f.lower().lstrip(".") for f in (args.to or g["formats"])]
    formats = list(dict.fromkeys(formats))
    if args.output and len(formats) > 1:
        sys.exit("-o can only be used with a single format.")

    author = args.author if args.author is not None else g["author"]
    pandoc = find_tool(settings["tools"]["pandoc"], "pandoc",
                       "Install Pandoc (https://pandoc.org/installing.html) "
                       "or set [tools] pandoc in settings.toml.")

    doc = Document(src, settings)
    ask = bool(ASK_METADATA & set(formats))
    doc.head = add_to_yaml(doc.head, ask_metadata(doc.head, doc.body, args, src, ask,
                                                   g["document_author"]))
    doc.body = drop_title_heading(doc.head, doc.body)
    try:
        citeproc_args, bibs = citation_setup(args, settings, doc.head)
    except SystemExit as e:
        # A missing bibliography only matters when there is something to cite.
        if any(not CROSSREF_KEY.match(k) for k in CITEKEY.findall(doc.body)):
            raise
        print(f"Warning: {e.code}\nNo citations in the document; exporting without a bibliography.",
              file=sys.stderr)
        citeproc_args, bibs = [], []
    warn_duplicate_keys(bibs)

    stamp = datetime.now().strftime(g["timestamp"]) if g["timestamp"] else ""
    date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    out_dir = expand(g["output_dir"], src.parent) if g["output_dir"] else src.parent
    name = g["filename"].format(stem=src.stem, timestamp=stamp)
    if not stamp:
        name = name.rstrip("@_- ")

    failures = 0
    exported: list[Path] = []
    with tempfile.TemporaryDirectory() as td:
        lua = Path(td) / "critic.lua"
        lua.write_text(LUA_FILTER, encoding="utf-8")
        pagebreak = Path(td) / "pagebreak.lua"
        pagebreak.write_text(PAGEBREAK_FILTER, encoding="utf-8")
        crossref = Path(td) / "crossref.lua"
        crossref.write_text(CROSSREF_FILTER, encoding="utf-8")
        crossref_args: list[str] = []
        if settings["crossref"]["enabled"]:
            # Must run before citeproc, which would read @fig:id as a citation.
            crossref_args = ["--lua-filter", str(crossref)]
            for key, value in settings["crossref"].items():
                if key != "enabled" and value:
                    crossref_args += ["--metadata", f"crossref-{key.replace('_', '-')}={value}"]
        zotero_lua = Path(td) / "zotero.lua"
        zotero_lua.write_text(ZOTERO_FILTER, encoding="utf-8")
        css = Path(td) / "critic.css"
        css.write_text(CRITIC_CSS, encoding="utf-8")
        css_header = Path(td) / "critic-style.html"
        css_header.write_text(f"<style>\n{CRITIC_CSS}</style>\n", encoding="utf-8")

        for fmt in formats:
            writer, ext = FORMAT_ALIASES.get(fmt, (fmt, fmt))
            mode, note = critic_mode_for(fmt, args.critic, settings)

            extra: list[str] = []
            if fmt == "pdf":
                engine = find_tool(settings["tools"]["pdf_engine"], "xelatex",
                                   "PDF needs a LaTeX engine such as TinyTeX or MiKTeX (see README), "
                                   "or set [tools] pdf_engine in settings.toml.")
                engine_name = Path(engine).stem.lower()
                writer = PDF_ENGINE_WRITERS.get(engine_name)
                if mode == "markup" and engine_name not in LATEX_ENGINES and writer != "html5":
                    mode, note = "accept", f"{engine_name} cannot show CriticMarkup; exporting with changes accepted."
                extra += ["--pdf-engine", engine]
                if writer == "html5":
                    extra += ["--css", str(css)]
            if writer in ("html5", "html"):
                extra += ["--embed-resources", "--include-in-header", str(css_header)]
                if settings["html"]["css"]:
                    extra += ["--css", str(expand(settings["html"]["css"], settings_base))]
            elif writer == "epub3":
                extra += ["--css", str(css)]
            elif writer in ("docx", "odt"):
                fields = settings["docx"]["zotero_fields"] if args.zotero_fields is None \
                    else args.zotero_fields
                if writer == "docx" and fields:
                    extra += zotero_field_setup(doc.body, doc.head, citeproc_args, Path(td),
                                                zotero_lua) or []
                ref = args.reference_doc or (settings["docx"]["reference_doc"] if writer == "docx" else "")
                if ref:
                    extra += ["--reference-doc",
                              str(expand(ref, Path.cwd() if args.reference_doc else settings_base))]
            if writer:
                extra += ["--to", writer]
            if note:
                print(f"Note: {note}", file=sys.stderr)

            text = doc.render(mode, fmt, author, date)

            out = args.output.resolve() if args.output else out_dir / f"{name}.{ext}"
            out.parent.mkdir(parents=True, exist_ok=True)
            ensure_writable(out)

            # The intermediate file sits next to the source so relative paths
            # (images etc.) resolve as in the original.
            tmp = src.with_name(f".{src.stem}.critic-export.{ext}.md")
            tmp.write_text(text, encoding="utf-8")

            cmd = [pandoc, str(tmp), "--from", "markdown+mark", "--standalone",
                   *crossref_args, *citeproc_args, *extra, "-o", str(out)]
            if settings["syntax"]["page_breaks"]:
                cmd[-2:-2] = ["--lua-filter", str(pagebreak)]
            if settings["syntax"]["remove_html_comments"]:
                # <!-- comments --> vanish from Word but would travel, hidden,
                # in HTML and EPUB. Pandoc's own option keeps lists separated
                # by a comment apart and leaves code alone.
                cmd.insert(4, "--strip-comments")
            if mode == "markup":
                cmd[-2:-2] = ["--metadata", f"critic-comment-label={crit['comment_label']}",
                              "--lua-filter", str(lua)]
            try:
                result = subprocess.run(cmd, check=True, cwd=src.parent, stderr=subprocess.PIPE)
                if result.stderr:
                    sys.stderr.write(result.stderr.decode("utf-8", errors="replace"))
                if writer == "odt" and settings["syntax"]["page_breaks"]:
                    add_odt_pagebreak_style(out)
                print(f"Exported {fmt} ({mode}): {out}")
                exported.append(out)
            except subprocess.CalledProcessError as e:
                if e.stderr:
                    sys.stderr.write(e.stderr.decode("utf-8", errors="replace"))
                failures += 1
                print(f"Failed: {fmt}. See Pandoc's message above.", file=sys.stderr)
            finally:
                if not args.keep:
                    tmp.unlink(missing_ok=True)

    if args.open or g["open_after_export"]:
        for f in exported:
            open_file(f)
    return 1 if failures else 0


class _StderrWatcher:
    """Notices warnings and errors written to stderr, so --pause-on-issues
    can keep the console window open long enough to read them."""

    def __init__(self, stream) -> None:
        self.stream = stream
        self.issues = False

    def write(self, text: str) -> int:
        if re.search(r"(?im)\bwarning\b|^note:|^failed|error", text):
            self.issues = True
        return self.stream.write(text)

    def __getattr__(self, name):
        return getattr(self.stream, name)


if __name__ == "__main__":
    if "--pause-on-issues" not in sys.argv:
        main()
    else:
        watcher = _StderrWatcher(sys.stderr)
        sys.stderr = watcher
        code = 0
        try:
            main()
        except SystemExit as e:
            if isinstance(e.code, str):
                print(e.code, file=sys.stderr)
                code = 1
            else:
                code = e.code or 0
        except Exception:
            import traceback
            traceback.print_exc()
            code = 1
        if code or watcher.issues:
            input("\nPress Enter to close...")
        sys.exit(code)
