"""Locators, code protection, page breaks and Obsidian syntax, without Pandoc."""
from pathlib import Path


def test_page_locators_become_numbers(module):
    assert module.fix_locators("[@a, {p. 13}; @b, {pp. 1–3}]", {}) == "[@a, 13; @b, 1–3]"


def test_other_locators_are_unwrapped_and_mapped(module):
    assert module.fix_locators("[@a, {chap. 3}]", {}) == "[@a, chap. 3]"
    assert module.fix_locators("[@a, {chap. 3}]", {"chap.": "kap."}) == "[@a, kap. 3]"


def test_code_is_protected(module):
    guard = module.CodeGuard()
    text = "a `{++x++}` b\n\n```\n{--y--}\n```\n"
    protected = guard.protect(text)
    assert "{++" not in protected and "{--" not in protected
    assert guard.restore(protected) == text


def test_page_breaks(module):
    out = module.mark_page_breaks("a\n\n\\newpage\n\nb\n\n{{< pagebreak >}}\n")
    assert out.count("::: pagebreak") == 2


def _converter(module, tmp_path: Path):
    return module.ObsidianConverter(module.Vault(tmp_path), tmp_path,
                                    {"remove_tags": True, "max_embed_depth": 5})


def test_obsidian_inline_syntax(module, tmp_path):
    conv = _converter(module, tmp_path)
    text = ("See [[Note]], [[Note|alias]] and [[Note#Part]]. ==key== idea ^abc12\n"
            "> [!warning] Careful\n> body\n#tag1 #tag2\nIssue #12 stays, #draft goes.")
    out = conv.convert(conv.strip_comments(text))
    assert "See Note, alias and Note > Part." in out
    assert "[key]{.mark} idea" in out and "^abc12" not in out
    assert "> **Careful**" in out
    assert "#tag1" not in out and "Issue #12 stays" in out and "#draft" not in out


def test_obsidian_comments_removed(module, tmp_path):
    conv = _converter(module, tmp_path)
    assert conv.strip_comments("a %%secret\nmore%% b") == "a  b"


def test_embeds_images_and_notes(module, tmp_path):
    (tmp_path / "img").mkdir()
    (tmp_path / "img" / "eye.png").write_bytes(b"")
    (tmp_path / "Part.md").write_text("---\nx: 1\n---\n# A\n\nFirst ^b1\n\n# B\n\nSecond\n",
                                      encoding="utf-8")
    conv = _converter(module, tmp_path)
    out = conv.expand_embeds("![[eye.png|Caption|300]]{#fig:eye}\n\n![[Part#B]]\n\n![[Part#^b1]]",
                             tmp_path)
    assert "![Caption](img/eye.png){#fig:eye width=300px}" in out
    assert "Second" in out and "First" in out and "x: 1" not in out


def test_circular_embed_is_stopped(module, tmp_path):
    (tmp_path / "A.md").write_text("![[B]]", encoding="utf-8")
    (tmp_path / "B.md").write_text("![[A]]", encoding="utf-8")
    conv = _converter(module, tmp_path)
    conv.expand_embeds("![[A]]", tmp_path, (tmp_path / "Main.md",))
    assert any("circular" in w for w in conv.warnings)


def test_yaml_metadata_detection_and_insertion(module):
    head = "---\ntitle: Kapitel\nauthor:\n  - A\n---\n"
    assert module.has_yaml_key(head, "title") and module.has_yaml_key(head, "author")
    assert not module.has_yaml_key("---\ntitle:\nlang: sv\n---\n", "title")
    lines = module.yaml_metadata_lines('Ett "citat"', ["Å", "B"])
    assert lines == ['title: "Ett \\"citat\\""', "author:", '  - "Å"', '  - "B"']
    assert module.add_to_yaml("", ["title: T"]) == "---\ntitle: T\n---\n"
    assert module.add_to_yaml("---\r\nlang: sv\r\n---\r\n", ["title: T"]) == \
        "---\r\nlang: sv\r\ntitle: T\r\n---\r\n"


def test_save_metadata_keeps_bom_and_line_endings(module, tmp_path):
    src = tmp_path / "m.md"
    src.write_bytes(b"\xef\xbb\xbf---\r\nlang: sv\r\n---\r\nText\r\n")
    module.save_metadata(src, ['title: "T"'])
    assert src.read_bytes() == b'\xef\xbb\xbf---\r\nlang: sv\r\ntitle: "T"\r\n---\r\nText\r\n'


def test_obsidian_lists_need_no_blank_lines(module, tmp_path):
    conv = _converter(module, tmp_path)
    text = ("Reasons:\n- one\n  more\n- two\nAfter the list.\n"
            "Year\n2610. not a list\nSteps:\n1. first\n2. second")
    assert conv.convert(text) == ("Reasons:\n\n- one\n  more\n- two\n\nAfter the list.\n"
                                  "Year\n2610. not a list\nSteps:\n\n1. first\n2. second")


def test_title_suggestion_and_duplicate_heading(module, tmp_path, monkeypatch):
    assert module.leading_h1("\n# Etikfall #\n\nText") == "Etikfall"
    assert module.leading_h1("Text\n# Later") is None
    head = '---\ntitle: "Etikfall"\n---\n'
    assert module.drop_title_heading(head, "# Etikfall\n\nText") == "Text"
    assert module.drop_title_heading(head, "# Other\n\nText") == "# Other\n\nText"

    import argparse
    args = argparse.Namespace(title=None, doc_author=None, save_metadata=False, no_prompt=False)
    monkeypatch.setattr("sys.stdin.isatty", lambda: True, raising=False)
    answers = iter(["", "n"])
    monkeypatch.setattr("builtins.input", lambda prompt="": next(answers))
    src = tmp_path / "Filnamn.md"
    assert module.ask_metadata("", "# Rubrik\nText", args, src, True, "A") == \
        ['title: "Rubrik"', 'author: "A"']
    answers = iter(["", "n"])
    assert module.ask_metadata("", "Text", args, src, True, "A")[0] == 'title: "Filnamn"'
    answers = iter(["-", "n"])
    assert module.ask_metadata("", "Text", args, src, True, "A") == ['author: "A"']
