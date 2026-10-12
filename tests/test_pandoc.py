"""End-to-end tests through Pandoc."""
import re
import zipfile

from conftest import needs_pandoc, run

MANUSCRIPT = """---
title: Test
lang: en-US
---
# Introduction {#sec:intro}

Design is an illusion [@dawkins1986, {p. 13}]. See @fig:eye and @sec:intro.

![The camera eye.](eye.png){#fig:eye}

Evolution {++often++} finds {--the--} same solutions {>>@Anna: source?<<}.
{==Convergence==}{>>Is this the right word?<<} matters <!-- private -->.

\\newpage

Second page.
"""


def _write(workdir, name="manuscript.md", text=MANUSCRIPT):
    (workdir / "eye.png").write_bytes(
        bytes.fromhex("89504e470d0a1a0a0000000d4948445200000001000000010806000000"
                      "1f15c4890000000d49444154789c6360000002000154a24f5d0000000049454e44ae426082"))
    path = workdir / name
    path.write_text(text, encoding="utf-8")
    return path


def _docx(workdir):
    out = next((workdir / "critic-exports").glob("*.docx"))
    with zipfile.ZipFile(out) as z:
        return z.read("word/document.xml").decode(), z.read("word/comments.xml").decode()


@needs_pandoc
def test_word_has_tracked_changes_comments_and_page_break(workdir):
    _write(workdir)
    run(workdir, "manuscript.md", "--bibliography", "library.json")
    doc, comments = _docx(workdir)
    assert re.findall(r'<w:ins [^>]*w:author="Tester"', doc)
    assert re.findall(r'<w:del [^>]*w:author="Tester"', doc)
    assert 'w:author="Anna"' in comments and comments.count("<w:comment ") == 2
    assert 'w:type="page"' in doc
    visible = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", doc))
    assert "Figure 1" in visible and "figure 1" in visible and "section 1" in visible
    assert "Dawkins 1986 , 13" in visible or "Dawkins 1986, 13" in visible


@needs_pandoc
def test_html_markup_and_private_comments(workdir):
    _write(workdir)
    run(workdir, "manuscript.md", "--to", "html", "--bibliography", "library.json")
    html = next((workdir / "critic-exports").glob("*.html")).read_text(encoding="utf-8")
    assert 'class="insertion"' in html and 'class="deletion"' in html
    assert "Anna:" in html and "private" not in html


@needs_pandoc
def test_check_reports_problems(workdir):
    _write(workdir, text=MANUSCRIPT + "\n[@nobody] @fig:none {++open\n")
    result = run(workdir, "manuscript.md", "--check", "--bibliography", "library.json", check=False)
    assert result.returncode == 1
    for expected in ("unknown citation key @nobody", "unknown label @fig:none", "unmatched"):
        assert expected in result.stdout
    assert "@Anna" not in result.stdout


@needs_pandoc
def test_count_leaves_out_comments_and_references(workdir):
    _write(workdir, text="# Title\n\nOne two {++three++} {--gone--} {>>not counted<<} [@x].\n\n"
                         "# References\n\nMany many words here.\n")
    out = run(workdir, "manuscript.md", "--count").stdout
    main = re.search(r"Main text\s+(\d+)", out)
    assert main and main.group(1) == "3"
    assert "1 citation(s) not counted" in out


@needs_pandoc
def test_import_round_trip(workdir):
    source = _write(workdir, text="Evolution finds the same solutions again and again.\n\n"
                                  "Convergence is common in nature.\n")
    reviewed = workdir / "reviewed.md"
    reviewed.write_text(
        'Evolution finds the [same]{.deletion author="Anna"} solutions'
        '[\\ over]{.insertion author="Anna"} again and again.\n\n'
        'Convergence [is common]{.comment-start id="1" author="Anna"}'
        'is common[]{.comment-end id="1"} in nature.\n', encoding="utf-8")
    import subprocess
    subprocess.run(["pandoc", str(reviewed), "-o", str(workdir / "reviewed.docx")], check=True)
    run(workdir, "manuscript.md", "--import", "reviewed.docx", "--in-place")
    text = source.read_text(encoding="utf-8")
    assert "{--same--}{>>@Anna<<}" in text
    assert "{++ over++}{>>@Anna<<}" in text
    assert "{==is common==}{>>@Anna: is common<<}" in text

    # Exported again and re-imported: nothing is added twice.
    run(workdir, "manuscript.md", "--to", "docx", "-o", "again.docx")
    before = source.read_text(encoding="utf-8")
    out = run(workdir, "manuscript.md", "--import", "again.docx", "--in-place").stdout
    assert source.read_text(encoding="utf-8") == before
    assert "already in the manuscript" in out


@needs_pandoc
def test_title_and_author_from_command_line(workdir):
    src = _write(workdir, text="Text.\n")
    run(workdir, str(src), "--to", "html", "--title", "Rubrik", "--doc-author", "Sebastian",
        "--save-metadata", "--no-prompt")
    html = next((workdir / "critic-exports").glob("*.html")).read_text(encoding="utf-8")
    assert "<title>Rubrik</title>" in html and "Sebastian" in html
    assert src.read_text(encoding="utf-8").startswith('---\ntitle: "Rubrik"\nauthor: "Sebastian"\n---\n')


def test_missing_pdf_metadata_is_reported_without_prompt(module, tmp_path, capsys):
    import argparse
    args = argparse.Namespace(title=None, doc_author=None, save_metadata=False, no_prompt=True)
    assert module.ask_metadata("---\ntitle: T\n---\n", "", args, tmp_path / "m.md", True) == []
    assert "no author" in capsys.readouterr().err


def test_default_document_author_fills_in(module, tmp_path, capsys):
    import argparse
    args = argparse.Namespace(title=None, doc_author=None, save_metadata=False, no_prompt=True)
    head = "---\ntitle: T\n---\n"
    assert module.ask_metadata(head, "", args, tmp_path / "m.md", True, "Sebastian") == \
        ['author: "Sebastian"']
    assert module.ask_metadata("---\nauthor: X\ntitle: T\n---\n", "", args, tmp_path / "m.md",
                               True, "Sebastian") == []
    assert capsys.readouterr().err == ""
