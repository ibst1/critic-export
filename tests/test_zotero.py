"""Zotero fields in Word: export (Better BibTeX faked) and import."""
import html
import json
import subprocess
import zipfile

from conftest import FIXTURES, needs_pandoc, run

DAWKINS = {"id": "dawkins1986", "type": "book", "title": "The Blind Watchmaker",
           "author": [{"family": "Dawkins", "given": "Richard"}],
           "issued": {"date-parts": [[1986]]}, "publisher": "Longman",
           "custom": {"itemID": 7, "uri": "http://zotero.org/users/1/items/AAAA1111"}}


@needs_pandoc
def test_citations_become_zotero_fields(module, tmp_path, monkeypatch):
    monkeypatch.setenv("CRITIC_EXPORT_ZOTERO_ITEMS", "")          # restored afterwards
    monkeypatch.setattr(module, "bbt_call",
                        lambda method, *params: {"items": {"dawkins1986": DAWKINS}, "errors": {}})
    text = "See [@dawkins1986, 13], [-@dawkins1986] and [@gould1989].\n"
    lua = tmp_path / "zotero.lua"
    lua.write_text(module.ZOTERO_FILTER, encoding="utf-8")
    args = module.zotero_field_setup(text, "---\nlang: sv-SE\n---\n", ["--citeproc"], tmp_path, lua)
    src = tmp_path / "m.md"
    src.write_text(text, encoding="utf-8")
    out = tmp_path / "m.docx"
    subprocess.run(["pandoc", str(src), "--citeproc", "--bibliography",
                    str(FIXTURES / "library.json"), *args, "-o", str(out)], check=True)
    with zipfile.ZipFile(out) as z:
        doc, props = z.read("word/document.xml").decode(), z.read("docProps/custom.xml").decode()
    codes = [json.loads(html.unescape(c)) for c in
             __import__("re").findall(r"CSL_CITATION (\{.*?\}) </w:instrText>", doc)]
    assert len(codes) == 2                                   # gould1989 is not in Zotero
    first, second = codes[0]["citationItems"][0], codes[1]["citationItems"][0]
    assert first["uris"] == ["http://zotero.org/users/1/items/AAAA1111"] and first["id"] == 7
    assert (first["locator"], first["label"]) == ("13", "page")
    assert second["suppress-author"] is True
    assert "ZOTERO_BIBL" in doc and "Gould" in doc
    assert "ZOTERO_PREF_1" in props and "sv-SE" in html.unescape(props)
    assert "zotero-items" not in props


def test_zotero_markdown_and_matching(module, tmp_path):
    fields = {"42": {"itemData": {"id": 42, "title": "Wonderful life", "issued": {"date-parts": [[1989]]}},
                     "uris": []},
              "43": {"itemData": {"id": 43, "title": "Unknown"}, "uris": []}}
    keys = module.resolve_zotero_ids(fields, [FIXTURES / "library.json"])
    assert keys == {"42": "gould1989"}
    cite = {"citationItems": [{"id": 42, "itemData": {"id": 42}, "locator": "3", "label": "chapter",
                               "suppress-author": True}]}
    assert module.zotero_markdown(cite, keys) == "[-@gould1989, chap. 3]"
    assert module.zotero_markdown({"citationItems": [{"id": 43, "itemData": {"id": 43}}]}, keys) is None


@needs_pandoc
def test_inserted_zotero_citation_is_imported_as_key(workdir):
    (workdir / "settings.toml").write_text(
        '[general]\nauthor = "Tester"\n[citations]\nbibliography = "library.json"\n', encoding="utf-8")
    source = workdir / "manuscript.md"
    source.write_text("Evolution finds the same solutions [@dawkins1986].\n\n"
                      "Convergence is common.\n", encoding="utf-8")
    reviewed = workdir / "reviewed.md"
    reviewed.write_text("Evolution finds the same solutions (Dawkins 1986).[\\ See XYZ.]"
                        '{.insertion author="Anna"}\n\nConvergence is common.\n', encoding="utf-8")
    docx = workdir / "reviewed.docx"
    subprocess.run(["pandoc", str(reviewed), "-o", str(docx)], check=True)

    # The reviewer's Zotero (no Better BibTeX, other library): no citation key.
    code = {"citationID": "x", "citationItems": [{
        "id": 42, "uris": ["http://zotero.org/users/999/items/ZZZZ9999"], "locator": "5", "label": "page",
        "itemData": {"id": 42, "type": "book", "title": "Wonderful Life",
                     "issued": {"date-parts": [[1989]]}}}]}
    field = ('</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">'
             ' ADDIN ZOTERO_ITEM CSL_CITATION ' + html.escape(json.dumps(code), quote=False) +
             ' </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
             '<w:r><w:t>(Gould 1989, 5)</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
             '<w:r><w:t xml:space="preserve">')
    with zipfile.ZipFile(docx) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts["word/document.xml"] = parts["word/document.xml"].decode().replace("XYZ", field, 1).encode()
    with zipfile.ZipFile(docx, "w") as z:
        for n, data in parts.items():
            z.writestr(n, data)

    run(workdir, "manuscript.md", "--import", "reviewed.docx", "--in-place")
    text = source.read_text(encoding="utf-8")
    assert "[@dawkins1986].{++ See [@gould1989, 5].++}{>>@Anna<<}" in text
