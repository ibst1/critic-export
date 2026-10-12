"""CriticMarkup conversion, without Pandoc."""


def test_accept_mode(module):
    text = "a {++new++} b {--old--} c {~~x~>y~~} d {>>note<<} e {==hl==}{>>n<<}"
    assert module.convert(text, "accept", "docx", "", "") == "a new b  c y d  e hl"


def test_track_mode_spans(module):
    out = module.convert("a {++new++} b", "track", "docx", "Me", "D")
    assert out == 'a [new]{.insertion author="Me" date="D"} b'


def test_edge_spaces_survive_pandoc(module):
    # Pandoc drops spaces at the edge of a span; they are kept as "\ ".
    out = module.convert("som{++ länge++} x {++stor ++}hund", "track", "docx", "A", "D")
    assert "[\\ länge]{.insertion" in out
    assert "[stor\\ ]{.insertion" in out


def test_change_spanning_paragraphs_is_split(module):
    out = module.convert("{++one\n\ntwo++}", "track", "docx", "A", "D")
    assert out.count(".insertion") == 2


def test_comment_with_highlight(module):
    out = module.convert("{==word==}{>>check<<}", "track", "docx", "A", "D")
    assert out.startswith('[check]{.comment-start id="0"')
    assert out.endswith('word[]{.comment-end id="0"}')


def test_reviewer_names(module):
    out = module.convert("{>>@Anna: fine<<} {++x++}{>>@Erik Berg<<}", "track", "docx", "Me", "D")
    assert 'author="Anna"' in out and 'critic-name="Anna"' in out
    assert '[x]{.insertion author="Erik Berg"' in out


def test_listed_reviewer_without_at(module):
    out = module.convert("{>>Erik: hm<<} {>>Source: Wiki<<}", "track", "docx", "Me", "D",
                         {"reviewers": ["Erik"]})
    assert 'author="Erik"' in out
    assert '[Source: Wiki]{.comment-start' in out          # not a name


def test_lone_author_tag_is_dropped(module):
    assert module.convert("x {>>@Anna<<} y", "track", "docx", "Me", "D") == "x  y"


def test_links_in_word_comments_become_text(module):
    out = module.convert("{>>see [site](https://x.org)<<}", "track", "docx", "", "D")
    assert "site (https://x.org)" in out
    out = module.convert("{>>see [site](https://x.org)<<}", "markup", "html", "", "D")
    assert "[site](https://x.org)" in out


def test_multi_paragraph_comment_is_flattened(module):
    out = module.convert("{>>one\n\n> two<<}", "track", "docx", "", "D")
    assert "one\\\ntwo" in out
