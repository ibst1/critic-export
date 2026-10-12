"""Settings files, project settings and bibliographies, without Pandoc."""
import tomllib


def test_update_settings_adds_missing_and_keeps_own(module, tmp_path, monkeypatch):
    example = tmp_path / "settings.example.toml"
    example.write_text('[general]\n# Your name.\nauthor = ""\n\n# Folder.\noutput_dir = "x"\n\n'
                       '[critic]\ndocx = "track"\n', encoding="utf-8")
    monkeypatch.setattr(module, "EXAMPLE_SETTINGS", example)
    mine = tmp_path / "settings.toml"
    mine.write_text('[general]\n# mine\nauthor = "Me"\n', encoding="utf-8")
    module.update_settings(mine)
    data = tomllib.loads(mine.read_text(encoding="utf-8"))
    assert data["general"] == {"author": "Me", "output_dir": "x"}
    assert data["critic"] == {"docx": "track"}
    assert "# mine" in mine.read_text(encoding="utf-8")
    assert (tmp_path / "settings.toml.bak").exists()
    assert module.missing_settings(mine.read_text(encoding="utf-8"),
                                   example.read_text(encoding="utf-8")) == []


def test_project_settings_override_and_resolve_paths(module, tmp_path):
    settings = module.load_settings(None)
    project = tmp_path / "critic-export.toml"
    project.write_text('[general]\nauthor = "P"\n[citations]\nbibliography = "lib.json"\n'
                       'style = "apa"\n', encoding="utf-8")
    sub = tmp_path / "a" / "b"
    sub.mkdir(parents=True)
    assert module.find_project_settings(sub) == project
    module.merge_project_settings(settings, project)
    assert settings["general"]["author"] == "P"
    assert settings["citations"]["bibliography"] == str(tmp_path / "lib.json")
    assert settings["citations"]["style"] == "apa"            # a Zotero style name stays


def test_bibliography_globs_and_duplicates(module, workdir, capsys):
    paths = module.resolve_bibliographies(["*.json"], workdir)
    assert [p.name for p in paths] == ["group.json", "library.json"]
    module.warn_duplicate_keys(paths)
    assert "gould1989" in capsys.readouterr().err


def test_comment_author(module):
    assert module.comment_author("@Anna Lund: text", []) == ("Anna Lund", "text")
    assert module.comment_author("Note: text", []) == (None, "Note: text")
    assert module.comment_author("erik: text", ["Erik"]) == ("erik", "text")
