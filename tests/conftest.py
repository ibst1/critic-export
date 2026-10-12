"""Shared test helpers. Tests that need Pandoc are skipped when it is missing."""
import importlib.util
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "app" / "critic_export.py"
FIXTURES = Path(__file__).resolve().parent / "fixtures"

_spec = importlib.util.spec_from_file_location("critic_export", SCRIPT)
ce = importlib.util.module_from_spec(_spec)
sys.modules["critic_export"] = ce
_spec.loader.exec_module(ce)

needs_pandoc = pytest.mark.skipif(shutil.which("pandoc") is None, reason="Pandoc not installed")


@pytest.fixture
def module():
    return ce


@pytest.fixture
def workdir(tmp_path):
    """A folder with an empty settings file, so tests never read or create
    the settings.toml next to the script."""
    (tmp_path / "settings.toml").write_text('[general]\nauthor = "Tester"\n', encoding="utf-8")
    for f in FIXTURES.iterdir():
        shutil.copy(f, tmp_path / f.name)
    return tmp_path


def run(workdir: Path, *args: str, check: bool = True) -> subprocess.CompletedProcess:
    """Run critic-export as a user would, with the test settings."""
    cmd = [sys.executable, str(SCRIPT), "--settings", str(workdir / "settings.toml"), *args]
    result = subprocess.run(cmd, cwd=workdir, capture_output=True, text=True, encoding="utf-8")
    if check and result.returncode:
        raise AssertionError(f"exit {result.returncode}\n{result.stdout}\n{result.stderr}")
    return result
