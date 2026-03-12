"""Validate Office fixture files are readable and not encrypted."""

import zipfile
from pathlib import Path

import pytest


def _office_fixture_files() -> list[Path]:
    """Return all Office Open XML fixture files under testdata."""
    testdata_root = Path(__file__).resolve().parent / "_templates" / "testdata"
    suffixes = {".docx", ".xlsx", ".pptx"}
    files = [
        path for path in testdata_root.rglob("*")
        if path.is_file() and path.suffix.lower() in suffixes
    ]
    return sorted(files)


@pytest.mark.parametrize("fixture_path", _office_fixture_files())
def test_fixture_file_is_not_encrypted(fixture_path: Path):
    """Each fixture should be a valid OOXML ZIP package without encrypted members."""
    assert zipfile.is_zipfile(fixture_path), (
        f"Fixture is not a valid OOXML ZIP package: {fixture_path}"
    )

    with zipfile.ZipFile(fixture_path, "r") as zf:
        bad_member = zf.testzip()
        assert bad_member is None, f"Corrupt ZIP member '{bad_member}' in {fixture_path}"

        encrypted_members = [
            info.filename for info in zf.infolist()
            if info.flag_bits & 0x1
        ]
        assert not encrypted_members, (
            f"Encrypted ZIP members found in {fixture_path}: {encrypted_members}"
        )


def test_office_fixture_inventory_not_empty():
    """Sanity-check fixture discovery so encryption validation always has coverage."""
    files = _office_fixture_files()
    assert files, "No Office fixture files found under tests/_templates/testdata"
