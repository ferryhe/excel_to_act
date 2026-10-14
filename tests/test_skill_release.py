"""Release bundles contain the matching runtime and no local conversion data."""

import hashlib
import json
from pathlib import Path
import runpy
from zipfile import ZipFile

import pytest


build_skill = runpy.run_path(str(Path(__file__).resolve().parents[1] / "tools/build_skill.py"))["build_skill"]


def _source_and_wheel(tmp_path: Path, fault: str = "") -> tuple[Path, Path]:
    root = tmp_path / "source"
    root.mkdir()
    (root / "pyproject.toml").write_text(
        '[project]\nname = "excel-to-act"\nversion = "0.1.0"\ndescription = "Conversion toolkit"\n'
        '[tool.setuptools.package-data]\n"excel_to_act" = ["agent.md"]\n', encoding="utf-8")
    package = root / "src/excel_to_act"
    package.mkdir(parents=True)
    (package / "__init__.py").write_bytes(b'__version__ = "0.1.0"\n')
    (package / "agent.md").write_bytes(b"Current stage contract.\n")
    skill = root / ".agents/skills/excel-to-act"
    (skill / "agents").mkdir(parents=True)
    (skill / "SKILL.md").write_bytes(b"---\nname: excel-to-act\ndescription: Convert Excel models.\n---\n")
    (skill / "agents/openai.yaml").write_bytes(b'interface:\n  display_name: "Excel to Act"\n')
    (skill / "unrelated.local").write_bytes(b"Do not distribute local files.")
    (root / "input").mkdir()
    (root / "input/client.xlsx").write_bytes(b"Private source data.")
    wheel = tmp_path / "excel_to_act-0.1.0-py3-none-any.whl"
    with ZipFile(wheel, "w") as archive:
        version = "9.0.0" if fault == "version" else "0.1.0"
        archive.writestr("excel_to_act-0.1.0.dist-info/METADATA",
                         f"Name: excel-to-act\nVersion: {version}\n")
        for name in ("__init__.py", "agent.md"):
            if fault == "missing_resource" and name == "agent.md":
                continue
            data = b"old code" if fault == "changed_code" and name == "__init__.py" else (package / name).read_bytes()
            archive.writestr(f"excel_to_act/{name}", data)
        if fault == "stale_file":
            archive.writestr("excel_to_act/deleted.py", b"obsolete code")
    return root, wheel


@pytest.mark.parametrize("with_license", [False, True])
def test_bundle_runtime_hashes_and_distribution_scope(tmp_path: Path, with_license: bool) -> None:
    root, wheel = _source_and_wheel(tmp_path)
    if with_license:
        (root / "LICENSE").write_bytes(b"Test distribution terms.\n")
    out = tmp_path / "dist"
    result = build_skill(root, wheel, out)
    with ZipFile(result) as archive:
        files = {name.removeprefix("excel-to-act/"): archive.read(name) for name in archive.namelist()}
    manifest = json.loads(files.pop("release.json"))
    assert manifest["status"] == "release_candidate"
    assert manifest["license_recorded"] is with_license
    assert manifest["files"] == {name: hashlib.sha256(data).hexdigest() for name, data in files.items()}
    assert files[manifest["runtime_wheel"]] == wheel.read_bytes()
    plugin = json.loads(files["plugin.json"])
    assert plugin["name"] == "excel-to-act" and plugin["version"] == "0.1.0"
    assert "skills/excel-to-act/SKILL.md" in files
    assert all("client.xlsx" not in name and "unrelated.local" not in name for name in files)
    if with_license:
        assert files["LICENSE"] == files["skills/excel-to-act/LICENSE"]
    for line in (out / "SHA256SUMS").read_text().splitlines():
        digest, name = line.split("  ")
        assert hashlib.sha256((out / name).read_bytes()).hexdigest() == digest
    repeated = build_skill(root, wheel, tmp_path / "second-dist")
    assert repeated.read_bytes() == result.read_bytes()
    before = result.read_bytes()
    with pytest.raises(ValueError, match="output already exists"):
        build_skill(root, wheel, out)
    assert result.read_bytes() == before


@pytest.mark.parametrize("fault", ["stale_file", "changed_code", "missing_resource", "version"])
def test_rejects_wheel_from_another_source_without_publishing(tmp_path: Path, fault: str) -> None:
    root, wheel = _source_and_wheel(tmp_path, fault)
    out = tmp_path / "dist"
    with pytest.raises(ValueError, match="wheel"):
        build_skill(root, wheel, out)
    assert not out.exists()
