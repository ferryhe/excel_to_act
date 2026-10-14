"""Bundle the repository skill with a matching wheel for local release review."""

from __future__ import annotations

import argparse
from email.parser import BytesParser
import hashlib
import json
from pathlib import Path
import tomllib
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


def build_skill(root: Path, wheel: Path, out: Path) -> Path:
    project = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))
    version = project["project"]["version"]
    source = root / "src"
    expected = {p.relative_to(source).as_posix(): p.read_bytes()
                for p in (source / "excel_to_act").rglob("*.py")}
    for package, patterns in project["tool"]["setuptools"]["package-data"].items():
        for pattern in patterns:
            for path in (source / package.replace(".", "/")).glob(pattern):
                if path.is_file():
                    expected[path.relative_to(source).as_posix()] = path.read_bytes()

    wheel_bytes = wheel.read_bytes()
    with ZipFile(wheel) as archive:
        metadata_paths = [p for p in archive.namelist() if p.endswith(".dist-info/METADATA")]
        if len(metadata_paths) != 1:
            raise ValueError("wheel must contain one distribution metadata record")
        metadata = BytesParser().parsebytes(archive.read(metadata_paths[0]))
        if metadata["Name"] != "excel-to-act" or metadata["Version"] != version:
            raise ValueError("wheel name/version differs from this source release")
        actual = {p: archive.read(p) for p in archive.namelist()
                  if p.startswith("excel_to_act/") and not p.endswith("/")}
        if actual != expected:
            raise ValueError("wheel code/resources differ from current source; rebuild without stale build cache")

    skill = root / ".agents" / "skills" / "excel-to-act"
    skill_path = "skills/excel-to-act"
    payload = {f"{skill_path}/{name}": (skill / name).read_bytes()
               for name in ("SKILL.md", "agents/openai.yaml")}
    payload[f"{skill_path}/runtime/{wheel.name}"] = wheel_bytes
    plugin = {"$schema": "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json",
              "name": "excel-to-act", "version": version,
              "description": project["project"]["description"]}
    license_path = root / "LICENSE"
    if license_path.is_file():
        payload["LICENSE"] = license_path.read_bytes()
        payload[f"{skill_path}/LICENSE"] = license_path.read_bytes()
    payload["plugin.json"] = (json.dumps(plugin, indent=2) + "\n").encode()
    manifest = {"schema_version": "excel-to-act.skill-release.v1", "version": version,
                "status": "release_candidate", "license_recorded": license_path.is_file(),
                "runtime_wheel": f"{skill_path}/runtime/{wheel.name}",
                "files": {name: hashlib.sha256(data).hexdigest()
                          for name, data in sorted(payload.items())}}
    payload["release.json"] = (json.dumps(manifest, indent=2) + "\n").encode()

    out.mkdir(parents=True, exist_ok=True)
    archive_path = out / f"excel-to-act-skill-{version}.zip"
    wheel_target = out / wheel.name
    sums_path = out / "SHA256SUMS"
    if archive_path.exists() or sums_path.exists():
        raise ValueError("release output already exists; use a new output directory")
    if wheel_target.exists():
        if wheel_target.read_bytes() != wheel_bytes:
            raise ValueError("output contains a different wheel; use a new output directory")
    else:
        with wheel_target.open("xb") as target:
            target.write(wheel_bytes)
    with ZipFile(archive_path, "x") as archive:
        for name, data in sorted(payload.items()):
            entry = ZipInfo(f"excel-to-act/{name}", date_time=(1980, 1, 1, 0, 0, 0))
            entry.external_attr = 0o644 << 16
            archive.writestr(entry, data, compress_type=ZIP_DEFLATED)
    artifacts = [wheel_target, archive_path]
    sdist = out / f"excel_to_act-{version}.tar.gz"
    if sdist.is_file():
        artifacts.append(sdist)
    with sums_path.open("x", encoding="utf-8", newline="\n") as sums:
        for artifact in sorted(artifacts):
            sums.write(f"{hashlib.sha256(artifact.read_bytes()).hexdigest()}  {artifact.name}\n")
    return archive_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--wheel", type=Path, required=True)
    parser.add_argument("--out", type=Path, default=Path("dist"))
    args = parser.parse_args()
    print(build_skill(Path(__file__).resolve().parents[1], args.wheel, args.out))
