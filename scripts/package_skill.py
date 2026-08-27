from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path

from privacy_scan import scan_path


EXCLUDED_DIRS = {".git", "__pycache__", ".pytest_cache", ".mypy_cache", "dist", "tests"}
EXCLUDED_SUFFIXES = {".docx", ".pdf", ".png", ".log", ".tmp", ".pyc"}


def package(skill_root: Path, output: Path, deny: list[str]) -> dict:
    findings = scan_path(skill_root, deny)
    if findings:
        return {"created": False, "findings": findings}
    output.parent.mkdir(parents=True, exist_ok=True)
    files = []
    for path in skill_root.rglob("*"):
        relative = path.relative_to(skill_root)
        if not path.is_file() or any(part in EXCLUDED_DIRS for part in relative.parts):
            continue
        if path.suffix.lower() in EXCLUDED_SUFFIXES:
            continue
        files.append((path, Path(skill_root.name) / relative))
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for source, archive_name in sorted(files, key=lambda item: item[1].as_posix()):
            info = zipfile.ZipInfo(archive_name.as_posix(), date_time=(2020, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            archive.writestr(info, source.read_bytes())
    archive_findings = scan_path(output, deny)
    if archive_findings:
        output.unlink(missing_ok=True)
        return {"created": False, "findings": archive_findings}
    return {"created": True, "output": str(output), "files": len(files), "findings": []}


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a clean standalone skill ZIP")
    parser.add_argument("--output")
    parser.add_argument("--deny", action="append", default=[])
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    output = Path(args.output).expanduser().resolve() if args.output else root.parent / "dist" / f"{root.name}.zip"
    result = package(root, output, args.deny)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["created"] else 1)


if __name__ == "__main__":
    main()
