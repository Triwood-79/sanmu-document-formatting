from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path

from common import configure_utf8_stdio


TEXT_SUFFIXES = {".md", ".txt", ".json", ".yaml", ".yml", ".py", ".ps1", ".toml", ".xml", ".rels"}
IGNORED_PARTS = {".git", "__pycache__", ".pytest_cache", ".mypy_cache", "dist"}
PATH_PATTERNS = {
    "absolute_user_path": re.compile(r"[A-Za-z]:[\\/](?:Users|Documents and Settings)[\\/][^\\/\s]+", re.IGNORECASE),
    "file_uri": re.compile("file:" + "//", re.IGNORECASE),
}


def decode_text(data: bytes) -> str | None:
    for encoding in ("utf-8", "utf-8-sig", "utf-16"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return None


def scan_text(label: str, text: str, deny: list[str]) -> list[dict]:
    findings: list[dict] = []
    for name, pattern in PATH_PATTERNS.items():
        match = pattern.search(text)
        if match:
            findings.append({"file": label, "rule": name, "match": match.group(0)[:80]})
    for value in deny:
        if value and value.casefold() in text.casefold():
            findings.append({"file": label, "rule": "deny_term", "match": value})
    return findings


def scan_archive(path: Path, deny: list[str]) -> list[dict]:
    findings: list[dict] = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if Path(name).suffix.lower() not in TEXT_SUFFIXES and not name.endswith(".rels"):
                continue
            text = decode_text(archive.read(name))
            if text is not None:
                findings.extend(scan_text(f"{path.name}!/{name}", text, deny))
    return findings


def scan_path(path: Path, deny: list[str]) -> list[dict]:
    findings: list[dict] = []
    if path.is_file():
        if path.suffix.lower() in {".zip", ".docx"} and zipfile.is_zipfile(path):
            return scan_archive(path, deny)
        if path.suffix.lower() in TEXT_SUFFIXES:
            text = decode_text(path.read_bytes())
            if text is not None:
                return scan_text(path.name, text, deny)
        return []
    for child in path.rglob("*"):
        if not child.is_file() or any(part in IGNORED_PARTS for part in child.relative_to(path).parts):
            continue
        findings.extend(scan_path(child, deny))
    return findings


def main() -> None:
    parser = argparse.ArgumentParser(description="Scan a skill folder or archive for release privacy risks")
    parser.add_argument("targets", nargs="+")
    parser.add_argument("--deny", action="append", default=[], help="Project-specific term that must not appear")
    args = parser.parse_args()
    findings: list[dict] = []
    for target in args.targets:
        findings.extend(scan_path(Path(target).expanduser().resolve(), args.deny))
    result = {"passed": not findings, "findings": findings}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    configure_utf8_stdio()
    main()
