from __future__ import annotations

import argparse
import hashlib
import hmac
import json
import stat
from pathlib import Path

from common import configure_utf8_stdio


ALLOWED_ROLES = {
    "main_title",
    "heading1",
    "heading2",
    "body",
    "reference_note",
    "description",
    "signature",
    "colophon",
    "skip",
}


def lexical_path(value: str) -> Path:
    return Path(value).expanduser().absolute()


def ensure_no_symlink(path: Path, label: str) -> None:
    current = path
    while current != current.parent:
        if current.is_symlink():
            raise ValueError(f"{label} must not be a symbolic link or be inside one")
        current = current.parent


def validate_structure(data: object) -> None:
    if not isinstance(data, dict):
        raise ValueError("Classification JSON must contain an object")
    classifications = data.get("classifications")
    if not isinstance(classifications, list):
        raise ValueError("Classification JSON must contain a classifications array")
    for position, item in enumerate(classifications):
        if not isinstance(item, dict) or not isinstance(item.get("index"), int):
            raise ValueError(f"Classification item {position} is invalid")
        if item.get("role") not in ALLOWED_ROLES:
            raise ValueError(f"Classification item {position} has an invalid role")
    tables = data.get("tables", [])
    if not isinstance(tables, list):
        raise ValueError("Classification JSON tables must be an array")
    for position, item in enumerate(tables):
        if not isinstance(item, dict) or not isinstance(item.get("index"), int):
            raise ValueError(f"Table item {position} is invalid")
        header_rows = item.get("header_rows")
        if not isinstance(header_rows, int) or header_rows not in {0, 1, 2}:
            raise ValueError(f"Table item {position} has invalid header_rows")


def inspect_candidate(classification_value: str, input_value: str, output_value: str) -> dict:
    classification_lexical = lexical_path(classification_value)
    input_lexical = lexical_path(input_value)
    output_lexical = lexical_path(output_value)
    ensure_no_symlink(classification_lexical, "Classification file")

    classification = classification_lexical.resolve(strict=True)
    source = input_lexical.resolve(strict=True)
    output = output_lexical.resolve(strict=True)
    if classification.suffix.lower() != ".json":
        raise ValueError("Classification file must use the .json extension")
    if not stat.S_ISREG(classification.stat().st_mode):
        raise ValueError("Classification path must be a regular file")
    if classification in {source, output}:
        raise ValueError("Classification file must not be the source or formatted DOCX")
    if not (classification.parent == source.parent == output.parent):
        raise ValueError("Classification, source, and formatted DOCX must be in the same directory")

    raw = classification.read_bytes()
    try:
        data = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("Classification file is not valid UTF-8 JSON") from exc
    validate_structure(data)
    return {
        "eligible": True,
        "classification": str(classification),
        "sha256": hashlib.sha256(raw).hexdigest(),
        "size_bytes": len(raw),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Safely remove a reviewed classification JSON after formatting")
    parser.add_argument("--classification", required=True)
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--confirmed", action="store_true")
    parser.add_argument("--expected-sha256")
    args = parser.parse_args()

    try:
        result = inspect_candidate(args.classification, args.input, args.output)
        if not args.confirmed:
            result.update({"deleted": False, "confirmation_required": True})
        else:
            if not args.expected_sha256:
                raise ValueError("Confirmed deletion requires --expected-sha256 from the eligibility check")
            if not hmac.compare_digest(result["sha256"], args.expected_sha256.lower()):
                raise ValueError("Classification file changed after the eligibility check")
            Path(result["classification"]).unlink()
            result.update({"deleted": True, "confirmation_required": False})
    except (FileNotFoundError, OSError, ValueError) as exc:
        result = {"eligible": False, "deleted": False, "error": str(exc)}

    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result.get("eligible") else 2)


if __name__ == "__main__":
    configure_utf8_stdio()
    main()
