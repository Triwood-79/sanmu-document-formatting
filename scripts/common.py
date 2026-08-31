from __future__ import annotations

import copy
import json
import os
import sys
import tempfile
from pathlib import Path
from typing import Any


SKILL_ROOT = Path(__file__).resolve().parents[1]
PRESET_PATH = SKILL_ROOT / "assets" / "presets" / "generic_official_v1.json"
STATE_NAME = "official-document-formatting"
ALIGNMENT_VALUES = {"left", "center", "right", "justify"}
PRINT_MODE_VALUES = {"single", "duplex"}
STYLE_FIELDS = {
    "font_cn",
    "font_fallback",
    "font_latin",
    "size_pt",
    "alignment",
    "first_line_chars",
    "line_spacing_pt",
    "space_before_pt",
}
CONFIGURABLE_EXACT_PATHS = {
    "page.print_mode",
    "global.bold",
    "page_number.font_cn",
    "page_number.font_fallback",
    "page_number.size_pt",
    "page_number.bold",
}


def configure_utf8_stdio() -> None:
    """Keep CLI JSON readable when an agent captures output on Windows."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            reconfigure(encoding="utf-8", errors="replace")


def read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        value = json.load(handle)
    if not isinstance(value, dict):
        raise ValueError(f"Expected a JSON object: {path}")
    return value


def write_json_atomic(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(temp_name, path)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)


def state_root() -> Path:
    codex_home = os.environ.get("CODEX_HOME")
    base = Path(codex_home).expanduser() if codex_home else Path.home() / ".codex"
    return base / "state" / STATE_NAME


def load_preset() -> dict[str, Any]:
    return read_json(PRESET_PATH)


def deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def active_record() -> dict[str, Any] | None:
    path = state_root() / "active.json"
    return read_json(path) if path.exists() else None


def active_profile(extra_override: dict[str, Any] | None = None) -> dict[str, Any]:
    preset = load_preset()
    record = active_record()
    if record:
        overrides = record.get("overrides", {})
        validate_override(overrides, preset)
        preset = deep_merge(preset, overrides)
    if extra_override:
        validate_override(extra_override, load_preset())
        preset = deep_merge(preset, extra_override)
    return preset


def validate_override(override: dict[str, Any], schema: dict[str, Any] | None = None, prefix: str = "") -> None:
    schema = schema or load_preset()
    for key, value in override.items():
        dotted = f"{prefix}.{key}" if prefix else key
        if key not in schema:
            raise ValueError(f"Unsupported profile field: {dotted}")
        expected = schema[key]
        if isinstance(value, dict):
            if not isinstance(expected, dict):
                raise ValueError(f"Field is not an object: {dotted}")
            validate_override(value, expected, dotted)
        elif isinstance(expected, bool):
            if not isinstance(value, bool):
                raise ValueError(f"Expected boolean at {dotted}")
        elif isinstance(expected, (int, float)) and not isinstance(expected, bool):
            if not isinstance(value, (int, float)) or isinstance(value, bool):
                raise ValueError(f"Expected number at {dotted}")
        elif not isinstance(value, type(expected)):
            raise ValueError(f"Unexpected value type at {dotted}")
        if not isinstance(value, dict):
            validate_configurable_value(dotted, value)


def validate_configurable_value(dotted: str, value: Any) -> None:
    parts = dotted.split(".")
    is_margin = len(parts) == 3 and parts[:2] == ["page", "margins_cm"]
    is_style = len(parts) == 3 and parts[0] == "styles" and parts[2] in STYLE_FIELDS
    if dotted not in CONFIGURABLE_EXACT_PATHS and not is_margin and not is_style:
        raise ValueError(f"Field is fixed in V1 and cannot be customized: {dotted}")

    if dotted == "page.print_mode" and value not in PRINT_MODE_VALUES:
        raise ValueError(f"Unsupported value at {dotted}; choose single or duplex")
    if dotted.endswith(".alignment") and value not in ALIGNMENT_VALUES:
        raise ValueError(f"Unsupported value at {dotted}; choose left, center, right, or justify")
    if dotted.endswith((".font_cn", ".font_fallback", ".font_latin")):
        if not isinstance(value, str) or not value.strip():
            raise ValueError(f"Font name cannot be empty at {dotted}")
    if is_margin and not 0.5 <= float(value) <= 10:
        raise ValueError(f"Margin must be between 0.5 and 10 cm at {dotted}")
    if dotted.endswith(".size_pt") and not 5 <= float(value) <= 72:
        raise ValueError(f"Font size must be between 5 and 72 pt at {dotted}")
    if dotted.endswith(".first_line_chars") and not 0 <= float(value) <= 20:
        raise ValueError(f"First-line indent must be between 0 and 20 characters at {dotted}")
    if dotted.endswith(".line_spacing_pt") and not 5 <= float(value) <= 100:
        raise ValueError(f"Line spacing must be between 5 and 100 pt at {dotted}")
    if dotted.endswith(".space_before_pt") and not 0 <= float(value) <= 100:
        raise ValueError(f"Space before must be between 0 and 100 pt at {dotted}")


def set_dotted(target: dict[str, Any], dotted: str, value: Any) -> None:
    schema: Any = load_preset()
    cursor: Any = target
    parts = dotted.split(".")
    for part in parts[:-1]:
        if not isinstance(schema, dict) or part not in schema or not isinstance(schema[part], dict):
            raise ValueError(f"Unsupported profile field: {dotted}")
        schema = schema[part]
        cursor = cursor.setdefault(part, {})
    leaf = parts[-1]
    if not isinstance(schema, dict) or leaf not in schema:
        raise ValueError(f"Unsupported profile field: {dotted}")
    expected = schema[leaf]
    validate_override({leaf: value}, {leaf: expected}, ".".join(parts[:-1]))
    cursor[leaf] = value


def flatten(value: dict[str, Any], prefix: str = "") -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, item in value.items():
        dotted = f"{prefix}.{key}" if prefix else key
        if isinstance(item, dict):
            result.update(flatten(item, dotted))
        else:
            result[dotted] = item
    return result


def choose_output_path(input_path: Path, output: str | None = None) -> Path:
    if output:
        candidate = Path(output).expanduser().resolve()
        if candidate.suffix.lower() != ".docx":
            raise ValueError("Output must use the .docx extension")
        if candidate == input_path.resolve():
            raise ValueError("Output path must not overwrite the input file")
        return unique_output_path(candidate)
    base = input_path.with_name(f"{input_path.stem}_排版后.docx")
    return unique_output_path(base)


def unique_output_path(candidate: Path) -> Path:
    if not candidate.exists():
        return candidate
    index = 2
    while True:
        numbered = candidate.with_name(f"{candidate.stem}_{index}{candidate.suffix}")
        if not numbered.exists():
            return numbered
        index += 1
