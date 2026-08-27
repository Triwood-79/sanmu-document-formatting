from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

from common import active_profile, read_json
from inspect_docx import classify_paragraphs


STYLE_ROLES = {
    "ODF Main Title": "main_title",
    "ODF Heading 1": "heading1",
    "ODF Heading 2": "heading2",
    "ODF Body": "body",
    "ODF Reference Note": "reference_note",
    "ODF Description": "description",
    "ODF Colophon": "colophon",
}
FONT_ALIASES = {
    "黑体": {"SimHei"},
    "楷体_GB2312": {"楷体GB2312", "KaiTi_GB2312", "KaiTi"},
    "仿宋_GB2312": {"仿宋GB2312", "FangSong_GB2312", "FangSong"},
    "SimSun": {"宋体", "simsun"},
}


def near(actual: float | None, expected: float, tolerance: float = 0.05) -> bool:
    return actual is not None and abs(actual - expected) <= tolerance


def validate_paragraph(paragraph, index: int, role: str, spec: dict, global_bold: bool, errors: list[str]) -> None:
    p_pr = paragraph._p.pPr
    spacing = p_pr.spacing if p_pr is not None else None
    line = int(spacing.get(qn("w:line"))) / 20 if spacing is not None and spacing.get(qn("w:line")) else None
    line_rule = spacing.get(qn("w:lineRule")) if spacing is not None else None
    if line_rule != "exact" or not near(line, spec["line_spacing_pt"], 0.01):
        errors.append(f"Paragraph {index} ({role}) has incorrect exact line spacing")
    ind = p_pr.ind if p_pr is not None else None
    chars = int(ind.get(qn("w:firstLineChars"))) / 100 if ind is not None and ind.get(qn("w:firstLineChars")) else 0
    if not near(chars, spec["first_line_chars"], 0.01):
        errors.append(f"Paragraph {index} ({role}) has incorrect first-line character indent")
    text_runs = [run for run in paragraph.runs if run.text]
    for run in text_runs:
        size = run.font.size.pt if run.font.size else None
        if not near(size, spec["size_pt"], 0.01):
            errors.append(f"Paragraph {index} ({role}) has incorrect font size")
            break
        if global_bold and run.font.bold is not True:
            errors.append(f"Paragraph {index} ({role}) is not fully bold")
            break
        r_pr = run._element.rPr
        east_asia = r_pr.rFonts.get(qn("w:eastAsia")) if r_pr is not None and r_pr.rFonts is not None else None
        allowed_fonts = {spec["font_cn"], spec["font_fallback"]} | FONT_ALIASES.get(spec["font_fallback"], set())
        if east_asia not in allowed_fonts:
            errors.append(f"Paragraph {index} ({role}) has an unexpected East Asian font")
            break


def validate_document(path: Path, profile: dict | None = None) -> dict:
    path = path.expanduser().resolve()
    errors: list[str] = []
    warnings: list[str] = []
    if path.suffix.lower() != ".docx" or not zipfile.is_zipfile(path):
        return {"valid": False, "errors": ["Not a valid .docx package"], "warnings": []}
    profile = profile or active_profile()
    try:
        document = Document(path)
    except Exception as exc:
        return {"valid": False, "errors": [f"DOCX open failed: {type(exc).__name__}"], "warnings": []}

    margins = profile["page"]["margins_cm"]
    for number, section in enumerate(document.sections, start=1):
        if section.orientation == WD_ORIENT.LANDSCAPE:
            warnings.append(f"Landscape section {number} was preserved")
            continue
        values = {
            "top": section.top_margin.cm,
            "bottom": section.bottom_margin.cm,
            "left": section.left_margin.cm,
            "right": section.right_margin.cm,
        }
        for key, expected in margins.items():
            if not near(values[key], expected):
                errors.append(f"Section {number} has incorrect {key} margin")

    inferred = classify_paragraphs(document)
    for item in inferred:
        styled_role = STYLE_ROLES.get(document.paragraphs[item["index"]].style.name)
        if styled_role:
            item["role"] = styled_role
        role = item["role"]
        if role in profile["styles"]:
            validate_paragraph(
                document.paragraphs[item["index"]],
                item["index"],
                role,
                profile["styles"][role],
                profile["global"]["bold"],
                errors,
            )

    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        footer_xml = b"".join(archive.read(name) for name in names if name.startswith("word/footer") and name.endswith(".xml"))
        if b"PAGE" not in footer_xml:
            errors.append("PAGE field is missing from the footer")
        settings = archive.read("word/settings.xml") if "word/settings.xml" in names else b""
        if profile["page"]["print_mode"] == "duplex" and b"evenAndOddHeaders" not in settings:
            errors.append("Duplex mode is missing even/odd footer settings")
        for metadata_name in ("docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"):
            if metadata_name == "docProps/custom.xml" and metadata_name in names:
                errors.append("Custom document properties were not removed")
        core = archive.read("docProps/core.xml") if "docProps/core.xml" in names else b""
        if b"creator" in core or b"lastModifiedBy" in core:
            errors.append("Identifying core properties were not removed")

    return {"valid": not errors, "errors": errors, "warnings": warnings, "mode": "structural"}


def main() -> None:
    parser = argparse.ArgumentParser(description="Structurally validate a formatted DOCX")
    parser.add_argument("input")
    parser.add_argument("--profile", help="Optional complete profile JSON")
    args = parser.parse_args()
    profile = read_json(Path(args.profile).expanduser().resolve()) if args.profile else None
    result = validate_document(Path(args.input), profile)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["valid"] else 1)


if __name__ == "__main__":
    main()
