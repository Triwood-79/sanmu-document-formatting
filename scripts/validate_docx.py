from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree

from common import active_profile, read_json


STYLE_ROLES = {
    "ODF Main Title": "main_title",
    "ODF Heading 1": "heading1",
    "ODF Heading 2": "heading2",
    "ODF Body": "body",
    "ODF Reference Note": "reference_note",
    "ODF Description": "description",
    "ODF Colophon": "skip",
}
ROLE_STYLES = {role: style for style, role in STYLE_ROLES.items() if role != "skip"}
ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
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
    space_before = paragraph.paragraph_format.space_before
    before_pt = space_before.pt if space_before is not None else 0
    if not near(before_pt, spec["space_before_pt"], 0.01):
        errors.append(f"Paragraph {index} ({role}) has incorrect space before")
    if paragraph.alignment != ALIGNMENTS[spec["alignment"]]:
        errors.append(f"Paragraph {index} ({role}) has incorrect alignment")
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
        if run.font.bold is not global_bold:
            errors.append(f"Paragraph {index} ({role}) has incorrect bold setting")
            break
        r_pr = run._element.rPr
        east_asia = r_pr.rFonts.get(qn("w:eastAsia")) if r_pr is not None and r_pr.rFonts is not None else None
        ascii_font = r_pr.rFonts.get(qn("w:ascii")) if r_pr is not None and r_pr.rFonts is not None else None
        hansi_font = r_pr.rFonts.get(qn("w:hAnsi")) if r_pr is not None and r_pr.rFonts is not None else None
        allowed_fonts = {spec["font_cn"], spec["font_fallback"]} | FONT_ALIASES.get(spec["font_fallback"], set())
        if east_asia not in allowed_fonts:
            errors.append(f"Paragraph {index} ({role}) has an unexpected East Asian font")
            break
        if ascii_font != spec["font_latin"] or hansi_font != spec["font_latin"]:
            errors.append(f"Paragraph {index} ({role}) has an unexpected Latin font")
            break


def validate_document(path: Path, profile: dict | None = None, classifications: list[dict] | None = None) -> dict:
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
        if not near(section.page_width.mm, 210, 0.1) or not near(section.page_height.mm, 297, 0.1):
            errors.append(f"Section {number} is not A4 portrait size")

    expected_roles = (
        {item["index"]: item["role"] for item in classifications}
        if classifications is not None
        else {index: STYLE_ROLES.get(paragraph.style.name) for index, paragraph in enumerate(document.paragraphs)}
    )
    for index, paragraph in enumerate(document.paragraphs):
        role = expected_roles.get(index)
        if role in profile["styles"]:
            if classifications is not None and paragraph.style.name != ROLE_STYLES[role]:
                errors.append(f"Paragraph {index} ({role}) is missing its managed style")
            validate_paragraph(
                paragraph,
                index,
                role,
                profile["styles"][role],
                profile["global"]["bold"],
                errors,
            )
        elif paragraph.text.strip():
            warnings.append(f"Paragraph {index} was preserved without managed formatting")

    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        footer_parts = [archive.read(name) for name in names if name.startswith("word/footer") and name.endswith(".xml")]
        footer_xml = b"".join(footer_parts)
        if b"PAGE" not in footer_xml:
            errors.append("PAGE field is missing from the footer")
        page_spec = profile["page_number"]
        footer_alignments: set[str] = set()
        for data in footer_parts:
            root = etree.fromstring(data)
            field = root.find(".//w:fldSimple", namespaces={"w": W_NS})
            if field is None or "PAGE" not in field.get(qn("w:instr"), ""):
                continue
            text = "".join(root.xpath(".//w:t/text()", namespaces={"w": W_NS}))
            if text != "— 1 —":
                errors.append("Page-number decoration is incorrect")
            paragraph = root.find(".//w:p", namespaces={"w": W_NS})
            jc = paragraph.find("./w:pPr/w:jc", namespaces={"w": W_NS}) if paragraph is not None else None
            if jc is not None:
                footer_alignments.add(jc.get(qn("w:val"), ""))
            r_pr = field.find(".//w:rPr", namespaces={"w": W_NS})
            r_fonts = r_pr.find("./w:rFonts", namespaces={"w": W_NS}) if r_pr is not None else None
            east_asia = r_fonts.get(qn("w:eastAsia")) if r_fonts is not None else None
            allowed_page_fonts = {page_spec["font_cn"], page_spec["font_fallback"]} | FONT_ALIASES.get(page_spec["font_fallback"], set())
            if east_asia not in allowed_page_fonts:
                errors.append("Page number has an unexpected font")
            size = r_pr.find("./w:sz", namespaces={"w": W_NS}) if r_pr is not None else None
            actual_half_points = int(size.get(qn("w:val"))) if size is not None and size.get(qn("w:val")) else None
            if actual_half_points != round(page_spec["size_pt"] * 2):
                errors.append("Page number has an incorrect font size")
            bold = r_pr.find("./w:b", namespaces={"w": W_NS}) if r_pr is not None else None
            actual_bold = bold is not None and bold.get(qn("w:val"), "1") not in {"0", "false", "off"}
            if actual_bold is not page_spec["bold"]:
                errors.append("Page number has an incorrect bold setting")
        expected_alignments = {"right", "left"} if profile["page"]["print_mode"] == "duplex" else {"center"}
        if not expected_alignments.issubset(footer_alignments):
            errors.append("Page-number footer alignment is incorrect")
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
