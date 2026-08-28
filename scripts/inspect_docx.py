from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT


H1_RE = re.compile(r"^[一二三四五六七八九十百]+、")
H2_RE = re.compile(r"^（[一二三四五六七八九十百]+）")
DATE_RE = re.compile(r"^(?:\d{4}年\d{1,2}月\d{1,2}日|某年某月某日)$")
NOTE_RE = re.compile(r"^(?:（.*）|\(.*\))$", re.DOTALL)
ALLOWED_EXTENSIONS = {".docx"}
STYLE_ROLES = {
    "ODF Main Title": "main_title",
    "ODF Heading 1": "heading1",
    "ODF Heading 2": "heading2",
    "ODF Body": "body",
    "ODF Reference Note": "reference_note",
    "ODF Description": "description",
    "ODF Colophon": "skip",
}


def package_risks(path: Path) -> tuple[list[str], list[str]]:
    blockers: list[str] = []
    warnings: list[str] = []
    if path.suffix.lower() not in ALLOWED_EXTENSIONS:
        blockers.append("Only .docx files are supported; convert legacy or macro-enabled files first")
        return blockers, warnings
    if not zipfile.is_zipfile(path):
        blockers.append("The file is corrupt, encrypted, or not a valid DOCX package")
        return blockers, warnings
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        if "word/comments.xml" in names or any(name.startswith("word/comments") for name in names):
            blockers.append("Comments are present; provide a clean copy without comments")
        document_xml = archive.read("word/document.xml") if "word/document.xml" in names else b""
        if b"<w:ins" in document_xml or b"<w:del" in document_xml or b"<w:moveFrom" in document_xml:
            blockers.append("Tracked changes are present; accept or reject them before formatting")
        if b"<w:documentProtection" in archive.read("word/settings.xml") if "word/settings.xml" in names else False:
            blockers.append("Document protection is enabled")
        if "word/header1.xml" in names or "word/footer1.xml" in names:
            warnings.append("Existing header or footer content may be replaced when standard page numbers are applied")
    return blockers, warnings


def classify_paragraphs(document: Document) -> list[dict]:
    entries: list[dict] = []
    nonempty = [index for index, paragraph in enumerate(document.paragraphs) if paragraph.text.strip()]
    title_index = nonempty[0] if nonempty else None
    description_indices: set[int] = set()
    if title_index is not None:
        after = [index for index in nonempty if index > title_index][:2]
        if after:
            if DATE_RE.fullmatch(document.paragraphs[after[0]].text.strip()):
                description_indices.add(after[0])
            elif len(after) > 1 and DATE_RE.fullmatch(document.paragraphs[after[1]].text.strip()):
                if len(document.paragraphs[after[0]].text.strip()) <= 40:
                    description_indices.add(after[0])
                description_indices.add(after[1])
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        styled_role = STYLE_ROLES.get(paragraph.style.name)
        if styled_role:
            role = styled_role
        elif not text:
            role = "skip"
        elif index == title_index:
            role = "main_title"
        elif index in description_indices:
            role = "description"
        elif H1_RE.match(text):
            role = "heading1"
        elif H2_RE.match(text):
            role = "heading2"
        elif NOTE_RE.fullmatch(text):
            role = "reference_note"
        else:
            role = "body"
        entries.append({"index": index, "role": role, "text_preview": text[:80]})
    return entries


def inspect_document(path: Path) -> dict:
    blockers, warnings = package_risks(path)
    if blockers and (path.suffix.lower() != ".docx" or not zipfile.is_zipfile(path)):
        return {"input": path.name, "blockers": blockers, "warnings": warnings}
    try:
        document = Document(path)
    except Exception as exc:
        blockers.append(f"DOCX could not be opened: {type(exc).__name__}")
        return {"input": path.name, "blockers": blockers, "warnings": warnings}

    drawing_count = 0
    textbox_count = 0
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml")
        drawing_count = xml.count(b"<w:drawing") + xml.count(b"<w:pict")
        textbox_count = xml.count(b"<w:txbxContent")
    landscape_sections = sum(1 for section in document.sections if section.orientation == WD_ORIENT.LANDSCAPE)
    legacy_colophon_count = sum(1 for paragraph in document.paragraphs if paragraph.style.name == "ODF Colophon")
    if landscape_sections:
        warnings.append("Landscape sections will be preserved and will not receive portrait page settings")
    if legacy_colophon_count:
        warnings.append("Colophon formatting is not supported in V1; detected colophon paragraphs will be preserved unchanged")
    if document.tables or drawing_count or textbox_count:
        warnings.append("Complex elements will be preserved without internal reformatting")

    classifications = classify_paragraphs(document)
    counts: dict[str, int] = {}
    for item in classifications:
        counts[item["role"]] = counts.get(item["role"], 0) + 1
    return {
        "input": path.name,
        "blockers": blockers,
        "warnings": warnings,
        "summary": {
            "top_level_paragraphs": len(document.paragraphs),
            "tables": len(document.tables),
            "drawings": drawing_count,
            "textboxes": textbox_count,
            "landscape_sections": landscape_sections,
            "unsupported_colophon_paragraphs": legacy_colophon_count,
            "classification_counts": counts,
        },
        "classifications": classifications,
        "confirmation_required": not blockers,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect and classify a DOCX before formatting")
    parser.add_argument("input")
    parser.add_argument("--write-map", help="Write the classification JSON to this path")
    args = parser.parse_args()
    result = inspect_document(Path(args.input).expanduser().resolve())
    if args.write_map and "classifications" in result:
        output = Path(args.write_map).expanduser().resolve()
        output.write_text(json.dumps({"classifications": result["classifications"]}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(2 if result.get("blockers") else 0)


if __name__ == "__main__":
    main()
