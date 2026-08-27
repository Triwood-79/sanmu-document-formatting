from __future__ import annotations

import argparse
import json
import os
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def clean_xml(name: str, data: bytes) -> bytes | None:
    if name == "docProps/custom.xml":
        return None
    try:
        root = etree.fromstring(data)
    except etree.XMLSyntaxError:
        return data

    if name == "docProps/core.xml":
        for element in list(root):
            if localname(element.tag) in {"creator", "lastModifiedBy", "revision"}:
                root.remove(element)
    elif name == "docProps/app.xml":
        for element in list(root):
            if localname(element.tag) in {"Manager", "Company", "HyperlinkBase", "Template"}:
                root.remove(element)
    elif name == "_rels/.rels":
        for element in list(root):
            if element.get("Type", "").endswith("/custom-properties"):
                root.remove(element)
    elif name == "[Content_Types].xml":
        for element in list(root):
            if element.get("PartName") == "/docProps/custom.xml":
                root.remove(element)
    elif name == "word/_rels/settings.xml.rels":
        for element in list(root):
            if element.get("Type", "").endswith("/attachedTemplate"):
                root.remove(element)
    elif name == "word/settings.xml":
        for element in list(root):
            if localname(element.tag) in {"attachedTemplate", "rsids"}:
                root.remove(element)

    if name.startswith("word/") and name.endswith(".xml"):
        for element in root.iter():
            for attribute in list(element.attrib):
                if localname(attribute).startswith("rsid"):
                    del element.attrib[attribute]
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def scrub_docx(path: Path) -> dict:
    path = path.expanduser().resolve()
    if path.suffix.lower() != ".docx" or not zipfile.is_zipfile(path):
        raise ValueError("Privacy scrub requires a valid .docx package")
    fd, temp_name = tempfile.mkstemp(prefix=f".{path.stem}.", suffix=".docx", dir=path.parent)
    os.close(fd)
    removed: list[str] = []
    try:
        with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temp_name, "w") as target:
            target.comment = source.comment
            for info in source.infolist():
                data = source.read(info.filename)
                cleaned = clean_xml(info.filename, data)
                if cleaned is None:
                    removed.append(info.filename)
                    continue
                target.writestr(info, cleaned)
        os.replace(temp_name, path)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)
    return {"file": path.name, "removed_parts": removed, "status": "scrubbed"}


def main() -> None:
    parser = argparse.ArgumentParser(description="Remove identifying DOCX metadata without changing body text")
    parser.add_argument("input")
    args = parser.parse_args()
    print(json.dumps(scrub_docx(Path(args.input)), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
