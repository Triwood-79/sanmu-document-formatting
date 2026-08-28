from __future__ import annotations

import base64
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from common import active_profile, load_preset  # noqa: E402
from inspect_docx import inspect_document  # noqa: E402
from privacy_scrub import scrub_docx  # noqa: E402
from validate_docx import validate_document  # noqa: E402


PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nWQAAAAASUVORK5CYII="
)


class SkillTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="odf-test-")
        self.root = Path(self.temp.name)
        self.env = os.environ.copy()
        self.env["CODEX_HOME"] = str(self.root / "codex-state")
        self.env["PYTHONIOENCODING"] = "utf-8"

    def tearDown(self) -> None:
        self.temp.cleanup()

    def run_script(self, name: str, *args: str, check: bool = True) -> subprocess.CompletedProcess:
        return subprocess.run(
            [sys.executable, str(SCRIPTS / name), *map(str, args)],
            env=self.env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            check=check,
        )

    def test_preset_exact_values(self) -> None:
        profile = load_preset()
        self.assertEqual(profile["page"]["margins_cm"], {"top": 3.7, "bottom": 3.5, "left": 2.8, "right": 2.6})
        self.assertEqual(profile["page"]["print_mode"], "duplex")
        self.assertEqual(profile["styles"]["main_title"]["size_pt"], 22)
        self.assertEqual(profile["styles"]["reference_note"]["size_pt"], 15)
        self.assertEqual(profile["page_number"]["size_pt"], 14)
        self.assertTrue(profile["global"]["bold"])
        self.assertNotIn("colophon", profile["styles"])

    def test_profile_draft_confirmation_and_restore(self) -> None:
        self.run_script("profile_manager.py", "begin", "--fresh")
        self.run_script("profile_manager.py", "set", "page.margins_cm.top", "4.0")
        diff = json.loads(self.run_script("profile_manager.py", "diff").stdout)
        self.assertEqual(diff["changes"][0]["current"], 3.7)
        self.assertEqual(diff["changes"][0]["proposed"], 4.0)
        before = json.loads(self.run_script("profile_manager.py", "view").stdout)
        self.assertEqual(before["profile"]["page"]["margins_cm"]["top"], 3.7)
        self.run_script("profile_manager.py", "confirm")
        after = json.loads(self.run_script("profile_manager.py", "view").stdout)
        self.assertEqual(after["profile"]["page"]["margins_cm"]["top"], 4.0)
        self.run_script("profile_manager.py", "restore")
        restored = json.loads(self.run_script("profile_manager.py", "view").stdout)
        self.assertEqual(restored["source"], "built_in_preset")

    def test_invalid_and_fixed_profile_fields_are_rejected(self) -> None:
        self.run_script("profile_manager.py", "begin", "--fresh")
        invalid_alignment = self.run_script(
            "profile_manager.py", "set", "styles.body.alignment", '"diagonal"', check=False
        )
        self.assertNotEqual(invalid_alignment.returncode, 0)
        self.assertIn("Unsupported value", invalid_alignment.stderr)
        fixed_field = self.run_script(
            "profile_manager.py", "set", "page_number.decoration", '"none"', check=False
        )
        self.assertNotEqual(fixed_field.returncode, 0)
        self.assertIn("fixed in V1", fixed_field.stderr)

    def test_page_number_bold_can_be_disabled_and_validated(self) -> None:
        self.run_script("profile_manager.py", "begin", "--fresh")
        self.run_script("profile_manager.py", "set", "page_number.bold", "false")
        self.run_script("profile_manager.py", "confirm")
        output = self.root / "page-number-not-bold.docx"
        self.run_script("docx_engine.py", "create", "--text", "# 示例标题\n正文。", "--output", output)
        profile = load_preset()
        profile["page_number"]["bold"] = False
        report = validate_document(output, profile)
        self.assertTrue(report["valid"], report["errors"])

    def test_create_and_validate_duplex_document(self) -> None:
        output = self.root / "generated.docx"
        text = "# 示例标题\n[说明]某单位\n[说明]某年某月某日\n## 一、一级标题\n正文 Test 123。\n（补充参考资料。）"
        result = json.loads(self.run_script("docx_engine.py", "create", "--text", text, "--output", output).stdout)
        self.assertEqual(result["validation"], "structural_pass")
        report = validate_document(output, load_preset())
        self.assertTrue(report["valid"], report["errors"])
        with zipfile.ZipFile(output) as archive:
            settings = archive.read("word/settings.xml")
            footers = [archive.read(name) for name in archive.namelist() if name.startswith("word/footer")]
        self.assertIn(b"evenAndOddHeaders", settings)
        self.assertTrue(any(b"PAGE" in footer for footer in footers))
        self.assertTrue(any(b'w:jc w:val="right"' in footer for footer in footers))
        self.assertTrue(any(b'w:jc w:val="left"' in footer for footer in footers))

    def test_existing_output_path_gets_numeric_suffix(self) -> None:
        existing = self.root / "named-output.docx"
        original = Document()
        original.add_paragraph("SENTINEL")
        original.save(existing)
        result = json.loads(
            self.run_script("docx_engine.py", "create", "--text", "# 新标题\n新正文。", "--output", existing).stdout
        )
        actual = Path(result["output"])
        self.assertEqual(actual.name, "named-output_2.docx")
        self.assertEqual(Document(existing).paragraphs[0].text, "SENTINEL")
        self.assertTrue(actual.exists())

    def test_new_colophon_content_is_preserved_without_formatting(self) -> None:
        output = self.root / "colophon-unsupported.docx"
        result = json.loads(
            self.run_script(
                "docx_engine.py",
                "create",
                "--text",
                "# 示例标题\n正文。\n[版记]某单位办公室 某年某月某日印发",
                "--output",
                output,
            ).stdout
        )
        self.assertTrue(any("preserved without managed formatting" in warning for warning in result["warnings"]))
        document = Document(output)
        self.assertEqual(document.paragraphs[-1].text, "某单位办公室 某年某月某日印发")
        self.assertEqual(document.paragraphs[-1].style.name, "Normal")
        report = validate_document(output, load_preset())
        self.assertTrue(report["valid"], report["errors"])
        self.assertTrue(any("preserved without managed formatting" in warning for warning in report["warnings"]))

    def test_existing_document_preserves_table_and_picture(self) -> None:
        source = self.root / "source.docx"
        image = self.root / "pixel.png"
        image.write_bytes(PNG_1X1)
        document = Document()
        document.add_paragraph("示例标题")
        document.add_paragraph("这是正文。")
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "表格内容"
        document.add_picture(str(image))
        document.save(source)
        inspection = inspect_document(source)
        self.assertEqual(inspection["summary"]["tables"], 1)
        self.assertGreaterEqual(inspection["summary"]["drawings"], 1)
        mapping = self.root / "classification.json"
        mapping.write_text(json.dumps({"classifications": inspection["classifications"]}, ensure_ascii=False), encoding="utf-8")
        output = self.root / "formatted.docx"
        self.run_script("docx_engine.py", "format-existing", "--input", source, "--output", output, "--classification", mapping, "--confirmed")
        formatted = Document(output)
        self.assertEqual(formatted.tables[0].cell(0, 0).text, "表格内容")
        with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
            before_media = {name: before.read(name) for name in before.namelist() if name.startswith("word/media/")}
            after_media = {name: after.read(name) for name in after.namelist() if name.startswith("word/media/")}
        self.assertEqual(before_media, after_media)

    def test_legacy_extension_and_tracked_changes_are_blocked(self) -> None:
        document = Document()
        document.add_paragraph("示例标题")
        source = self.root / "source.docx"
        document.save(source)
        legacy = self.root / "source.docm"
        legacy.write_bytes(source.read_bytes())
        self.assertTrue(inspect_document(legacy)["blockers"])
        tracked = self.root / "tracked.docx"
        with zipfile.ZipFile(source) as before, zipfile.ZipFile(tracked, "w") as after:
            for info in before.infolist():
                data = before.read(info.filename)
                if info.filename == "word/document.xml":
                    data = data.replace(b"<w:body>", b'<w:body><w:ins w:id="1"/>', 1)
                after.writestr(info, data)
        self.assertTrue(any("Tracked changes" in item for item in inspect_document(tracked)["blockers"]))

    def test_privacy_scrub_removes_authorship(self) -> None:
        output = self.root / "metadata.docx"
        document = Document()
        document.add_paragraph("示例标题")
        document.core_properties.author = "Sample Author"
        document.core_properties.last_modified_by = "Sample Editor"
        document.save(output)
        scrub_docx(output)
        with zipfile.ZipFile(output) as archive:
            core = archive.read("docProps/core.xml")
        self.assertNotIn(b"Sample Author", core)
        self.assertNotIn(b"Sample Editor", core)
        self.assertNotIn(b"lastModifiedBy", core)


if __name__ == "__main__":
    unittest.main(verbosity=2)
