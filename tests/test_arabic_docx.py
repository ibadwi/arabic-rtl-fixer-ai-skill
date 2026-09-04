import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("arabic_docx", ROOT / "scripts" / "arabic_docx.py")
DOCX = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = DOCX
SPEC.loader.exec_module(DOCX)

CONTENT_TYPES = b"""<?xml version='1.0' encoding='UTF-8'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
  <Default Extension='xml' ContentType='application/xml'/>
</Types>"""


def document_xml(text: str, configured: bool = False) -> bytes:
    props = "<w:pPr><w:bidi/></w:pPr>" if configured else ""
    run_props = "<w:rPr><w:rFonts w:cs='Arial'/><w:lang w:bidi='ar-SA'/></w:rPr>" if configured else ""
    return f"""<?xml version='1.0' encoding='UTF-8'?>
<w:document xmlns:w='{DOCX.W_NS}'><w:body><w:p>{props}<w:r>{run_props}<w:t>{text}</w:t></w:r></w:p></w:body></w:document>""".encode()


def make_docx(path: Path, text: str, configured: bool = False) -> None:
    with ZipFile(path, "w", ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", CONTENT_TYPES)
        package.writestr("word/document.xml", document_xml(text, configured))
        package.writestr("word/media/keep.bin", b"unchanged")


class ArabicDocxTests(unittest.TestCase):
    def test_audit_detects_missing_rtl_and_metadata(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "broken.docx"
            make_docx(source, "ارتفع الأداء بنسبة 25% في Q4 2026.")
            result = DOCX.audit_docx(source)
            codes = {issue.code for issue in result.issues}
            self.assertIn("MISSING_PARAGRAPH_BIDI", codes)
            self.assertIn("MISSING_RUN_RTL", codes)
            self.assertIn("MISSING_ARABIC_LANGUAGE", codes)
            self.assertIn("MISSING_COMPLEX_FONT", codes)
            self.assertEqual(result.mixed_paragraphs, 1)

    def test_repair_preserves_assets_and_is_idempotent(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "broken.docx"
            fixed = Path(directory) / "fixed.docx"
            second = Path(directory) / "second.docx"
            make_docx(source, "مرحبا\u202e 25%", configured=False)
            stats = DOCX.repair_docx(source, fixed, "ar-SA", "Arial", False)
            self.assertGreater(stats["paragraphs_repaired"], 0)
            self.assertEqual(DOCX.audit_docx(fixed).counts()["error"], 0)
            with ZipFile(fixed) as package:
                self.assertEqual(package.read("word/media/keep.bin"), b"unchanged")
            second_stats = DOCX.repair_docx(fixed, second, "ar-SA", "Arial", False)
            self.assertEqual(second_stats["xml_parts_changed"], 0)

    def test_repair_requires_distinct_output(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "file.docx"
            make_docx(source, "مرحبا")
            with self.assertRaises(ValueError):
                DOCX.repair_docx(source, source, "ar-SA", None, False)


if __name__ == "__main__":
    unittest.main()
