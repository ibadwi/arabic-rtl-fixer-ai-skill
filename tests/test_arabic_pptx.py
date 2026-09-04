import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("arabic_pptx", ROOT / "scripts" / "arabic_pptx.py")
PPTX = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = PPTX
SPEC.loader.exec_module(PPTX)

CONTENT_TYPES = b"""<?xml version='1.0' encoding='UTF-8'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>
"""
PRESENTATION = f"""<?xml version='1.0' encoding='UTF-8'?>
<p:presentation xmlns:p='{PPTX.P_NS}'/>
""".encode()


def slide_xml(text: str) -> bytes:
    return f"""<?xml version='1.0' encoding='UTF-8'?>
<p:sld xmlns:p='{PPTX.P_NS}' xmlns:a='{PPTX.A_NS}'>
 <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id='2' name='Arabic text'/></p:nvSpPr>
 <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody>
 </p:sp></p:spTree></p:cSld>
</p:sld>""".encode()


def make_pptx(path: Path, text: str) -> None:
    with ZipFile(path, "w", ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", CONTENT_TYPES)
        package.writestr("ppt/presentation.xml", PRESENTATION)
        package.writestr("ppt/slides/slide1.xml", slide_xml(text))
        package.writestr("ppt/media/keep.bin", b"unchanged")


class ArabicPptxTests(unittest.TestCase):
    def test_repair_cleans_controls_and_preserves_assets(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "broken.pptx"
            fixed = Path(directory) / "fixed.pptx"
            make_pptx(source, "يعتمد النظام\u202e على GPS وTCP/IP.")
            PPTX.repair_pptx(source, fixed, "ar-SA", "Arial", False, True, False)
            result = PPTX.audit_pptx(fixed, expect_presentation_rtl=True)
            self.assertEqual(result.counts()["error"], 0)
            with ZipFile(fixed) as package:
                self.assertEqual(package.read("ppt/media/keep.bin"), b"unchanged")

    def test_repair_requires_distinct_output(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "file.pptx"
            make_pptx(source, "مرحبا")
            with self.assertRaises(ValueError):
                PPTX.repair_pptx(source, source, "ar-SA", None, False, False, False)

    def test_repair_is_idempotent(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.pptx"
            fixed = Path(directory) / "fixed.pptx"
            second = Path(directory) / "second.pptx"
            make_pptx(source, "مرحبا 25%")
            PPTX.repair_pptx(source, fixed, "ar-SA", "Arial", False, True, False)
            stats = PPTX.repair_pptx(fixed, second, "ar-SA", "Arial", False, True, False)
            self.assertEqual(stats["xml_parts_changed"], 0)


if __name__ == "__main__":
    unittest.main()
