#!/usr/bin/env python3
"""Generate small synthetic DOCX and PPTX examples for this repository."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1]
EXAMPLES = ROOT / "examples"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
CONTENT_TYPES = b"<?xml version='1.0' encoding='UTF-8'?><Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>"


def load(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def write_docx(path: Path) -> None:
    xml = f"""<?xml version='1.0' encoding='UTF-8'?>
<w:document xmlns:w='{W_NS}'><w:body><w:p><w:r><w:t>ارتفع الأداء\u202e بنسبة 25% في Q4 2026.</w:t></w:r></w:p></w:body></w:document>""".encode()
    with ZipFile(path, "w", ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", CONTENT_TYPES)
        package.writestr("word/document.xml", xml)


def write_pptx(path: Path) -> None:
    presentation = f"<?xml version='1.0' encoding='UTF-8'?><p:presentation xmlns:p='{P_NS}'/>".encode()
    slide = f"""<?xml version='1.0' encoding='UTF-8'?>
<p:sld xmlns:p='{P_NS}' xmlns:a='{A_NS}'><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id='2' name='Arabic example'/></p:nvSpPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>يعتمد النظام\u202e على GPS وTCP/IP.</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>""".encode()
    with ZipFile(path, "w", ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", CONTENT_TYPES)
        package.writestr("ppt/presentation.xml", presentation)
        package.writestr("ppt/slides/slide1.xml", slide)


def main() -> None:
    broken = EXAMPLES / "broken"
    expected = EXAMPLES / "expected"
    broken.mkdir(parents=True, exist_ok=True)
    expected.mkdir(parents=True, exist_ok=True)
    docx_source = broken / "arabic-mixed.docx"
    pptx_source = broken / "arabic-mixed.pptx"
    write_docx(docx_source)
    write_pptx(pptx_source)
    docx = load("example_arabic_docx", ROOT / "scripts" / "arabic_docx.py")
    pptx = load("example_arabic_pptx", ROOT / "scripts" / "arabic_pptx.py")
    docx.repair_docx(docx_source, expected / "arabic-mixed-fixed.docx", "ar-SA", "Arial", True)
    pptx.repair_pptx(pptx_source, expected / "arabic-mixed-fixed.pptx", "ar-SA", "Arial", False, True, True)
    print("Generated synthetic broken and repaired OOXML examples.")


if __name__ == "__main__":
    main()
