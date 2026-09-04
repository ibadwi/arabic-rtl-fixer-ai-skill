#!/usr/bin/env python3
"""Audit and normalize Arabic WordprocessingML text in editable DOCX files."""

from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
import unicodedata
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
ET.register_namespace("w", W_NS)


def qname(local: str) -> str:
    return f"{{{W_NS}}}{local}"


W_P = qname("p")
W_PPR = qname("pPr")
W_R = qname("r")
W_RPR = qname("rPr")
W_T = qname("t")
W_BIDI = qname("bidi")
W_JC = qname("jc")
W_LANG = qname("lang")
W_RFONTS = qname("rFonts")
W_CS = qname("cs")
W_RTL = qname("rtl")
W_VAL = qname("val")
W_BIDI_LANG = qname("bidi")
W_CS_FONT = qname("cs")
W_ASCII_FONT = qname("ascii")
W_HANSI_FONT = qname("hAnsi")

PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]
RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
    "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
    "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
    "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
]

BIDI_CONTROLS = {
    "\u061c": "ARABIC LETTER MARK",
    "\u200e": "LEFT-TO-RIGHT MARK",
    "\u200f": "RIGHT-TO-LEFT MARK",
    "\u202a": "LEFT-TO-RIGHT EMBEDDING",
    "\u202b": "RIGHT-TO-LEFT EMBEDDING",
    "\u202c": "POP DIRECTIONAL FORMATTING",
    "\u202d": "LEFT-TO-RIGHT OVERRIDE",
    "\u202e": "RIGHT-TO-LEFT OVERRIDE",
    "\u2066": "LEFT-TO-RIGHT ISOLATE",
    "\u2067": "RIGHT-TO-LEFT ISOLATE",
    "\u2068": "FIRST STRONG ISOLATE",
    "\u2069": "POP DIRECTIONAL ISOLATE",
}


@dataclass
class Issue:
    severity: str
    code: str
    part: str
    paragraph: int | None
    message: str
    text: str


@dataclass
class AuditResult:
    file: str
    xml_parts: int
    arabic_paragraphs: int
    mixed_paragraphs: int
    issues: list[Issue]

    def counts(self) -> dict[str, int]:
        return {
            severity: sum(issue.severity == severity for issue in self.issues)
            for severity in ("error", "warning", "notice")
        }


def has_arabic_letters(text: str) -> bool:
    return any(
        unicodedata.category(char).startswith(("L", "M"))
        and "ARABIC" in unicodedata.name(char, "")
        for char in text
    )


def has_ltr_or_number(text: str) -> bool:
    return any(unicodedata.bidirectional(char) in {"L", "EN"} for char in text)


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(W_T))


def direct_child(parent: ET.Element | None, tag: str) -> ET.Element | None:
    if parent is None:
        return None
    return next((child for child in parent if child.tag == tag), None)


def enabled(element: ET.Element | None) -> bool:
    if element is None:
        return False
    return element.get(W_VAL, "1").lower() not in {"0", "false", "off"}


def docx_xml_parts(package: ZipFile) -> list[str]:
    return [
        name
        for name in package.namelist()
        if name.startswith("word/") and name.endswith(".xml")
    ]


def inspect_xml(part: str, data: bytes) -> tuple[int, int, list[Issue]]:
    root = ET.fromstring(data)
    arabic_count = 0
    mixed_count = 0
    issues: list[Issue] = []
    for index, paragraph in enumerate(root.iter(W_P), start=1):
        text = paragraph_text(paragraph)
        if not has_arabic_letters(text):
            continue
        arabic_count += 1
        if has_ltr_or_number(text):
            mixed_count += 1
        controls = [BIDI_CONTROLS[c] for c in text if c in BIDI_CONTROLS]
        if controls:
            issues.append(Issue("error", "BIDI_CONTROL", part, index, f"Hidden bidi controls: {', '.join(sorted(set(controls)))}", text))

        ppr = direct_child(paragraph, W_PPR)
        if not enabled(direct_child(ppr, W_BIDI)):
            issues.append(Issue("error", "MISSING_PARAGRAPH_BIDI", part, index, "Arabic paragraph is missing native w:bidi direction", text))

        for run in paragraph.findall(W_R):
            run_text = "".join(node.text or "" for node in run.iter(W_T))
            if not has_arabic_letters(run_text):
                continue
            rpr = direct_child(run, W_RPR)
            if not enabled(direct_child(rpr, W_RTL)):
                issues.append(Issue("warning", "MISSING_RUN_RTL", part, index, "Arabic run is missing native w:rtl", run_text))
            lang = direct_child(rpr, W_LANG)
            if lang is None or not lang.get(W_BIDI_LANG):
                issues.append(Issue("warning", "MISSING_ARABIC_LANGUAGE", part, index, "Arabic run is missing w:lang/@w:bidi", run_text))
            fonts = direct_child(rpr, W_RFONTS)
            if fonts is None or not fonts.get(W_CS_FONT):
                issues.append(Issue("warning", "MISSING_COMPLEX_FONT", part, index, "Arabic run is missing the complex-script font slot", run_text))
    return arabic_count, mixed_count, issues


def audit_docx(path: Path) -> AuditResult:
    with ZipFile(path, "r") as package:
        if "[Content_Types].xml" not in package.namelist() or "word/document.xml" not in package.namelist():
            raise ValueError("Not a DOCX package: required OOXML parts are missing")
        parts = docx_xml_parts(package)
        arabic = mixed = 0
        issues: list[Issue] = []
        for part in parts:
            try:
                part_arabic, part_mixed, part_issues = inspect_xml(part, package.read(part))
            except ET.ParseError as exc:
                issues.append(Issue("error", "INVALID_XML", part, None, str(exc), ""))
                continue
            arabic += part_arabic
            mixed += part_mixed
            issues.extend(part_issues)
    return AuditResult(str(path), len(parts), arabic, mixed, issues)


def ensure_first(parent: ET.Element, tag: str) -> tuple[ET.Element, bool]:
    child = direct_child(parent, tag)
    if child is not None:
        return child, False
    child = ET.Element(tag)
    parent.insert(0, child)
    return child, True


def ensure_child(parent: ET.Element, tag: str) -> tuple[ET.Element, bool]:
    child = direct_child(parent, tag)
    if child is not None:
        return child, False
    child = ET.SubElement(parent, tag)
    return child, True


def ensure_ordered_child(parent: ET.Element, tag: str, order: list[str]) -> tuple[ET.Element, bool]:
    child = direct_child(parent, tag)
    if child is not None:
        return child, False
    child = ET.Element(tag)
    target_name = tag.rsplit("}", 1)[-1]
    target_rank = order.index(target_name)
    for index, existing in enumerate(parent):
        existing_name = existing.tag.rsplit("}", 1)[-1]
        if existing_name in order and order.index(existing_name) > target_rank:
            parent.insert(index, child)
            return child, True
    parent.append(child)
    return child, True


def set_attr(element: ET.Element, name: str, value: str) -> bool:
    if element.get(name) == value:
        return False
    element.set(name, value)
    return True


def repair_xml(data: bytes, language: str, font: str | None) -> tuple[bytes, bool, dict[str, int]]:
    root = ET.fromstring(data)
    changed = False
    stats = {"controls_removed": 0, "paragraphs_repaired": 0, "runs_repaired": 0}
    for paragraph in root.iter(W_P):
        text = paragraph_text(paragraph)
        if not has_arabic_letters(text):
            continue
        paragraph_changed = False
        ppr, created = ensure_first(paragraph, W_PPR)
        paragraph_changed |= created
        bidi, created = ensure_ordered_child(ppr, W_BIDI, PPR_ORDER)
        paragraph_changed |= created
        paragraph_changed |= set_attr(bidi, W_VAL, "1")

        for run in paragraph.findall(W_R):
            content = "".join(node.text or "" for node in run.iter(W_T))
            if not has_arabic_letters(content):
                continue
            run_changed = False
            rpr, created = ensure_first(run, W_RPR)
            run_changed |= created
            rtl, created = ensure_ordered_child(rpr, W_RTL, RPR_ORDER)
            run_changed |= created
            run_changed |= set_attr(rtl, W_VAL, "1")
            lang, created = ensure_ordered_child(rpr, W_LANG, RPR_ORDER)
            run_changed |= created
            run_changed |= set_attr(lang, W_BIDI_LANG, language)
            fonts, created = ensure_ordered_child(rpr, W_RFONTS, RPR_ORDER)
            run_changed |= created
            chosen_font = font or fonts.get(W_CS_FONT) or fonts.get(W_ASCII_FONT) or fonts.get(W_HANSI_FONT)
            if chosen_font:
                run_changed |= set_attr(fonts, W_CS_FONT, chosen_font)
            for node in run.iter(W_T):
                original = node.text or ""
                cleaned = "".join(c for c in original if c not in BIDI_CONTROLS)
                if cleaned != original:
                    node.text = cleaned
                    if cleaned[:1].isspace() or cleaned[-1:].isspace():
                        node.set(f"{{{XML_NS}}}space", "preserve")
                    stats["controls_removed"] += len(original) - len(cleaned)
                    run_changed = True
            if run_changed:
                stats["runs_repaired"] += 1
                paragraph_changed = True
        if paragraph_changed:
            stats["paragraphs_repaired"] += 1
            changed = True
    if not changed:
        return data, False, stats
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), True, stats


def repair_docx(source: Path, output: Path, language: str, font: str | None, force: bool) -> dict[str, int]:
    if source.resolve() == output.resolve():
        raise ValueError("Repair requires a distinct output path; preserve the source DOCX")
    if output.exists() and not force:
        raise FileExistsError(f"Output already exists: {output}. Pass --force to replace it.")
    output.parent.mkdir(parents=True, exist_ok=True)
    totals = {"xml_parts_changed": 0, "controls_removed": 0, "paragraphs_repaired": 0, "runs_repaired": 0}
    handle, temp_name = tempfile.mkstemp(prefix=f".{output.stem}.", suffix=".docx", dir=output.parent)
    os.close(handle)
    temporary = Path(temp_name)
    try:
        with ZipFile(source, "r") as incoming, ZipFile(temporary, "w") as outgoing:
            if "[Content_Types].xml" not in incoming.namelist() or "word/document.xml" not in incoming.namelist():
                raise ValueError("Not a DOCX package: required OOXML parts are missing")
            outgoing.comment = incoming.comment
            for item in incoming.infolist():
                data = incoming.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    try:
                        data, part_changed, stats = repair_xml(data, language, font)
                    except ET.ParseError as exc:
                        raise ValueError(f"Invalid XML in {item.filename}: {exc}") from exc
                    if part_changed:
                        totals["xml_parts_changed"] += 1
                    for key in ("controls_removed", "paragraphs_repaired", "runs_repaired"):
                        totals[key] += stats[key]
                outgoing.writestr(item, data)
        os.replace(temporary, output)
    finally:
        if temporary.exists():
            temporary.unlink()
    return totals


def result_to_json(result: AuditResult) -> str:
    payload = asdict(result)
    payload["counts"] = result.counts()
    return json.dumps(payload, ensure_ascii=False, indent=2)


def print_result(result: AuditResult, strict: bool, max_issues: int = 80) -> int:
    counts = result.counts()
    print(f"Arabic DOCX audit: {result.file}")
    print(f"XML parts: {result.xml_parts} | Arabic paragraphs: {result.arabic_paragraphs} | mixed paragraphs: {result.mixed_paragraphs}")
    visible = result.issues if max_issues == 0 else result.issues[:max_issues]
    for issue in visible:
        location = issue.part + (f":paragraph-{issue.paragraph}" if issue.paragraph else "")
        print(f"{issue.severity.upper()} [{issue.code}] {location}: {issue.message}")
        if issue.text:
            print(f"  text: {issue.text!r}")
    if len(result.issues) > len(visible):
        print(f"... {len(result.issues) - len(visible)} additional issues omitted; use --max-issues 0 or --json to show all.")
    failed = counts["error"] > 0 or (strict and counts["warning"] > 0)
    print(f"{'FAIL' if failed else 'PASS'}: errors={counts['error']} warnings={counts['warning']} notices={counts['notice']} strict={'yes' if strict else 'no'}")
    return 1 if failed else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    audit = subparsers.add_parser("audit", help="Inspect Arabic DOCX structure without modifying the file")
    audit.add_argument("file", type=Path)
    audit.add_argument("--strict", action="store_true", help="Treat warnings as failures")
    audit.add_argument("--json", action="store_true", help="Print machine-readable JSON")
    audit.add_argument("--max-issues", type=int, default=80)
    repair = subparsers.add_parser("repair", help="Write a normalized copy, then run a strict audit")
    repair.add_argument("source", type=Path)
    repair.add_argument("output", type=Path)
    repair.add_argument("--lang", default="ar-SA")
    repair.add_argument("--font")
    repair.add_argument("--force", action="store_true")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        if args.command == "audit":
            result = audit_docx(args.file)
            if args.json:
                print(result_to_json(result))
                counts = result.counts()
                return 1 if counts["error"] or (args.strict and counts["warning"]) else 0
            return print_result(result, args.strict, args.max_issues)
        stats = repair_docx(args.source, args.output, args.lang, args.font, args.force)
        print("Repair summary: " + " | ".join(f"{key}={value}" for key, value in stats.items()))
        return print_result(audit_docx(args.output), strict=True)
    except (BadZipFile, FileNotFoundError, FileExistsError, PermissionError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
