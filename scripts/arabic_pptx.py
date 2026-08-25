#!/usr/bin/env python3
"""Audit and normalize Arabic DrawingML text in editable PPTX files."""

from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
import tempfile
import unicodedata
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

for prefix, uri in {
    "a": A_NS,
    "p": P_NS,
    "c": C_NS,
    "r": R_NS,
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
}.items():
    ET.register_namespace(prefix, uri)


def qname(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


A_P = qname(A_NS, "p")
A_PPR = qname(A_NS, "pPr")
A_BODYPR = qname(A_NS, "bodyPr")
A_R = qname(A_NS, "r")
A_FLD = qname(A_NS, "fld")
A_RPR = qname(A_NS, "rPr")
A_DEFRPR = qname(A_NS, "defRPr")
A_ENDRPR = qname(A_NS, "endParaRPr")
A_T = qname(A_NS, "t")
A_LATIN = qname(A_NS, "latin")
A_EA = qname(A_NS, "ea")
A_CS = qname(A_NS, "cs")
A_BUCHAR = qname(A_NS, "buChar")

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
BULLET_RE = re.compile(r"^\s*([•◦▪‣⁃])\s+")
TRUE_VALUES = {"1", "true", "on"}
RTL_ALIGNMENTS = {"r", "ctr", "just", "justLow", "dist", "thaiDist"}

PPR_ORDER = [
    "lnSpc", "spcBef", "spcAft", "buClr", "buClrTx", "buSzPct", "buSzPts",
    "buSzTx", "buFont", "buFontTx", "buNone", "buAutoNum", "buChar", "buBlip",
    "tabLst", "defRPr", "extLst",
]
RPR_ORDER = [
    "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
    "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill",
    "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "extLst",
]


@dataclass
class Issue:
    severity: str
    code: str
    part: str
    paragraph: int | None
    shape: str | None
    message: str
    text: str


@dataclass
class AuditResult:
    file: str
    presentation_rtl: bool | None
    xml_parts: int
    arabic_paragraphs: int
    mixed_paragraphs: int
    native_bullets: int
    issues: list[Issue]

    def counts(self) -> dict[str, int]:
        return {
            severity: sum(issue.severity == severity for issue in self.issues)
            for severity in ("error", "warning", "notice")
        }


def local_name(element_or_tag: ET.Element | str) -> str:
    tag = element_or_tag.tag if isinstance(element_or_tag, ET.Element) else element_or_tag
    return tag.rsplit("}", 1)[-1]


def has_arabic_letters(text: str) -> bool:
    for char in text:
        if unicodedata.category(char).startswith(("L", "M")) and "ARABIC" in unicodedata.name(char, ""):
            return True
    return False


def has_ltr_or_number(text: str) -> bool:
    return any(unicodedata.bidirectional(char) in {"L", "EN"} for char in text)


def bidi_controls_in(text: str) -> list[str]:
    return [BIDI_CONTROLS[char] for char in text if char in BIDI_CONTROLS]


def strip_bidi_controls(text: str) -> tuple[str, int]:
    cleaned = "".join(char for char in text if char not in BIDI_CONTROLS)
    return cleaned, len(text) - len(cleaned)


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(A_T))


def direct_child(parent: ET.Element | None, tag: str) -> ET.Element | None:
    if parent is None:
        return None
    return next((child for child in parent if child.tag == tag), None)


def build_parent_map(root: ET.Element) -> dict[ET.Element, ET.Element]:
    return {child: parent for parent in root.iter() for child in parent}


def nearest_container(paragraph: ET.Element, parents: dict[ET.Element, ET.Element]) -> ET.Element | None:
    current = parents.get(paragraph)
    while current is not None:
        if local_name(current) in {"txBody", "txPr", "rich"}:
            return current
        current = parents.get(current)
    return None


def nearest_shape_name(paragraph: ET.Element, parents: dict[ET.Element, ET.Element]) -> str | None:
    current = parents.get(paragraph)
    while current is not None:
        if local_name(current) in {"sp", "graphicFrame", "cxnSp", "pic"}:
            for descendant in current.iter():
                if local_name(descendant) == "cNvPr" and descendant.get("name"):
                    return descendant.get("name")
            return None
        current = parents.get(current)
    return None


def direct_runs(paragraph: ET.Element) -> Iterable[ET.Element]:
    return (child for child in paragraph if child.tag in {A_R, A_FLD})


def run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.iter(A_T))


def insert_ordered(parent: ET.Element, child: ET.Element, order: list[str]) -> None:
    rank = order.index(local_name(child)) if local_name(child) in order else len(order)
    for index, existing in enumerate(parent):
        existing_name = local_name(existing)
        existing_rank = order.index(existing_name) if existing_name in order else len(order)
        if existing_rank > rank:
            parent.insert(index, child)
            return
    parent.append(child)


def ensure_ppr(paragraph: ET.Element) -> tuple[ET.Element, bool]:
    ppr = direct_child(paragraph, A_PPR)
    if ppr is not None:
        return ppr, False
    ppr = ET.Element(A_PPR)
    paragraph.insert(0, ppr)
    return ppr, True


def ensure_rpr(run: ET.Element) -> tuple[ET.Element, bool]:
    rpr = direct_child(run, A_RPR)
    if rpr is not None:
        return rpr, False
    rpr = ET.Element(A_RPR)
    run.insert(0, rpr)
    return rpr, True


def ensure_defrpr(ppr: ET.Element) -> tuple[ET.Element, bool]:
    defrpr = direct_child(ppr, A_DEFRPR)
    if defrpr is not None:
        return defrpr, False
    defrpr = ET.Element(A_DEFRPR)
    insert_ordered(ppr, defrpr, PPR_ORDER)
    return defrpr, True


def set_attr(element: ET.Element, name: str, value: str) -> bool:
    if element.get(name) == value:
        return False
    element.set(name, value)
    return True


def explicit_typeface(properties: ET.Element | None) -> str | None:
    if properties is None:
        return None
    for tag in (A_CS, A_LATIN, A_EA):
        font = direct_child(properties, tag)
        if font is not None and font.get("typeface"):
            return font.get("typeface")
    return None


def ensure_complex_font(properties: ET.Element, preferred: str | None) -> tuple[bool, bool]:
    cs = direct_child(properties, A_CS)
    typeface = preferred or explicit_typeface(properties)
    if cs is not None:
        changed = bool(typeface) and set_attr(cs, "typeface", typeface)
        return changed, bool(cs.get("typeface"))
    if not typeface:
        return False, False
    cs = ET.Element(A_CS, {"typeface": typeface})
    insert_ordered(properties, cs, RPR_ORDER)
    return True, True


def effective_property(run: ET.Element, ppr: ET.Element | None, attribute: str) -> str | None:
    rpr = direct_child(run, A_RPR)
    if rpr is not None and rpr.get(attribute):
        return rpr.get(attribute)
    defrpr = direct_child(ppr, A_DEFRPR) if ppr is not None else None
    return defrpr.get(attribute) if defrpr is not None else None


def effective_font_nodes(run: ET.Element, ppr: ET.Element | None) -> tuple[ET.Element | None, ET.Element | None]:
    rpr = direct_child(run, A_RPR)
    defrpr = direct_child(ppr, A_DEFRPR) if ppr is not None else None
    cs = direct_child(rpr, A_CS)
    if cs is None:
        cs = direct_child(defrpr, A_CS)
    latin = direct_child(rpr, A_LATIN)
    if latin is None:
        latin = direct_child(defrpr, A_LATIN)
    return cs, latin


def audit_xml(part: str, data: bytes) -> tuple[int, int, int, list[Issue]]:
    root = ET.fromstring(data)
    parents = build_parent_map(root)
    arabic_count = 0
    mixed_count = 0
    native_bullets = 0
    issues: list[Issue] = []

    for element in root.iter():
        if element.text:
            controls = bidi_controls_in(element.text)
            if controls:
                issues.append(Issue(
                    "error", "bidi-control", part, None, None,
                    f"Manual bidi control characters found: {', '.join(sorted(set(controls)))}.",
                    element.text[:120],
                ))

    for index, paragraph in enumerate(root.iter(A_P), start=1):
        text = paragraph_text(paragraph)
        if not has_arabic_letters(text):
            continue
        arabic_count += 1
        if has_ltr_or_number(text):
            mixed_count += 1
        shape = nearest_shape_name(paragraph, parents)
        snippet = text[:160]
        ppr = direct_child(paragraph, A_PPR)

        if ppr is None or (ppr.get("rtl") or "").lower() not in TRUE_VALUES:
            issues.append(Issue(
                "error", "missing-paragraph-rtl", part, index, shape,
                "Arabic paragraph is missing native a:pPr/@rtl=1.", snippet,
            ))

        alignment = ppr.get("algn") if ppr is not None else None
        if alignment not in RTL_ALIGNMENTS:
            issues.append(Issue(
                "error", "bad-alignment", part, index, shape,
                f"Arabic paragraph alignment must be explicit RTL-compatible alignment; found {alignment!r}.", snippet,
            ))

        container = nearest_container(paragraph, parents)
        bodypr = direct_child(container, A_BODYPR)
        if container is not None and (bodypr is None or (bodypr.get("rtlCol") or "").lower() not in TRUE_VALUES):
            issues.append(Issue(
                "error", "missing-body-rtl", part, index, shape,
                "Arabic text body is missing a:bodyPr/@rtlCol=1.", snippet,
            ))
        elif container is None:
            issues.append(Issue(
                "notice", "unscoped-rich-text", part, index, shape,
                "Arabic paragraph has no auditable text-body container; verify it in PowerPoint.", snippet,
            ))

        if BULLET_RE.match(text):
            issues.append(Issue(
                "error", "literal-bullet", part, index, shape,
                "Paragraph begins with a typed bullet glyph; use a native DrawingML bullet.", snippet,
            ))

        if ppr is not None and any(local_name(child) in {"buChar", "buAutoNum", "buBlip"} for child in ppr):
            native_bullets += 1

        for run in direct_runs(paragraph):
            content = run_text(run)
            if not has_arabic_letters(content):
                continue
            language = effective_property(run, ppr, "lang")
            if not language or not language.lower().startswith("ar"):
                issues.append(Issue(
                    "warning", "missing-arabic-language", part, index, shape,
                    f"Arabic run lacks Arabic BCP-47 language metadata; found {language!r}.", content[:160],
                ))
            cs, latin = effective_font_nodes(run, ppr)
            if latin is not None and (cs is None or not cs.get("typeface")):
                issues.append(Issue(
                    "warning", "missing-complex-font", part, index, shape,
                    "An explicit Latin font exists but the Arabic complex-script font slot is missing.", content[:160],
                ))

    chart_text = "".join(element.text or "" for element in root.iter())
    if part.startswith("ppt/charts/") and has_arabic_letters(chart_text):
        issues.append(Issue(
            "notice", "arabic-chart", part, None, None,
            "Arabic appears in a native chart. Structural paragraph checks cannot prove every chart label; verify in PowerPoint.",
            chart_text[:160],
        ))

    return arabic_count, mixed_count, native_bullets, issues


def audit_pptx(path: Path, expect_presentation_rtl: bool = False) -> AuditResult:
    issues: list[Issue] = []
    xml_parts = 0
    arabic_count = 0
    mixed_count = 0
    native_bullets = 0
    presentation_rtl: bool | None = None

    with ZipFile(path, "r") as package:
        if "[Content_Types].xml" not in package.namelist():
            raise ValueError("Not an OOXML package: [Content_Types].xml is missing")
        for name in package.namelist():
            if not (name.startswith("ppt/") and name.endswith(".xml")):
                continue
            xml_parts += 1
            data = package.read(name)
            if name == "ppt/presentation.xml":
                try:
                    presentation_root = ET.fromstring(data)
                    presentation_rtl = (presentation_root.get("rtl") or "").lower() in TRUE_VALUES
                except ET.ParseError:
                    pass
            try:
                a_count, m_count, bullet_count, part_issues = audit_xml(name, data)
            except ET.ParseError as exc:
                issues.append(Issue("error", "invalid-xml", name, None, None, str(exc), ""))
                continue
            arabic_count += a_count
            mixed_count += m_count
            native_bullets += bullet_count
            issues.extend(part_issues)

    if expect_presentation_rtl and presentation_rtl is not True:
        issues.append(Issue(
            "error", "missing-presentation-rtl", "ppt/presentation.xml", None, None,
            "Arabic-first deck is missing p:presentation/@rtl=1.", "",
        ))

    return AuditResult(str(path), presentation_rtl, xml_parts, arabic_count, mixed_count, native_bullets, issues)


def first_paragraph_typeface(paragraph: ET.Element, ppr: ET.Element, preferred: str | None) -> str | None:
    if preferred:
        return preferred
    defrpr = direct_child(ppr, A_DEFRPR)
    candidate = explicit_typeface(defrpr)
    if candidate:
        return candidate
    for run in direct_runs(paragraph):
        candidate = explicit_typeface(direct_child(run, A_RPR))
        if candidate:
            return candidate
    return None


def convert_literal_bullet(paragraph: ET.Element, ppr: ET.Element) -> bool:
    first_text = next((node for node in paragraph.iter(A_T) if node.text and node.text.strip()), None)
    if first_text is None:
        return False
    match = BULLET_RE.match(first_text.text or "")
    if match is None:
        return False
    marker = match.group(1)
    first_text.text = (first_text.text or "")[match.end():]
    changed = True

    for child in list(ppr):
        if local_name(child) in {"buNone", "buAutoNum", "buChar", "buBlip"}:
            ppr.remove(child)
    insert_ordered(ppr, ET.Element(A_BUCHAR, {"char": marker}), PPR_ORDER)
    if not ppr.get("marR"):
        ppr.set("marR", "457200")
    if not ppr.get("indent"):
        ppr.set("indent", "-457200")
    return changed


def repair_xml(
    part: str,
    data: bytes,
    language: str,
    font: str | None,
    convert_bullets: bool,
    presentation_rtl: bool,
) -> tuple[bytes, bool, dict[str, int]]:
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True))
    root = ET.fromstring(data, parser=parser)
    changed = False
    stats = {"controls_removed": 0, "paragraphs_repaired": 0, "bullets_converted": 0}

    if part == "ppt/presentation.xml" and presentation_rtl:
        changed |= set_attr(root, "rtl", "1")

    for element in root.iter():
        if element.text:
            cleaned, removed = strip_bidi_controls(element.text)
            if removed:
                element.text = cleaned
                stats["controls_removed"] += removed
                changed = True

    parents = build_parent_map(root)
    for paragraph in root.iter(A_P):
        text = paragraph_text(paragraph)
        if not has_arabic_letters(text):
            continue

        paragraph_changed = False
        ppr, created = ensure_ppr(paragraph)
        paragraph_changed |= created
        paragraph_changed |= set_attr(ppr, "rtl", "1")
        if ppr.get("algn") not in RTL_ALIGNMENTS:
            paragraph_changed |= set_attr(ppr, "algn", "r")

        container = nearest_container(paragraph, parents)
        if container is not None:
            bodypr = direct_child(container, A_BODYPR)
            if bodypr is None:
                bodypr = ET.Element(A_BODYPR)
                container.insert(0, bodypr)
                paragraph_changed = True
            paragraph_changed |= set_attr(bodypr, "rtlCol", "1")

        paragraph_font = first_paragraph_typeface(paragraph, ppr, font)
        defrpr, created = ensure_defrpr(ppr)
        paragraph_changed |= created
        paragraph_changed |= set_attr(defrpr, "lang", language)
        font_changed, _ = ensure_complex_font(defrpr, paragraph_font)
        paragraph_changed |= font_changed

        for run in direct_runs(paragraph):
            content = run_text(run)
            if not has_arabic_letters(content):
                continue
            rpr, created = ensure_rpr(run)
            paragraph_changed |= created
            paragraph_changed |= set_attr(rpr, "lang", language)
            run_font = font or explicit_typeface(rpr) or paragraph_font
            font_changed, _ = ensure_complex_font(rpr, run_font)
            paragraph_changed |= font_changed

        endrpr = direct_child(paragraph, A_ENDRPR)
        if endrpr is not None:
            paragraph_changed |= set_attr(endrpr, "lang", language)
            font_changed, _ = ensure_complex_font(endrpr, font or paragraph_font)
            paragraph_changed |= font_changed

        if convert_bullets and convert_literal_bullet(paragraph, ppr):
            stats["bullets_converted"] += 1
            paragraph_changed = True

        if paragraph_changed:
            stats["paragraphs_repaired"] += 1
            changed = True

    if not changed:
        return data, False, stats
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), True, stats


def repair_pptx(
    source: Path,
    output: Path,
    language: str,
    font: str | None,
    convert_bullets: bool,
    presentation_rtl: bool,
    force: bool,
) -> dict[str, int]:
    if source.resolve() == output.resolve():
        raise ValueError("Repair requires a distinct output path; preserve the source PPTX")
    if output.exists() and not force:
        raise FileExistsError(f"Output already exists: {output}. Pass --force to replace it.")
    output.parent.mkdir(parents=True, exist_ok=True)
    totals = {"xml_parts_changed": 0, "controls_removed": 0, "paragraphs_repaired": 0, "bullets_converted": 0}

    handle, temp_name = tempfile.mkstemp(prefix=f".{output.stem}.", suffix=".pptx", dir=output.parent)
    os.close(handle)
    temp_path = Path(temp_name)
    try:
        with ZipFile(source, "r") as incoming, ZipFile(temp_path, "w") as outgoing:
            if "[Content_Types].xml" not in incoming.namelist():
                raise ValueError("Not an OOXML package: [Content_Types].xml is missing")
            outgoing.comment = incoming.comment
            for item in incoming.infolist():
                data = incoming.read(item.filename)
                if item.filename.startswith("ppt/") and item.filename.endswith(".xml"):
                    try:
                        data, part_changed, stats = repair_xml(
                            item.filename,
                            data,
                            language,
                            font,
                            convert_bullets,
                            presentation_rtl,
                        )
                    except ET.ParseError as exc:
                        raise ValueError(f"Invalid XML in {item.filename}: {exc}") from exc
                    if part_changed:
                        totals["xml_parts_changed"] += 1
                    for key in ("controls_removed", "paragraphs_repaired", "bullets_converted"):
                        totals[key] += stats[key]
                outgoing.writestr(item, data)
        os.replace(temp_path, output)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return totals


def result_to_json(result: AuditResult) -> str:
    payload = asdict(result)
    payload["counts"] = result.counts()
    return json.dumps(payload, ensure_ascii=False, indent=2)


def print_result(result: AuditResult, strict: bool, max_issues: int = 80) -> int:
    counts = result.counts()
    print(f"Arabic PPTX audit: {result.file}")
    print(
        f"Presentation RTL: "
        f"{'yes' if result.presentation_rtl else 'no' if result.presentation_rtl is False else 'unknown'} | "
        f"XML parts: {result.xml_parts} | Arabic paragraphs: {result.arabic_paragraphs} | "
        f"mixed paragraphs: {result.mixed_paragraphs} | native bullets: {result.native_bullets}"
    )
    visible_issues = result.issues if max_issues == 0 else result.issues[:max_issues]
    for issue in visible_issues:
        location = issue.part
        if issue.paragraph is not None:
            location += f":paragraph-{issue.paragraph}"
        if issue.shape:
            location += f":{issue.shape}"
        print(f"{issue.severity.upper()} [{issue.code}] {location}: {issue.message}")
        if issue.text:
            print(f"  text: {issue.text!r}")
    omitted = len(result.issues) - len(visible_issues)
    if omitted:
        print(f"... {omitted} additional issues omitted; use --max-issues 0 or --json to show all.")
    failed = counts["error"] > 0 or (strict and counts["warning"] > 0)
    status = "FAIL" if failed else "PASS"
    print(
        f"{status}: errors={counts['error']} warnings={counts['warning']} "
        f"notices={counts['notice']} strict={'yes' if strict else 'no'}"
    )
    return 1 if failed else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    audit = subparsers.add_parser("audit", help="Inspect Arabic PPTX structure without modifying the file")
    audit.add_argument("file", type=Path)
    audit.add_argument("--strict", action="store_true", help="Treat warnings as delivery blockers")
    audit.add_argument(
        "--expect-presentation-rtl",
        action="store_true",
        help="Require p:presentation/@rtl=1 for an Arabic-first deck",
    )
    audit.add_argument("--json", action="store_true", help="Emit machine-readable JSON")
    audit.add_argument(
        "--max-issues",
        type=int,
        default=80,
        help="Maximum issues to print; use 0 for all (default: 80)",
    )

    repair = subparsers.add_parser("repair", help="Write a normalized copy, then run a strict audit")
    repair.add_argument("source", type=Path)
    repair.add_argument("output", type=Path)
    repair.add_argument("--lang", default="ar-SA", help="BCP-47 language tag for Arabic runs (default: ar-SA)")
    repair.add_argument("--font", help="Arabic-capable typeface to write into the complex-script slot; does not embed it")
    repair.add_argument("--convert-bullets", action="store_true", help="Convert leading bullet glyphs to native bullets")
    repair.add_argument(
        "--presentation-rtl",
        action="store_true",
        help="Set p:presentation/@rtl=1 for an Arabic-first deck",
    )
    repair.add_argument("--force", action="store_true", help="Replace an existing output file")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        if args.command == "audit":
            result = audit_pptx(args.file, expect_presentation_rtl=args.expect_presentation_rtl)
            if args.json:
                print(result_to_json(result))
                counts = result.counts()
                return 1 if counts["error"] or (args.strict and counts["warning"]) else 0
            return print_result(result, args.strict, max_issues=max(0, args.max_issues))

        totals = repair_pptx(
            args.source,
            args.output,
            args.lang,
            args.font,
            args.convert_bullets,
            args.presentation_rtl,
            args.force,
        )
        print(
            "Repair complete: "
            + " | ".join(f"{key}={value}" for key, value in totals.items())
        )
        return print_result(
            audit_pptx(args.output, expect_presentation_rtl=args.presentation_rtl),
            strict=True,
        )
    except (BadZipFile, FileNotFoundError, FileExistsError, PermissionError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
