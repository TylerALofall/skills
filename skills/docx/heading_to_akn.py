#!/usr/bin/env python3
"""
Extract heading-based structure from a DOCX and emit Akoma Ntoso XML.

Features:
- Reads word/document.xml from a DOCX zip (no subprocesses).
- Detects headings (H1–H4a / Heading1–Heading4a styles) and assigns the text
  that follows to the nearest ancestor heading ("possession" until the next
  heading of the same or higher level).
- Skips boilerplate such as "Case No. XX-XXXX" banners and center page numbers.
- Builds a nested Akoma Ntoso `<section>` hierarchy with `<heading>` and `<p>` nodes.
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Iterator, List, Optional, Tuple

from defusedxml import ElementTree as SafeET
from xml.etree import ElementTree as ET  # Standard lib for emission

# Supported heading style names mapped to levels
HEADING_STYLE_LEVELS: dict[str, int] = {
    "heading1": 1,
    "h1": 1,
    "heading2": 2,
    "h2": 2,
    "heading3": 3,
    "h3": 3,
    "heading4": 4,
    "heading4a": 4,
    "h4": 4,
    "h4a": 4,
}

# Patterns for skipping boilerplate paragraphs
CASE_NO_PATTERN = re.compile(r"(?i)^case\s+no\.?\s+[\w-]+$")
PAGE_NUMBER_PATTERN = re.compile(r"^\d+(\s+of\s+\d+)?$", re.IGNORECASE)


@dataclass
class Section:
    title: str
    level: int
    content: list[str] = field(default_factory=list)
    children: list["Section"] = field(default_factory=list)


def _read_document_xml(docx_path: Path) -> str:
    with zipfile.ZipFile(docx_path, "r") as archive:
        try:
            return archive.read("word/document.xml").decode("utf-8")
        except KeyError as exc:
            raise ValueError("word/document.xml not found in DOCX") from exc


def _extract_paragraphs(document_xml: str) -> Iterator[Tuple[str, Optional[str]]]:
    """
    Yield (text, style) for each paragraph.

    style is the raw w:pStyle value (lowercased), or None.
    """
    root = SafeET.fromstring(document_xml)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for p in root.findall(".//w:p", ns):
        style_elem = p.find("w:pPr/w:pStyle", ns)
        style_val = style_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if style_elem is not None else None
        text_parts = [t.text or "" for t in p.findall(".//w:t", ns)]
        text = "".join(text_parts).strip()
        style = style_val.lower() if style_val else None
        yield text, style


def _heading_level(style: Optional[str]) -> Optional[int]:
    if not style:
        return None
    if style.lower() in HEADING_STYLE_LEVELS:
        return HEADING_STYLE_LEVELS[style.lower()]
    # Fallback: HeadingX or HX patterns
    match = re.match(r"(?i)^heading\s*([1-4])([aA]?)$", style)
    if not match:
        match = re.match(r"(?i)^h([1-4])([aA]?)$", style)
    if match:
        level = int(match.group(1))
        return min(level, 4)
    return None


def _should_skip_paragraph(text: str) -> bool:
    if not text.strip():
        return True
    if CASE_NO_PATTERN.match(text.strip()):
        return True
    if PAGE_NUMBER_PATTERN.match(text.strip()):
        return True
    return False


def _slugify(text: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "-", text.strip().lower()).strip("-")
    return cleaned or "section"


def _build_sections(paragraphs: Iterable[Tuple[str, Optional[str]]]) -> Section:
    root = Section(title="root", level=0)
    stack: List[Section] = [root]

    for text, style in paragraphs:
        level = _heading_level(style)
        if level:
            # New heading; pop until parent level is lower
            while stack and stack[-1].level >= level:
                stack.pop()
            parent = stack[-1] if stack else root
            section = Section(title=text or f"Heading{level}", level=level)
            parent.children.append(section)
            stack.append(section)
        else:
            if _should_skip_paragraph(text):
                continue
            stack[-1].content.append(text)

    return root


def _section_to_xml(section: Section) -> ET.Element:
    elem = ET.Element("section", attrib={"id": _slugify(section.title), "level": str(section.level)})
    heading = ET.SubElement(elem, "heading")
    heading.text = section.title
    for para in section.content:
        p = ET.SubElement(elem, "p")
        p.text = para
    for child in section.children:
        elem.append(_section_to_xml(child))
    return elem


def build_akn_tree(root: Section) -> ET.Element:
    akn = ET.Element("akomaNtoso")
    judgment = ET.SubElement(akn, "judgment")
    body = ET.SubElement(judgment, "body")
    for child in root.children:
        body.append(_section_to_xml(child))
    return akn


def save_xml(tree: ET.Element, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    ET.ElementTree(tree).write(output_path, encoding="utf-8", xml_declaration=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract DOCX headings into Akoma Ntoso XML.")
    parser.add_argument("docx_path", type=Path, help="Path to the source DOCX file.")
    parser.add_argument("output_xml", type=Path, help="Path to write the Akoma Ntoso XML.")
    parser.add_argument("--toc", action="store_true", help="Print the detected heading outline (level + text).")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if not args.docx_path.exists():
        raise SystemExit(f"Input DOCX not found: {args.docx_path}")

    document_xml = _read_document_xml(args.docx_path)
    paragraphs = list(_extract_paragraphs(document_xml))
    sections_root = _build_sections(paragraphs)

    if args.toc:
        for section in sections_root.children:
            _print_outline(section)

    akn_tree = build_akn_tree(sections_root)
    save_xml(akn_tree, args.output_xml)
    print(f"Akoma Ntoso XML written to {args.output_xml}")


def _print_outline(section: Section, indent: int = 0) -> None:
    prefix = "  " * indent
    print(f"{prefix}H{section.level}: {section.title}")
    for child in section.children:
        _print_outline(child, indent + 1)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
