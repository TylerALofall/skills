#!/usr/bin/env python3
"""
Generate a cover page by swapping scoped placeholders inside a DOCX template.

- Reads the DOCX as a zip file
- Replaces specific <w:t> contents (case number, filing name, judge)
- Validates the XML with defusedxml before writing
- Repackages without touching the source template or spawning subprocesses
"""

import argparse
import html
import re
import sys
import zipfile
from pathlib import Path
from typing import Dict, Tuple

from defusedxml import ElementTree


PLACEHOLDERS: Dict[str, str] = {
    "case_number": "No. 6461",
    "filing_name": "APPELLANTS OPENING BRIEF",
    "judge": "Hon. Stacy Beckerman",
}


def _read_document_xml(template_path: Path) -> str:
    with zipfile.ZipFile(template_path, "r") as docx:
        try:
            xml_bytes = docx.read("word/document.xml")
        except KeyError as exc:
            raise ValueError("word/document.xml not found in template") from exc
    return xml_bytes.decode("utf-8")


def _validate_xml(xml_text: str, context: str) -> None:
    try:
        ElementTree.fromstring(xml_text)
    except ElementTree.ParseError as exc:
        raise ValueError(f"XML became invalid after replacing {context}: {exc}") from exc


def _scoped_replace(xml_text: str, placeholder: str, replacement: str) -> Tuple[str, int]:
    pattern = re.compile(rf"(<w:t[^>]*>){re.escape(placeholder)}(</w:t>)")
    matches = list(pattern.finditer(xml_text))
    if not matches:
        raise ValueError(f"Placeholder '{placeholder}' not found in any <w:t> node")

    escaped_value = html.escape(replacement)
    updated = pattern.sub(lambda m: f"{m.group(1)}{escaped_value}{m.group(2)}", xml_text)
    return updated, len(matches)


def apply_replacements(xml_text: str, values: Dict[str, str]) -> Tuple[str, Dict[str, int]]:
    updated = xml_text
    counts: Dict[str, int] = {}
    for key, placeholder in PLACEHOLDERS.items():
        replacement = values.get(key)
        if replacement is None:
            raise ValueError(f"Missing replacement value for {key}")

        updated, count = _scoped_replace(updated, placeholder, replacement)
        _validate_xml(updated, placeholder)
        counts[key] = count

    return updated, counts


def write_docx(template_path: Path, output_path: Path, updated_xml: str) -> None:
    if output_path.resolve() == template_path.resolve():
        raise ValueError("Output path must differ from template to preserve the source file")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(template_path, "r") as src, zipfile.ZipFile(output_path, "w") as dst:
        for info in src.infolist():
            data = updated_xml.encode("utf-8") if info.filename == "word/document.xml" else src.read(info)
            dst.writestr(info, data)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a cover DOCX from a template")
    parser.add_argument("template", type=Path, help="Path to the DOCX template")
    parser.add_argument("output", type=Path, help="Path for the generated DOCX")
    parser.add_argument("--case-number", required=True, help="Case number text (e.g., 'No. 1234')")
    parser.add_argument(
        "--filing-name", required=True, help="Filing name text (e.g., 'APPELLANT'S BRIEF')"
    )
    parser.add_argument("--judge", required=True, help="Judge name text")
    parser.add_argument("--dry-run", action="store_true", help="Preview replacements without writing output")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    template = args.template
    output = args.output

    if not template.exists():
        raise SystemExit(f"Template not found: {template}")

    xml_text = _read_document_xml(template)

    replacement_values = {
        "case_number": args.case_number,
        "filing_name": args.filing_name,
        "judge": args.judge,
    }
    updated_xml, counts = apply_replacements(xml_text, replacement_values)

    if args.dry_run:
        print("Dry run: proposed replacements")
        for key in ("case_number", "filing_name", "judge"):
            placeholder = PLACEHOLDERS[key]
            replacement = replacement_values[key]
            count = counts.get(key, 0)
            print(f"- {key}: '{placeholder}' -> '{replacement}' (matches: {count})")
        print("No files were written (dry run).")
        return

    write_docx(template, output, updated_xml)
    print(f"Generated cover saved to {output}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
