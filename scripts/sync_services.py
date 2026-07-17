#!/usr/bin/env python3
"""Refresh the calculator's embedded service catalog from the source XLSX."""

from __future__ import annotations

import argparse
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DATA_START = "    /* SERVICES_DATA_START */"
DATA_END = "    /* SERVICES_DATA_END */"


def column_index(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference.upper())
    if not letters:
        raise ValueError(f"Invalid cell reference: {reference}")
    value = 0
    for letter in letters.group(0):
        value = value * 26 + ord(letter) - ord("A") + 1
    return value - 1


def shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(item.itertext()) for item in root.findall(f"{{{MAIN_NS}}}si")]


def worksheet_paths(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        item.attrib["Id"]: item.attrib["Target"]
        for item in relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }
    sheets = []
    for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet"):
        relationship_id = sheet.attrib[f"{{{REL_NS}}}id"]
        target = targets[relationship_id]
        path = posixpath.normpath(posixpath.join("xl", target))
        sheets.append((sheet.attrib.get("name", path), path))
    return sheets


def cell_value(cell: ET.Element, strings: list[str]):
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(cell.itertext())
    value = cell.find(f"{{{MAIN_NS}}}v")
    if value is None or value.text is None:
        return ""
    raw = value.text
    if cell_type == "s":
        return strings[int(raw)]
    if cell_type in {"str", "e"}:
        return raw
    try:
        number = float(raw)
        return int(number) if number.is_integer() else number
    except ValueError:
        return raw


def worksheet_rows(archive: zipfile.ZipFile, path: str, strings: list[str]):
    root = ET.fromstring(archive.read(path))
    for row in root.findall(f".//{{{MAIN_NS}}}row"):
        values: dict[int, object] = {}
        for cell in row.findall(f"{{{MAIN_NS}}}c"):
            reference = cell.attrib.get("r", "")
            values[column_index(reference)] = cell_value(cell, strings)
        yield values


def normalize_header(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def parse_rate(value: object, vendor: bool):
    if value is None or str(value).strip() == "":
        return None, None, "custom"

    raw = str(value).strip().replace("−", "-").replace("–", "-")
    range_match = re.fullmatch(
        r"\s*(\d+(?:\.\d+)?)\s*%?\s*-\s*(\d+(?:\.\d+)?)\s*%\s*",
        raw,
    )
    if range_match:
        low, high = (float(part) / 100 for part in range_match.groups())
        return min(low, high), max(low, high), "range"

    percent_value = raw.endswith("%")
    numeric = float(raw.removesuffix("%").strip())
    if percent_value or numeric > 1:
        numeric /= 100
    if not 0 <= numeric <= 1:
        raise ValueError(f"Discount is outside the supported 0-100% range: {value}")
    return numeric, numeric, "up-to" if vendor else "exact"


def parse_services(workbook_path: Path) -> list[dict[str, object]]:
    with zipfile.ZipFile(workbook_path) as archive:
        strings = shared_strings(archive)
        offers: list[dict[str, object]] = []
        seen: set[tuple[str, str]] = set()
        counters = {"Partner service": 0, "Vendor service": 0}

        for _, path in worksheet_paths(archive):
            rows = list(worksheet_rows(archive, path, strings))
            header_index = None
            columns: dict[str, int] = {}
            for index, row in enumerate(rows):
                headers = {normalize_header(value): column for column, value in row.items()}
                classification = next(
                    (column for header, column in headers.items() if header.startswith("classification")),
                    None,
                )
                service = next(
                    (column for header, column in headers.items() if header == "service"),
                    None,
                )
                if classification is not None and service is not None:
                    columns["classification"] = classification
                    columns["service"] = service
                    columns["discount"] = next(
                        (column for header, column in headers.items() if header.startswith("discount")),
                        -1,
                    )
                    columns["note"] = next(
                        (column for header, column in headers.items() if "what is required" in header and "negotiations" in header),
                        -1,
                    )
                    header_index = index
                    break

            if header_index is None:
                continue

            for row in rows[header_index + 1 :]:
                raw_classification = normalize_header(row.get(columns["classification"], ""))
                if raw_classification == "partner service":
                    classification = "Partner service"
                elif raw_classification == "vendor service":
                    classification = "Vendor service"
                else:
                    continue

                name = str(row.get(columns["service"], "") or "").strip()
                if not name:
                    continue
                key = (classification, name.casefold())
                if key in seen:
                    continue
                seen.add(key)

                source_index = counters[classification]
                counters[classification] += 1
                discount = row.get(columns["discount"], "") if columns["discount"] >= 0 else ""
                minimum, maximum, discount_type = parse_rate(
                    discount,
                    vendor=classification == "Vendor service",
                )
                note_value = row.get(columns["note"], "") if columns["note"] >= 0 else ""
                note = str(note_value or "").strip() or None
                prefix = "partner" if classification == "Partner service" else "vendor"
                offers.append(
                    {
                        "id": f"{prefix}-{source_index}",
                        "classification": classification,
                        "name": name,
                        "discountMin": minimum,
                        "discountMax": maximum,
                        "discountType": discount_type,
                        "negotiationNote": note,
                    }
                )

    partner_count = sum(item["classification"] == "Partner service" for item in offers)
    vendor_count = sum(item["classification"] == "Vendor service" for item in offers)
    if partner_count == 0 or vendor_count == 0 or len(offers) < 20:
        raise ValueError(
            f"Workbook did not contain the expected service catalog: "
            f"{partner_count} partner and {vendor_count} vendor rows"
        )
    return offers


def update_html(html_path: Path, offers: list[dict[str, object]]) -> bool:
    html = html_path.read_text(encoding="utf-8")
    if html.count(DATA_START) != 1 or html.count(DATA_END) != 1:
        raise ValueError("The calculator data markers are missing or duplicated")
    start = html.index(DATA_START)
    end = html.index(DATA_END, start) + len(DATA_END)
    payload = json.dumps(offers, ensure_ascii=False, separators=(",", ":"))
    replacement = f"{DATA_START}\n    const services = {payload};\n{DATA_END}"
    updated = html[:start] + replacement + html[end:]
    if updated == html:
        return False
    html_path.write_text(updated, encoding="utf-8")
    return True


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("html", type=Path)
    args = parser.parse_args()
    offers = parse_services(args.workbook)
    changed = update_html(args.html, offers)
    partner_count = sum(item["classification"] == "Partner service" for item in offers)
    vendor_count = len(offers) - partner_count
    print(
        f"Synced {len(offers)} services "
        f"({partner_count} partner, {vendor_count} vendor); "
        f"calculator {'updated' if changed else 'already current'}."
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        print(f"Sync failed: {error}", file=sys.stderr)
        raise SystemExit(1)
