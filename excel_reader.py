"""Minimal Excel reader for the included QA plan.

This module avoids third-party dependencies by parsing the XLSX file
as a zip archive with XML content. It exposes helpers to extract headers
and rows from a sheet.
"""
from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile
from typing import Iterable, List, Tuple

# Namespaces used by XLSX files
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

EXCEL_FILE = Path(__file__).resolve().parent / "QA_Planilla_Revision_Manual_Meridian_Transelec.xlsx"


def _cell_text(cell: ET.Element) -> str:
    """Extract the text contained in a cell element.

    The workbook uses inline strings instead of shared strings, so we
    need to check for the ``inlineStr`` type and look for the nested
    ``t`` element.
    """

    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        inline = cell.find(f"{_NS}is")
        if inline is not None:
            text_node = inline.find(f"{_NS}t")
            if text_node is not None:
                return text_node.text or ""
        return ""

    value = cell.find(f"{_NS}v")
    return value.text if value is not None else ""


def _sheet_rows(sheet_xml: bytes) -> Iterable[List[str]]:
    """Yield rows as lists of string values from the sheet XML."""

    sheet = ET.fromstring(sheet_xml)
    for row in sheet.findall(f".{_NS}sheetData/{_NS}row"):
        yield [_cell_text(cell) for cell in row.findall(f"{_NS}c")]


def load_sheet(sheet_number: int = 1) -> Tuple[List[str], List[dict]]:
    """Load headers and row dictionaries from a worksheet.

    Args:
        sheet_number: Worksheet number (1-indexed) inside the workbook.

    Returns:
        A tuple ``(headers, rows)`` where ``headers`` is a list of column
        names from the first row and ``rows`` is a list of dictionaries
        mapping those headers to row values.
    """

    sheet_path = f"xl/worksheets/sheet{sheet_number}.xml"
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {EXCEL_FILE}")

    with zipfile.ZipFile(EXCEL_FILE) as archive:
        try:
            sheet_bytes = archive.read(sheet_path)
        except KeyError as exc:  # sheet does not exist
            raise FileNotFoundError(
                f"No se encontró la hoja {sheet_number} en el archivo XLSX"
            ) from exc

    rows = list(_sheet_rows(sheet_bytes))
    if not rows:
        return [], []

    headers = rows[0]
    data_rows = []
    for row in rows[1:]:
        # Skip completely empty rows
        if all(cell == "" for cell in row):
            continue

        max_len = max(len(headers), len(row))
        item = {}
        for idx in range(max_len):
            key = headers[idx] if idx < len(headers) else f"col_{idx + 1}"
            value = row[idx] if idx < len(row) else ""
            item[key] = value
        data_rows.append(item)

    return headers, data_rows
