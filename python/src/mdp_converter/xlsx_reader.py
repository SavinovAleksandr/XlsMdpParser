"""Direct OOXML XLSX reader."""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def colnum(ref: str) -> int:
    letters = "".join(c for c in ref if c.isalpha())
    num = 0
    for char in letters:
        num = num * 26 + ord(char.upper()) - 64
    return num


def _cell_value(cell: ET.Element, shared: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value_el = cell.find("a:v", NS)
    formula_el = cell.find("a:f", NS)
    val = ""
    if cell_type == "inlineStr":
        val = "".join(x.text or "" for x in cell.findall(".//a:t", NS))
    elif value_el is not None:
        val = value_el.text or ""
        if cell_type == "s":
            val = shared[int(val)]
    if formula_el is not None and not val.strip():
        val = "=" + (formula_el.text or "")
    return val


def read_xlsx(path: str | Path) -> dict[str, list[list[str]]]:
    with zipfile.ZipFile(path) as archive:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", NS):
                shared.append("".join(t.text or "" for t in item.iter(f"{{{NS['a']}}}t")))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {x.attrib["Id"]: x.attrib["Target"] for x in rels}

        sheets: dict[str, list[list[str]]] = {}
        for sheet in workbook.find("a:sheets", NS):
            name = sheet.attrib["name"]
            target = rel_map[sheet.attrib[f"{{{NS['r']}}}id"]].lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            root = ET.fromstring(archive.read(target))

            merge_map: dict[tuple[int, int], str] = {}
            merge_cells = root.find("a:mergeCells", NS)
            if merge_cells is not None:
                for merge in merge_cells.findall("a:mergeCell", NS):
                    ref = merge.attrib.get("ref", "")
                    if ":" not in ref:
                        continue
                    start, end = ref.split(":", 1)
                    start_col, start_row = _split_ref(start)
                    end_col, end_row = _split_ref(end)
                    # Value will be filled after row scan; placeholder key only.
                    merge_map[(start_col, start_row, end_col, end_row)] = ""

            rows: list[list[str]] = []
            row_cells: dict[int, dict[int, str]] = {}
            for row in root.findall(".//a:sheetData/a:row", NS):
                row_idx = int(row.attrib.get("r", "0") or 0)
                cells: dict[int, str] = {}
                for cell in row.findall("a:c", NS):
                    ref = cell.attrib.get("r", "A1")
                    col = colnum(ref)
                    cells[col] = _cell_value(cell, shared)
                if cells:
                    row_cells[row_idx] = cells

            if row_cells:
                max_row = max(row_cells)
                max_col = max(max(cells) for cells in row_cells.values())
                for row_idx in range(1, max_row + 1):
                    cells = row_cells.get(row_idx, {})
                    # Apply merged cell fill-down/right from top-left value.
                    for (sc, sr, ec, er), _ in list(merge_map.items()):
                        if sr <= row_idx <= er:
                            anchor = row_cells.get(sr, {}).get(sc, "")
                            if anchor:
                                for r in range(sr, er + 1):
                                    for c in range(sc, ec + 1):
                                        if r == row_idx:
                                            cells.setdefault(c, anchor)
                    rows.append([cells.get(i, "") for i in range(1, max_col + 1)])

            sheets[name] = rows
        return sheets


def _split_ref(ref: str) -> tuple[int, int]:
    match = re.match(r"([A-Za-z]+)(\d+)", ref)
    if not match:
        return 1, 1
    return colnum(match.group(1)), int(match.group(2))
