"""Automatic table and header detection."""
from __future__ import annotations

import re

from .diagnostics import ParseDiagnostics
from .normalization import low, norm

HEADER_ALIASES = {
    "num": ["№ п/п", "номер", "№"],
    "scheme": ["схема сети", "ремонтная схема"],
    "temp": ["тнв", "температура", "условие 2", "гр. уст. арпм", "группа уставок арпм"],
    "mdp": ["мдп без па"],
    "mdp_pa": ["мдп с па"],
    "adp": ["адп"],
    "criteria": ["критерий определения", "критерии определения"],
    "control": ["контроль дополнительных параметров"],
    "note": ["примечание к схеме", "примечание"],
}


def find_sheet(sheets: dict, token: str) -> tuple[str | None, list[list[str]]]:
    token = low(token)
    for name, rows in sheets.items():
        if token in low(name):
            return name, rows
    return None, []


def header_map(rows: list[list[str]], diag: ParseDiagnostics | None = None) -> dict[str, tuple[int, int]]:
    best: dict[str, tuple[int, int]] = {}
    scan_rows = min(12, len(rows))
    for row_idx, row in enumerate(rows[:scan_rows]):
        for col_idx, value in enumerate(row):
            lv = low(value)
            for key, aliases in HEADER_ALIASES.items():
                if any(alias in lv for alias in aliases):
                    best.setdefault(key, (col_idx, row_idx))

    if "criteria" in best:
        start_col, header_row = best["criteria"]
        # The criteria group ends where the following top-level group starts.
        # In the source workbooks that is usually "Контроль дополнительных
        # параметров".  Excluding it made us accidentally read the *control*
        # subheaders as criteria and silently shifted all criteria columns.
        next_starts = [
            pos[0]
            for key, pos in best.items()
            if pos[0] > start_col and key not in {"mdp", "mdp_pa", "adp"}
        ]
        end_col = min(next_starts) if next_starts else start_col + 4
        for col in range(start_col, end_col):
            texts = " ".join(
                low(rows[r][col] if col < len(rows[r]) else "")
                for r in range(min(5, len(rows)))
            )
            if "без па" in texts and "мдп" in texts:
                best["criteria_mdp"] = (col, header_row)
            elif "с па" in texts and "мдп" in texts:
                best["criteria_mdp_pa"] = (col, header_row)
            elif "адп" in texts:
                best["criteria_adp"] = (col, header_row)
        best.setdefault("criteria_mdp", (start_col, header_row))

        # The ADP header is often centred over a merged range while actual
        # values are anchored in another column of that range.  Select the
        # populated column after the last MDP criteria anchor and before the
        # service/control group.  This covers both left- and right-anchored
        # templates without filename-specific rules.
        explicit_adp = best.get("criteria_adp")
        explicit_populated = bool(
            explicit_adp and _body_content_count(rows, explicit_adp[0], scan_rows) > 0
        )
        if "adp" in best and not explicit_populated:
            previous = best.get("criteria_mdp_pa", best.get("criteria_mdp", (start_col, header_row)))[0]
            candidates = list(range(previous + 1, end_col))
            if candidates:
                counts = {col: _body_content_count(rows, col, scan_rows) for col in candidates}
                populated = [col for col in candidates if counts[col] > 0]
                if populated:
                    # Merged criteria cells are often anchored one column to
                    # the left of their visible subheader.  The nearest
                    # populated column is the correct anchor; choosing the
                    # densest one incorrectly selects an AОПО/PA criterion.
                    adp_col = max(populated)
                    best["criteria_adp"] = (adp_col, header_row)

    if "mdp" not in best and diag:
        diag.warn("Не найден заголовок «МДП без ПА»")
    if "scheme" not in best and diag:
        diag.warn("Не найден заголовок «Схема сети»")
    return best


def cell(row: list[str], header_map_: dict[str, tuple[int, int]], key: str) -> str:
    col, _ = header_map_.get(key, (999, 0))
    return norm(row[col]) if col < len(row) else ""


def column_has_content(rows: list[list[str]], header_map_: dict[str, tuple[int, int]], key: str, start_row: int) -> bool:
    col, _ = header_map_.get(key, (999, 0))
    if col == 999:
        return False
    for row in rows[start_row:]:
        value = norm(row[col]) if col < len(row) else ""
        if value and low(value) not in ("-", "—") and not low(value).startswith("минимальное из"):
            if key in ("mdp", "mdp_pa") and not _meaningful_mdp_value(value):
                continue
            return True
    return False


def detect_columns(rows: list[list[str]], header_map_: dict[str, tuple[int, int]]) -> tuple[bool, bool, bool]:
    start = body_start_row(rows, header_map_)
    has_mdp = "mdp" in header_map_ and column_has_content(rows, header_map_, "mdp", start)
    has_mdp_pa = bool(mdp_pa_columns(rows, header_map_, start))
    has_adp = "adp" in header_map_ and column_has_content(rows, header_map_, "adp", start)
    return has_mdp, has_mdp_pa, has_adp


def body_start_row(rows: list[list[str]], header_map_: dict[str, tuple[int, int]]) -> int:
    """Find the first physical data row after a multi-level header."""
    minimum = max((row for _, row in header_map_.values()), default=0) + 1
    num_col = header_map_.get("num", (999, 0))[0]
    scheme_col = header_map_.get("scheme", (999, 0))[0]
    for index, row in enumerate(rows[minimum:], minimum):
        num = norm(row[num_col]) if num_col < len(row) else ""
        scheme = norm(row[scheme_col]) if scheme_col < len(row) else ""
        if (num or scheme) and "схема сети" not in low(scheme) and "№" not in num:
            return index
    return minimum


def mdp_pa_columns(
    rows: list[list[str]], header_map_: dict[str, tuple[int, int]], start_row: int | None = None
) -> list[int]:
    """Return every PA value column between the base MDP and ADP columns.

    Several official workbooks label these columns with the automation name
    (АОПО/АПНУ) instead of the literal heading ``МДП с ПА``.  They still form
    one logical MDP-with-PA group in the output.
    """
    if "adp" not in header_map_:
        return []
    if "mdp" in header_map_:
        first = header_map_["mdp"][0] + 1
    elif "mdp_pa" in header_map_:
        first = header_map_["mdp_pa"][0]
    else:
        return []
    last = header_map_["adp"][0]
    start = start_row if start_row is not None else max((r for _, r in header_map_.values()), default=1) + 1
    result: list[int] = []
    for col in range(first, last):
        values = [norm(row[col]) for row in rows[start:] if col < len(row)]
        if any(_meaningful_mdp_value(v) for v in values):
            result.append(col)
    return result


def mdp_pa_criteria_columns(
    rows: list[list[str]], header_map_: dict[str, tuple[int, int]], start_row: int | None = None
) -> list[int]:
    """Return populated PA criteria columns in their source order."""
    if "criteria_mdp" not in header_map_ or "criteria_adp" not in header_map_:
        return []
    first = header_map_["criteria_mdp"][0] + 1
    last = header_map_["criteria_adp"][0]
    start = start_row if start_row is not None else max((r for _, r in header_map_.values()), default=1) + 1
    return [
        col
        for col in range(first, last)
        if _body_content_count(rows, col, start) > 0
    ]


def control_columns(
    rows: list[list[str]],
    header_map_: dict[str, tuple[int, int]],
    has_mdp: bool,
    has_mdp_pa: bool,
) -> dict[str, list[int]]:
    """Map the variable-width control group to three logical output columns."""
    if "control" not in header_map_:
        return {"mdp": [], "mdp_pa": [], "adp": []}
    first = header_map_["control"][0]
    last = header_map_.get("note", (max((len(row) for row in rows), default=first), 0))[0]
    header_end = body_start_row(rows, header_map_)
    headers = {
        col: " ".join(
            low(rows[row][col] if col < len(rows[row]) else "")
            for row in range(min(header_end, len(rows)))
        )
        for col in range(first, last)
    }
    adp = [col for col, text in headers.items() if "адп" in text]
    remaining = [col for col in range(first, last) if col not in adp]
    pa_explicit = [
        col
        for col in remaining
        if ("с па" in headers[col] or ("па" in headers[col] and "без па" not in headers[col]))
        and "мдп без па" not in headers[col]
    ]
    if has_mdp_pa:
        # Named PA devices are used as subheaders instead of the words
        # “МДП с ПА”.  A non-empty header after the base anchor belongs to PA.
        pa_named = [
            col
            for col in remaining
            if col != first
            and headers[col]
            and "мдп без па" not in headers[col]
            and "контроль дополнительных параметров" not in headers[col]
        ]
        pa = sorted(set(pa_explicit + pa_named))
    else:
        pa = []
    mdp = [col for col in remaining if col not in pa] if has_mdp else []
    return {"mdp": mdp, "mdp_pa": pa, "adp": adp}


def _meaningful_mdp_value(value: str) -> bool:
    """Reject technical placeholders such as ``1)`` or numbered dashes."""
    text = re.sub(r"^\s*\d+\)\s*", "", norm(value))
    return bool(text and text not in ("-", "—", "–", "−"))


def _body_content_count(rows: list[list[str]], col: int, start_row: int) -> int:
    count = 0
    for row in rows[start_row:]:
        value = norm(row[col]) if col < len(row) else ""
        if value and value not in ("-", "—", "–", "−"):
            count += 1
    return count


def row_axis_label(rows: list[list[str]], header_map_: dict[str, tuple[int, int]]) -> str:
    """Return the first-column heading for logical data rows (TNV or ARPM group)."""
    col, _ = header_map_.get("temp", (999, 0))
    if col == 999:
        return "ТНВ"
    for row_idx in range(min(6, len(rows))):
        if col >= len(rows[row_idx]):
            continue
        text = low(rows[row_idx][col])
        if "арпм" in text and ("уст" in text or "устав" in text):
            return "Группа уставок АРПМ"
        if "тнв" in text or "температура" in text:
            return "ТНВ"
    return "ТНВ"
