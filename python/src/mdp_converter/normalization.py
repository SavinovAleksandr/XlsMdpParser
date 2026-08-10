"""Text normalization utilities."""
from __future__ import annotations

import re

DASHES = frozenset({"-", "—", "–", "−"})


def norm(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def low(value: object) -> str:
    return norm(value).lower().replace("ё", "е")


def is_empty_or_dash(value: str) -> bool:
    v = norm(value)
    return not v or v in DASHES


def is_non_formula_text(value: str) -> bool:
    """Return True if cell content is explanatory text, not a formula."""
    v = norm(value)
    if not v or is_empty_or_dash(v):
        return True
    lv = low(v)
    if lv.startswith("минимальное из"):
        return True
    if lv.startswith("не контролируется"):
        return True
    # Pure number without operators is still a formula (constant).
    if re.search(r"[+\-*/^=<>]", v):
        return False
    if re.match(r"^if\s*\(", v, re.I):
        return False
    if re.search(r"[A-Za-zА-Яа-яЁё_]", v) and re.search(r"[\d.]", v):
        return False
    # Long text without formula markers.
    if len(v) > 80 and not re.search(r"[*+\-/^]", v):
        return True
    return False


def split_numbered(text: str) -> list[dict]:
    # Numbered items are line-oriented.  A generic whitespace split also
    # mistakes the final numeric argument in ``IF(x, 220, 200)`` for item
    # ``200)``.  Preserve line boundaries until the individual items are cut.
    raw_text = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    if not norm(raw_text):
        return []
    parts = re.split(r"(?m)(?=^\s*\d+\)\s*)", raw_text)
    out: list[dict] = []
    for part in parts:
        part = norm(part)
        if not part:
            continue
        match = re.match(r"^(\d+)\)\s*(.*)$", part, re.S)
        if match:
            out.append({"number": int(match.group(1)), "formula": match.group(2).strip()})
        else:
            out.append({"number": len(out) + 1, "formula": part})
    return out


def split_criteria_numbered(text: str) -> list[dict]:
    return split_numbered(text)


def factor_key(name: str) -> str:
    # Excel-generated formulas replace spaces and punctuation from the
    # information-sheet label with underscores. Brackets, parentheses,
    # asterisks, slashes, commas and dashes therefore must not affect identity.
    normalized = low(name).replace("№", "n")
    # Telemetry identifiers in brackets (I18134, S22109, P1234) are the most
    # reliable identity when the descriptive part was typed inconsistently.
    codes = re.findall(r"(?<![0-9a-zа-яё])([isp]\d{3,})(?![0-9a-zа-яё])", normalized)
    if codes:
        return codes[-1]
    return re.sub(r"[^0-9a-zа-яё]+", "", normalized)


def merge_factor_names(names: list[str]) -> list[str]:
    result: list[str] = []
    keys: set[str] = set()
    for name in names:
        key = factor_key(name)
        if key and key not in keys:
            keys.add(key)
            result.append(name)
    return result


def renumber_formula_items(items: list) -> tuple[list, dict[int, int]]:
    """Renumber formula items sequentially and return the old→new map."""
    ordered = sorted(items, key=lambda item: item.number)
    mapping: dict[int, int] = {}
    renumbered = []
    for index, item in enumerate(ordered, start=1):
        mapping[item.number] = index
        renumbered.append(type(item)(number=index, raw=item.raw, ast=item.ast, is_computable=item.is_computable))
    return renumbered, mapping


def renumber_criteria_text(text: str, mapping: dict[int, int]) -> str:
    if not text or not mapping:
        return text
    parts = split_numbered(text)
    if not parts:
        return text
    lines = []
    for part in parts:
        old = part["number"]
        new = mapping.get(old, old)
        lines.append(f"{new}) {part['formula']}")
    return "\n".join(lines)


def strip_formula_prefix(text: str) -> str:
    text = norm(text)
    text = re.sub(r"^\s*\d+\)\s*", "", text)
    if text.startswith("="):
        text = text[1:]
    # Operational notes such as "[пл]" are annotations, not expression
    # syntax.  Keep them in FormulaItem.raw for display, but omit them from AST.
    text = re.sub(r"\s*\[[^\[\]]+\]\s*$", "", text)
    # Some source tables carry a one-digit footnote marker after a numeric
    # constant (for example ``105 1``).  It is presentation metadata rather
    # than multiplication and must not make the constant non-computable.
    if re.fullmatch(r"[+-]?\d+(?:[.,]\d+)?\s+\d+", text):
        text = text.rsplit(None, 1)[0]
    # CASE expressions in operational tables sometimes carry a trailing
    # footnote number after the complete expression.
    if re.match(r"^case\s*\(", text, re.I) and re.search(r"\)\s+\d+$", text):
        text = text.rsplit(None, 1)[0]
    # VR/operational tables append a one-digit footnote after a full expression
    # (for example ``320-Pотб 1`` or ``310-Pотб-Pнб2 2``).  Do not treat a
    # trailing arithmetic operand such as ``expr - 1`` as a footnote marker.
    if re.search(r"\s+[1-9]$", text) and not re.search(r"[\+\-\*/]\s+[1-9]$", text):
        text = re.sub(r"\s+[1-9]$", "", text)
    return text.strip()
