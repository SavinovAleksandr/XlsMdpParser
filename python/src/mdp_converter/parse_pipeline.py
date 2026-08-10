"""Main XLSX parsing pipeline."""
from __future__ import annotations

import re
from pathlib import Path

from .diagnostics import ParseDiagnostics
from .expression import parse_expression
from .models import FactorDefinition, FormulaItem, Model, ModeOption, ModeParam, PaVariant, RowData, Scheme
from .normalization import (
    is_empty_or_dash,
    is_non_formula_text,
    low,
    norm,
    renumber_criteria_text,
    renumber_formula_items,
    split_criteria_numbered,
    split_numbered,
    strip_formula_prefix,
)
from .parameter_analysis import analyze_parameters
from .section_profiles import pa_season_group_label, pa_season_param_label
from .table_detector import (
    cell,
    body_start_row,
    detect_columns,
    find_sheet,
    header_map,
    mdp_pa_columns,
    mdp_pa_criteria_columns,
    control_columns,
    row_axis_label as detect_row_axis_label,
)
from .xlsx_reader import read_xlsx


def parse(path: str | Path, diag: ParseDiagnostics | None = None) -> Model:
    diag = diag or ParseDiagnostics()
    sheets = read_xlsx(path)
    _, repair_rows = find_sheet(sheets, "ремонтные схемы")
    if not repair_rows:
        raise ValueError('Не найден лист «Ремонтные схемы»')

    hm = header_map(repair_rows, diag)
    model = Model()
    model.row_axis_label = detect_row_axis_label(repair_rows, hm)
    (
        model.title,
        model.elements,
        model.irregular_oscillation_mw,
        model.weather_stations,
    ) = _read_info_sheet(sheets, diag)

    header_rows = body_start_row(repair_rows, hm)
    model.has_mdp, model.has_mdp_pa, model.has_adp = detect_columns(repair_rows, hm)
    pa_cols = mdp_pa_columns(repair_rows, hm, header_rows)
    pa_criteria_cols = mdp_pa_criteria_columns(repair_rows, hm, header_rows)
    controls = control_columns(
        repair_rows, hm, model.has_mdp, model.has_mdp_pa
    )

    schemes: list[Scheme] = []
    current: Scheme | None = None
    pending_group: RowData | None = None
    stabilization_block = 0
    info_factors: list[str] = []

    for row in repair_rows[header_rows:]:
        num = cell(row, hm, "num")
        name = cell(row, hm, "scheme")
        if num or name:
            if current:
                _commit_group(current, pending_group)
                pending_group = None
                schemes.append(current)
            current = Scheme(num or str(len(schemes) + 1), name or "Без названия")
            stabilization_block = 0

        if not current:
            continue

        row_data = RowData(
            temperature=cell(row, hm, "temp"),
            mdp=cell(row, hm, "mdp") if model.has_mdp else "",
            mdp_pa=_cells(row, pa_cols) if model.has_mdp_pa else "",
            adp=cell(row, hm, "adp") if model.has_adp else "",
            crit_mdp=cell(row, hm, "criteria_mdp"),
            crit_mdp_pa=_cells(row, pa_criteria_cols) if model.has_mdp_pa else "",
            crit_adp=cell(row, hm, "criteria_adp") if model.has_adp else "",
            control_mdp=_unique_cells(row, controls["mdp"]) if model.has_mdp else "",
            control_mdp_pa=_unique_cells(row, controls["mdp_pa"]) if model.has_mdp_pa else "",
            control_adp=_unique_cells(row, controls["adp"]) if model.has_adp else "",
        )
        starts_group = _is_minimum_label(row_data.mdp) or _is_minimum_label(row_data.mdp_pa)
        row_data.mdp = _drop_minimum_label(row_data.mdp)
        row_data.mdp_pa = _drop_minimum_label(row_data.mdp_pa)

        if starts_group:
            _commit_group(current, pending_group)
            pending_group = RowData()
            stabilization_block += 1

        row_data.pa_season = _pa_season_marker(row, stabilization_block)

        if not _row_has_content(row_data):
            note = cell(row, hm, "note")
            status = _special_status(row)
            if status:
                current.is_controlled = False
                note = _join_cell(note, status)
            if note:
                current.note = (current.note + " " + note).strip()
            continue

        row_data.mdp_items = _parse_formula_items(row_data.mdp, diag)
        row_data.mdp_pa_items = _parse_formula_items(row_data.mdp_pa, diag)
        row_data.adp_items = _parse_formula_items(row_data.adp, diag)

        if pending_group is not None:
            if _starts_next_group(pending_group, row_data):
                _commit_group(current, pending_group)
                pending_group = RowData()
            _merge_group_row(pending_group, row_data)
            note = cell(row, hm, "note")
            if note:
                current.note = (current.note + " " + note).strip()
            continue

        # A logical temperature row is commonly spread across several physical
        # Excel rows (one numbered formula/criterion per row).  Continuation
        # rows have no temperature and must be folded into the preceding row;
        # otherwise the HTML shows dozens of sparse rows and cannot choose a
        # minimum across all criteria of that temperature.
        if current.rows and not row_data.temperature and _is_continuation(row_data):
            _append_continuation(current.rows[-1], row_data)
        else:
            current.rows.append(row_data)

        note = cell(row, hm, "note")
        if note:
            current.note = (current.note + " " + note).strip()

    if current:
        _commit_group(current, pending_group)
        schemes.append(current)
    for scheme in schemes:
        _consolidate_pa_stabilization(scheme)
        _consolidate_scheme_controls(scheme)
        _ensure_control_for_ddtn_criteria(scheme, diag)
    model.schemes = schemes

    _, info_rows = find_sheet(sheets, "информация о сечении")
    if not info_rows:
        _, info_rows = find_sheet(sheets, "общая информация")
    if not info_rows and sheets:
        info_rows = next(iter(sheets.values()))
    model.factor_definitions = _read_factors(info_rows)
    info_factors = [factor.name for factor in model.factor_definitions]

    ast_nodes = []
    for scheme in model.schemes:
        for row in scheme.rows:
            for item in row.mdp_items + row.mdp_pa_items + row.adp_items:
                if item.ast:
                    from .expression.evaluator import _from_dict

                    try:
                        ast_nodes.append(_from_dict(item.ast))
                    except Exception:
                        pass

    model.mode_params, model.factors = analyze_parameters(ast_nodes, info_factors)
    _apply_pa_season_modes(model)
    return model


def _read_info_sheet(
    sheets: dict, diag: ParseDiagnostics | None = None
) -> tuple[str, list[str], float | None, list[str]]:
    title = "Контролируемое сечение"
    elements: list[str] = []
    irregular_oscillation_mw: float | None = None
    weather_stations: list[str] = []
    in_composition = False
    in_weather = False
    _, info = find_sheet(sheets, "информация о сечении")
    if not info:
        _, info = find_sheet(sheets, "общая информация")
    for row in info:
        vals = list(dict.fromkeys(norm(x) for x in row if norm(x)))
        if not vals:
            continue
        text = " ".join(vals)
        text_low = low(text)
        if title == "Контролируемое сечение" and "допустимые перетоки" in low(text):
            title = text
        if "значение амплитуды" in text_low or "нерегулярн" in text_low:
            match = re.search(r"(-?\d+(?:[.,]\d+)?)\s*(?:мвт)?", text, re.IGNORECASE)
            if match:
                irregular_oscillation_mw = float(match.group(1).replace(",", "."))
            elif diag:
                diag.warn(f"Не разобрана величина нерегулярных колебаний: {text}")
        if "метеостанци" in text_low:
            in_weather = True
            tail = re.split(r"метеостанци(?:я|и|й)?\s*:", text, maxsplit=1, flags=re.IGNORECASE)
            if len(tail) == 2 and tail[1].strip():
                weather_stations.append(tail[1].strip())
            continue
        if in_weather and any(
            marker in text_low
            for marker in (
                "влияющие факторы",
                "группы ремонтов",
                "сопроводительная информация",
                "ремонтные схемы",
                "примечани",
            )
        ):
            in_weather = False
        elif in_weather:
            weather_stations.append(vals[-1])
            continue
        if "состав контролируемого сечения" in text_low:
            in_composition = True
            continue
        if in_composition and any(
            marker in text_low
            for marker in ("сечение не контролируется", "условия для снятия", "значение амплитуды", "метеостанции")
        ):
            in_composition = False
        if in_composition:
            # The composition may contain lines, transformers and other
            # equipment, not only values starting with "ВЛ".
            elements.append(vals[-1])
    return (
        title,
        list(dict.fromkeys(elements)),
        irregular_oscillation_mw,
        list(dict.fromkeys(weather_stations)),
    )


def _read_factors(info_rows: list[list[str]]) -> list[FactorDefinition]:
    factors: list[FactorDefinition] = []
    in_section = False
    for row in info_rows:
        vals = [norm(x) for x in row if norm(x)]
        if not vals:
            continue
        text = " ".join(vals)
        text_low = low(text)
        if "влияющие факторы" in text_low:
            in_section = True
            continue
        if in_section and any(
            marker in text_low
            for marker in (
                "группы ремонтов",
                "сопроводительная информация",
                "ремонтные схемы",
            )
        ):
            break
        if not in_section:
            continue
        parsed = False
        for cell in reversed(vals):
            if ":" not in cell:
                continue
            name, _, description = cell.partition(":")
            name = name.strip()
            if name:
                factors.append(FactorDefinition(name=name, description=description.strip()))
                parsed = True
                break
        if parsed:
            continue
        if len(vals) >= 2:
            name = vals[0].rstrip(":").strip()
            description = vals[-1].strip()
            if name and description and name != description:
                factors.append(FactorDefinition(name=name, description=description))
    return factors


def _row_has_content(row: RowData) -> bool:
    fields = [
        row.temperature, row.mdp, row.mdp_pa, row.adp,
        row.crit_mdp, row.crit_mdp_pa, row.crit_adp,
        row.control_mdp, row.control_mdp_pa, row.control_adp,
    ]
    meaningful = [f for f in fields if f and not low(f).startswith("минимальное из")]
    return bool(meaningful)


def _special_status(row: list[str]) -> str:
    for value in row:
        text = norm(value)
        if "не контролируется" in low(text):
            return text
    return ""


def _is_continuation(row: RowData) -> bool:
    """Whether a physical row continues a numbered logical data row."""
    return bool(
        row.mdp_items
        or row.mdp_pa_items
        or row.adp
        or row.crit_mdp
        or row.crit_mdp_pa
        or row.crit_adp
        or row.control_mdp
        or row.control_mdp_pa
        or row.control_adp
    )


def _drop_minimum_label(value: str) -> str:
    return "" if _is_minimum_label(value) else value


def _is_minimum_label(value: str) -> bool:
    return low(value).startswith("минимальное из")


def _join_cell(left: str, right: str) -> str:
    if not right:
        return left
    if not left:
        return right
    return f"{left}\n{right}"


def _cells(row: list[str], columns: list[int]) -> str:
    value = ""
    for col in columns:
        if col < len(row):
            value = _join_cell(value, norm(row[col]))
    return value


def _unique_cells(row: list[str], columns: list[int]) -> str:
    values: list[str] = []
    for col in columns:
        value = norm(row[col]) if col < len(row) else ""
        if value and value not in values:
            values.append(value)
    return "\n".join(values)


def _append_continuation(target: RowData, source: RowData) -> None:
    target.mdp = _join_cell(target.mdp, source.mdp)
    target.mdp_pa = _join_cell(target.mdp_pa, source.mdp_pa)
    target.adp = _join_cell(target.adp, source.adp)
    target.crit_mdp = _join_cell(target.crit_mdp, source.crit_mdp)
    target.crit_mdp_pa = _join_cell(target.crit_mdp_pa, source.crit_mdp_pa)
    target.crit_adp = _join_cell(target.crit_adp, source.crit_adp)
    target.control_mdp = _join_cell(target.control_mdp, source.control_mdp)
    target.control_mdp_pa = _join_cell(target.control_mdp_pa, source.control_mdp_pa)
    target.control_adp = _join_cell(target.control_adp, source.control_adp)
    target.mdp_items.extend(source.mdp_items)
    target.mdp_pa_items.extend(source.mdp_pa_items)
    target.adp_items.extend(source.adp_items)


def _merge_group_row(target: RowData, source: RowData) -> None:
    if source.temperature and not target.temperature:
        target.temperature = source.temperature
    _append_continuation(target, source)


def _starts_next_group(target: RowData, source: RowData) -> bool:
    if source.temperature and target.temperature and source.temperature != target.temperature:
        return True
    for existing, incoming in (
        (target.mdp_items, source.mdp_items),
        (target.mdp_pa_items, source.mdp_pa_items),
    ):
        if any(item.number == 1 for item in existing) and any(item.number == 1 for item in incoming):
            return True
    return False


def _parse_formula_items(text: str, diag: ParseDiagnostics) -> list[FormulaItem]:
    if is_empty_or_dash(text):
        return []
    items: list[FormulaItem] = []
    for part in split_numbered(text):
        raw = part["formula"]
        if is_non_formula_text(raw):
            items.append(
                FormulaItem(number=part["number"], raw=raw, ast=None, is_computable=False)
            )
            continue
        expr = strip_formula_prefix(raw)
        try:
            ast = parse_expression(expr)
            items.append(
                FormulaItem(number=part["number"], raw=raw, ast=ast.to_dict(), is_computable=True)
            )
        except Exception as exc:
            diag.warn(f"Формула не разобрана ({part['number']}): {raw[:80]} — {exc}")
            items.append(
                FormulaItem(number=part["number"], raw=raw, ast=None, is_computable=False)
            )
    if not items and norm(text) and not is_non_formula_text(text):
        expr = strip_formula_prefix(text)
        try:
            ast = parse_expression(expr)
            items.append(FormulaItem(number=1, raw=text, ast=ast.to_dict(), is_computable=True))
        except Exception:
            items.append(FormulaItem(number=1, raw=text, ast=None, is_computable=False))
    return items


def _commit_group(scheme: Scheme, group: RowData | None) -> None:
    if group is not None and _row_has_content(group):
        scheme.rows.append(group)


def _pa_season_marker(row: list[str], block_index: int) -> str:
    markers: list[str] = []
    for col in (3, 4):
        if col < len(row):
            value = norm(row[col])
            if value and value not in DASHES and not low(value).startswith("минимальное"):
                markers.append(value)
    if markers:
        return " / ".join(markers)
    return str(max(block_index, 1))


def _consolidate_scheme_controls(scheme: Scheme) -> None:
    """Hoist scheme-wide control parameters from merged Excel cells onto the first row."""
    if not scheme.rows:
        return
    for attr in ("control_mdp", "control_mdp_pa", "control_adp"):
        parts: list[str] = []
        for row in scheme.rows:
            text = getattr(row, attr)
            if not text:
                continue
            for line in text.split("\n"):
                line = line.strip()
                if line and line not in parts:
                    parts.append(line)
        if len(parts) <= 1:
            merged = "\n".join(parts)
            for row in scheme.rows:
                setattr(row, attr, "")
            if merged:
                setattr(scheme.rows[0], attr, merged)


def _criteria_has_ddtn(text: str) -> bool:
    normalized = low(text)
    if "ддтн" in normalized or "длительно допустим" in normalized:
        return True
    for part in split_criteria_numbered(text):
        formula = low(part.get("formula", ""))
        if "ддтн" in formula or "длительно допустим" in formula:
            return True
    return False


def _extract_ddtn_control_text(criteria: str) -> str:
    lines: list[str] = []
    for part in split_criteria_numbered(criteria):
        formula = norm(part.get("formula", ""))
        formula_low = low(formula)
        if formula and ("ддтн" in formula_low or "длительно допустим" in formula_low):
            if formula not in lines:
                lines.append(formula)
    return "\n".join(lines)


def _scheme_control_value(rows: list[RowData], control_attr: str) -> str:
    for row in rows:
        value = norm(getattr(row, control_attr, ""))
        if value:
            return getattr(row, control_attr, "")
    return ""


def _ensure_control_for_ddtn_criteria(scheme: Scheme, diag: ParseDiagnostics | None = None) -> None:
    """Ensure control exists when DDTN criteria are present, without duplicating per TNV row."""
    pairs = (
        ("crit_mdp", "control_mdp", "МДП без ПА"),
        ("crit_mdp_pa", "control_mdp_pa", "МДП с ПА"),
        ("crit_adp", "control_adp", "АДП"),
    )
    for crit_attr, control_attr, label in pairs:
        if _scheme_control_value(scheme.rows, control_attr):
            continue

        inferred = ""
        missing_rows: list[RowData] = []
        for row in scheme.rows:
            criteria = getattr(row, crit_attr, "")
            if not _criteria_has_ddtn(criteria):
                continue
            missing_rows.append(row)
            if not inferred:
                inferred = _extract_ddtn_control_text(criteria)

        if not missing_rows:
            continue

        if inferred:
            setattr(scheme.rows[0], control_attr, inferred)
            continue

        if diag:
            for row in missing_rows:
                diag.warn(
                    f"Схема {scheme.number} «{scheme.name}», "
                    f"ТНВ {row.temperature or '—'}: в критериях {label} указан ДДТН, "
                    "но контроль доп. параметров не задан."
                )


def _consolidate_pa_stabilization(scheme: Scheme) -> None:
    """Fold duplicate temperature rows from multiple PA blocks into one logical row."""
    if not scheme.rows:
        return
    grouped: dict[str, list[RowData]] = {}
    order: list[str] = []
    for index, row in enumerate(scheme.rows):
        key = row.temperature or f"__row_{index}"
        if key not in grouped:
            order.append(key)
            grouped[key] = []
        grouped[key].append(row)

    consolidated: list[RowData] = []
    for key in order:
        rows = grouped[key]
        if len(rows) == 1:
            consolidated.append(_finalize_pa_row(rows[0]))
            continue
        consolidated.append(_merge_pa_temperature_rows(rows))
    scheme.rows = consolidated


def _finalize_pa_row(row: RowData) -> RowData:
    if not row.mdp_pa_items:
        return row
    items, mapping = renumber_formula_items(row.mdp_pa_items)
    row.mdp_pa_items = items
    row.crit_mdp_pa = renumber_criteria_text(row.crit_mdp_pa, mapping)
    row.mdp_pa = _join_numbered_items(items)
    return row


def _merge_pa_temperature_rows(rows: list[RowData]) -> RowData:
    base = rows[0]
    merged = RowData(
        temperature=base.temperature,
        mdp=base.mdp,
        adp=base.adp,
        crit_mdp=base.crit_mdp,
        crit_adp=base.crit_adp,
        control_mdp=base.control_mdp,
        control_adp=base.control_adp,
        mdp_items=list(base.mdp_items),
        adp_items=list(base.adp_items),
    )
    for index, row in enumerate(rows, start=1):
        if not row.mdp_pa_items and not row.mdp_pa:
            continue
        season = row.pa_season or str(index)
        label = f"Стабилизация {season}" if season.isdigit() else f"Стабилизация ({season})"
        items, mapping = renumber_formula_items(row.mdp_pa_items)
        crit = renumber_criteria_text(row.crit_mdp_pa, mapping)
        merged.mdp_pa_variants.append(
            PaVariant(
                season=str(index),
                label=label,
                mdp_pa=_join_numbered_items(items),
                crit_mdp_pa=crit,
                mdp_pa_items=items,
            )
        )
        if row.mdp_items and len(row.mdp_items) > len(merged.mdp_items):
            merged.mdp_items = list(row.mdp_items)
            merged.mdp = row.mdp
            merged.crit_mdp = row.crit_mdp
        if row.control_mdp_pa:
            merged.control_mdp_pa = _join_cell(merged.control_mdp_pa, row.control_mdp_pa)
        if row.control_mdp and not merged.control_mdp:
            merged.control_mdp = row.control_mdp

    if merged.mdp_pa_variants:
        first = merged.mdp_pa_variants[0]
        merged.mdp_pa_items = list(first.mdp_pa_items)
        merged.mdp_pa = first.mdp_pa
        merged.crit_mdp_pa = first.crit_mdp_pa
    return merged


def _join_numbered_items(items: list[FormulaItem]) -> str:
    return "\n".join(f"{item.number}) {item.raw}" for item in items if item.raw)


def _apply_pa_season_modes(model: Model) -> None:
    max_variants = 0
    for scheme in model.schemes:
        for row in scheme.rows:
            if row.mdp_pa_variants:
                max_variants = max(max_variants, len(row.mdp_pa_variants))
    if max_variants < 2:
        return
    model.has_pa_seasons = True
    model.pa_season_label = pa_season_param_label(model)
    model.pa_season_options = [
        ModeOption(value=str(index), label=pa_season_group_label(model, index))
        for index in range(1, max_variants + 1)
    ]
    _apply_pa_variant_labels(model)
    if not any(p.name == "pa_season" for p in model.mode_params):
        model.mode_params.insert(
            0,
            ModeParam(
                name="pa_season",
                kind="select",
                options=list(model.pa_season_options),
                default="1",
            ),
        )


def _apply_pa_variant_labels(model: Model) -> None:
    for scheme in model.schemes:
        for row in scheme.rows:
            for variant in row.mdp_pa_variants:
                try:
                    index = int(variant.season)
                except ValueError:
                    index = 0
                if index:
                    variant.label = pa_season_group_label(model, index)


DASHES = frozenset({"-", "—", "–", "−"})
