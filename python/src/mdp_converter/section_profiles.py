"""Section-specific presentation rules."""
from __future__ import annotations

import re

from .models import Model

DEFAULT_PA_SEASON_LABEL = "Группа стабилизации МДП с ПА"
VOLOGDA_ARKHANGELSK_PA_SEASON_LABEL = "Группа уставок АОПО"
VOLOGDA_ARKHANGELSK_AOPO_SEASONS = {
    "1": "Летняя уставка",
    "2": "Зимняя уставка",
    "3": "Весенне-осенняя уставка",
}


def section_name_from_title(title: str) -> str:
    match = re.search(r"[«\"]([^»\"]+)[»\"]", title)
    return match.group(1).strip() if match else ""


def _normalize_section_name(name: str) -> str:
    text = name.casefold().replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", text).strip()


def is_vologda_arkhangelsk_section(model: Model) -> bool:
    """True for the planned Vologda–Arkhangelsk section (not the VR workbook)."""
    normalized = _normalize_section_name(section_name_from_title(model.title))
    if not normalized or "вр" in normalized.split():
        return False
    return "вологда" in normalized and "архангельск" in normalized


def pa_season_param_label(model: Model) -> str:
    if is_vologda_arkhangelsk_section(model):
        return VOLOGDA_ARKHANGELSK_PA_SEASON_LABEL
    return DEFAULT_PA_SEASON_LABEL


def pa_season_group_label(model: Model, index: int) -> str:
    value = str(index)
    if is_vologda_arkhangelsk_section(model):
        return VOLOGDA_ARKHANGELSK_AOPO_SEASONS.get(value, f"Группа {index}")
    return f"Группа {index}"
