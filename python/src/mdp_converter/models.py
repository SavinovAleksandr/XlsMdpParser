"""Data models for MDP converter."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class FormulaItem:
    number: int
    raw: str
    ast: dict[str, Any] | None = None
    is_computable: bool = True


@dataclass
class PaVariant:
    """One MDP-with-PA stabilization block (season / AOPO device group)."""
    season: str
    label: str
    mdp_pa: str = ""
    crit_mdp_pa: str = ""
    mdp_pa_items: list[FormulaItem] = field(default_factory=list)


@dataclass
class RowData:
    temperature: str = ""
    mdp: str = ""
    mdp_pa: str = ""
    adp: str = ""
    crit_mdp: str = ""
    crit_mdp_pa: str = ""
    crit_adp: str = ""
    control_mdp: str = ""
    control_mdp_pa: str = ""
    control_adp: str = ""
    mdp_items: list[FormulaItem] = field(default_factory=list)
    mdp_pa_items: list[FormulaItem] = field(default_factory=list)
    adp_items: list[FormulaItem] = field(default_factory=list)
    pa_season: str = ""
    mdp_pa_variants: list[PaVariant] = field(default_factory=list)


@dataclass
class Scheme:
    number: str
    name: str
    rows: list[RowData] = field(default_factory=list)
    note: str = ""
    is_controlled: bool = True


@dataclass
class ModeOption:
    value: str
    label: str


@dataclass
class ModeParam:
    name: str
    kind: str  # bool | select | number
    options: list[ModeOption] = field(default_factory=list)
    default: str = "0"


@dataclass
class Factor:
    name: str
    default: float = 0.0


@dataclass
class FactorDefinition:
    name: str
    description: str = ""


@dataclass
class Model:
    title: str = "Контролируемое сечение"
    elements: list[str] = field(default_factory=list)
    irregular_oscillation_mw: float | None = None
    weather_stations: list[str] = field(default_factory=list)
    factors: list[Factor] = field(default_factory=list)
    factor_definitions: list[FactorDefinition] = field(default_factory=list)
    mode_params: list[ModeParam] = field(default_factory=list)
    schemes: list[Scheme] = field(default_factory=list)
    has_mdp: bool = False
    has_mdp_pa: bool = False
    has_adp: bool = False
    has_pa_seasons: bool = False
    pa_season_options: list[ModeOption] = field(default_factory=list)
    pa_season_label: str = "Группа стабилизации МДП с ПА"
    row_axis_label: str = "ТНВ"
