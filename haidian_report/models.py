from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class SourceFile:
    path: Path
    kind: str


@dataclass
class TableSheet:
    title: str
    rows: list[list[str]]
    source: Path
    index: int


@dataclass
class StreetSummary:
    street_name: str
    region: str
    current_month_total: int
    previous_month_total: int
    month_on_month: str
    current_period_total: int
    previous_period_total: int
    year_on_year: str


@dataclass
class StreetReportContext:
    street_name: str
    report_month_label: str
    current_year: int
    current_month: int
    previous_month: int
    previous_year: int
    normal_management: dict[str, Any] = field(default_factory=dict)
    waste_management: dict[str, Any] = field(default_factory=dict)
    city_furniture: dict[str, Any] = field(default_factory=dict)
    backstreets: dict[str, Any] = field(default_factory=dict)
    grid_management: dict[str, Any] = field(default_factory=dict)
    law_enforcement: dict[str, Any] = field(default_factory=dict)
    suggestions: list[dict[str, Any]] = field(default_factory=list)
    report_sections: list[dict[str, Any]] = field(default_factory=list)
