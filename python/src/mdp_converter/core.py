"""Public API for unified MDP parsing and export."""
from __future__ import annotations

from pathlib import Path
from typing import Callable, Literal

from .diagnostics import ParseDiagnostics
from .html_generator import generate as generate_html
from .models import Model
from .parse_pipeline import parse
from .xlsx_generator import default_xlsx_output_path, generate_xlsx
from .xlsx_reader import read_xlsx

OutputFormat = Literal["html", "xlsx"]

__all__ = [
    "convert",
    "convert_directory",
    "parse",
    "generate_html",
    "generate_xlsx",
    "read_xlsx",
    "ParseDiagnostics",
    "OutputFormat",
]


def _resolve_output_path(
    input_path: Path,
    output_path: Path | None,
    output_format: OutputFormat,
) -> Path:
    if output_path is None:
        return default_xlsx_output_path(input_path) if output_format == "xlsx" else input_path.with_suffix(".html")
    output_path = Path(output_path)
    if output_path.is_dir():
        if output_format == "xlsx":
            return output_path / default_xlsx_output_path(input_path).name
        return output_path / f"{input_path.stem}.html"
    return output_path


def convert(
    input_path: str | Path,
    output_path: str | Path | None = None,
    include_calculation: bool = True,
    include_chart: bool = True,
    diagnostics: ParseDiagnostics | None = None,
    *,
    output_format: OutputFormat = "html",
) -> Model:
    diag = diagnostics or ParseDiagnostics()
    source = Path(input_path)
    target = _resolve_output_path(source, Path(output_path) if output_path is not None else None, output_format)
    model = parse(source, diag)

    if output_format == "html":
        generate_html(model, target, include_calculation, include_chart)
    elif output_format == "xlsx":
        generate_xlsx(model, source, target)
    else:
        raise ValueError(f"Неизвестный формат вывода: {output_format}")

    diag_path = target.with_suffix(".diagnostics.json")
    if diag.warnings or diag.errors:
        diag.save(diag_path)
    elif diag_path.exists():
        diag_path.unlink()
    return model


def convert_directory(
    input_directory: str | Path,
    output_directory: str | Path,
    include_calculation: bool = True,
    include_chart: bool = True,
    progress_callback: Callable[[int, int, Path], None] | None = None,
    *,
    output_format: OutputFormat = "html",
) -> tuple[list[Path], list[tuple[Path, str]]]:
    """Convert every XLSX in a directory, continuing after individual failures."""
    source_dir = Path(input_directory)
    output_dir = Path(output_directory)
    if not source_dir.is_dir():
        raise ValueError(f"Папка с исходными файлами не найдена: {source_dir}")

    sources = sorted(
        (
            path
            for path in source_dir.iterdir()
            if path.is_file()
            and path.suffix.lower() == ".xlsx"
            and not path.name.startswith("~$")
            and not path.stem.endswith("_корр")
        ),
        key=lambda path: path.name.casefold(),
    )
    if not sources:
        raise ValueError("В выбранной папке нет файлов XLSX.")

    output_dir.mkdir(parents=True, exist_ok=True)
    converted: list[Path] = []
    failures: list[tuple[Path, str]] = []
    total = len(sources)
    for index, source in enumerate(sources, start=1):
        if progress_callback:
            progress_callback(index, total, source)
        target = _resolve_output_path(source, output_dir, output_format)
        try:
            convert(
                source,
                target,
                include_calculation,
                include_chart,
                diagnostics=ParseDiagnostics(),
                output_format=output_format,
            )
            converted.append(target)
        except Exception as exc:
            failures.append((source, str(exc)))
    return converted, failures
