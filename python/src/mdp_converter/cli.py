import argparse
import sys
from pathlib import Path

from .core import convert


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Единый парсер Excel МДП: автономный HTML или исправленный XLSX",
    )
    parser.add_argument("input", help="Исходный файл XLSX или папка с файлами")
    parser.add_argument("-o", "--output", help="Выходной файл или папка")
    parser.add_argument(
        "--format",
        choices=("html", "xlsx"),
        default="html",
        help="Формат результата: html (интерактивный HTML) или xlsx (лист «Ремонтные схемы»)",
    )
    parser.add_argument("--no-calc", action="store_true", help="Только для HTML: без расчётной части")
    parser.add_argument("--no-chart", action="store_true", help="Только для HTML: без графика")
    args = parser.parse_args()

    input_path = Path(args.input)
    if input_path.is_dir():
        from .core import convert_directory

        output_dir = Path(args.output) if args.output else input_path / ("_html" if args.format == "html" else "_корр")
        converted, failures = convert_directory(
            input_path,
            output_dir,
            not args.no_calc,
            not args.no_chart,
            output_format=args.format,
        )
        print(f"Готово: {len(converted)} файлов в {output_dir}")
        for path in converted:
            print(f"  ✓ {path.name}")
        if failures:
            print(f"Ошибок: {len(failures)}", file=sys.stderr)
            for path, error in failures:
                print(f"  ✗ {path.name}: {error}", file=sys.stderr)
            return 1
        return 0

    out = Path(args.output) if args.output else None
    try:
        model = convert(
            input_path,
            out,
            not args.no_calc,
            not args.no_chart,
            output_format=args.format,
        )
        target = out or (
            input_path.with_suffix(".html")
            if args.format == "html"
            else input_path.with_name(f"{input_path.stem}_корр.xlsx")
        )
        print(
            f"Готово: {target}\n"
            f"Формат: {args.format}\n"
            f"Схем: {len(model.schemes)}; факторов: {len(model.factors)}"
        )
    except Exception as exc:
        print("Ошибка:", exc, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
