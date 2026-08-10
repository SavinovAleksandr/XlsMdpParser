import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

from .core import convert, convert_directory
from .diagnostics import ParseDiagnostics
from .xlsx_generator import default_xlsx_output_path


def main() -> None:
    root = tk.Tk()
    root.title("МДП Excel → HTML / XLSX")
    root.geometry("780x560")

    source = tk.StringVar()
    source_mode = tk.StringVar(value="file")
    output_format = tk.StringVar(value="html")
    calc = tk.BooleanVar(value=True)
    chart = tk.BooleanVar(value=True)
    progress = tk.StringVar()

    tk.Label(root, text="Единый парсер МДП", font=("Segoe UI", 12, "bold")).pack(pady=(18, 5))
    tk.Label(
        root,
        text="Один и тот же входной XLSX можно сохранить как интерактивный HTML или как исправленный XLSX.",
        fg="#64748b",
        wraplength=700,
    ).pack(padx=20)

    mode_row = tk.Frame(root)
    mode_row.pack(fill="x", padx=20, pady=(12, 7))
    tk.Radiobutton(mode_row, text="Один файл XLSX", variable=source_mode, value="file").pack(side="left")
    tk.Radiobutton(
        mode_row,
        text="Все XLSX из папки",
        variable=source_mode,
        value="folder",
    ).pack(side="left", padx=(20, 0))

    row = tk.Frame(root)
    row.pack(fill="x", padx=20)
    tk.Entry(row, textvariable=source).pack(side="left", fill="x", expand=True)
    select_button = tk.Button(row, text="Выбрать")
    select_button.pack(side="left", padx=8)

    def choose_source() -> None:
        if source_mode.get() == "folder":
            selected = filedialog.askdirectory(title="Выберите папку с файлами XLSX")
        else:
            selected = filedialog.askopenfilename(filetypes=[("Excel XLSX", "*.xlsx")])
        if selected:
            source.set(selected)

    select_button.configure(command=choose_source)

    def sync_source_mode(*_) -> None:
        source.set("")
        select_button.configure(
            text="Выбрать папку" if source_mode.get() == "folder" else "Выбрать файл"
        )

    source_mode.trace_add("write", sync_source_mode)
    sync_source_mode()

    format_frame = tk.LabelFrame(root, text="Формат результата", font=("Segoe UI", 10, "bold"))
    format_frame.pack(fill="x", padx=20, pady=16)
    tk.Radiobutton(
        format_frame,
        text="HTML — автономный интерактивный файл с расчётом и копированием МДП",
        variable=output_format,
        value="html",
        anchor="w",
        justify="left",
        wraplength=650,
    ).pack(fill="x", padx=12, pady=(10, 4))
    tk.Radiobutton(
        format_frame,
        text="XLSX — исправленный файл *_корр.xlsx с листом «Ремонтные схемы»",
        variable=output_format,
        value="xlsx",
        anchor="w",
        justify="left",
        wraplength=650,
    ).pack(fill="x", padx=12, pady=(4, 10))

    opts = tk.LabelFrame(root, text="Опции HTML", font=("Segoe UI", 10, "bold"))
    opts.pack(fill="x", padx=20, pady=(0, 12))

    calc_cb = tk.Checkbutton(
        opts,
        text="Добавить расчётную часть, влияющие факторы и выбор минимального МДП",
        variable=calc,
        anchor="w",
        justify="left",
        wraplength=650,
    )
    calc_cb.pack(fill="x", padx=12, pady=(10, 5))

    chart_cb = tk.Checkbutton(
        opts,
        text="Добавить график зависимости МДП от выбранного влияющего фактора",
        variable=chart,
        anchor="w",
        justify="left",
        wraplength=650,
    )
    chart_cb.pack(fill="x", padx=12, pady=(5, 10))

    note = tk.Label(
        opts,
        text="Опции HTML игнорируются при выборе формата XLSX.",
        fg="#64748b",
        anchor="w",
    )
    note.pack(fill="x", padx=14, pady=(0, 10))

    def sync_options(*_) -> None:
        html_mode = output_format.get() == "html"
        state = "normal" if html_mode else "disabled"
        calc_cb.configure(state=state)
        chart_cb.configure(state=state if calc.get() and html_mode else "disabled")
        if not html_mode:
            chart.set(False)

    def sync_chart(*_) -> None:
        if output_format.get() == "html" and calc.get():
            chart_cb.configure(state="normal")
        else:
            chart_cb.configure(state="disabled")

    output_format.trace_add("write", sync_options)
    calc.trace_add("write", sync_chart)
    sync_options()
    sync_chart()

    def run() -> None:
        if not source.get():
            target = "папку с файлами XLSX" if source_mode.get() == "folder" else "исходный XLSX"
            messagebox.showwarning("Исходные данные", f"Выберите {target}.")
            return

        fmt = output_format.get()
        if source_mode.get() == "folder":
            if fmt == "html":
                out = filedialog.askdirectory(title="Выберите папку для готовых HTML")
            else:
                out = filedialog.askdirectory(title="Выберите папку для файлов *_корр.xlsx")
        elif fmt == "html":
            out = filedialog.asksaveasfilename(
                defaultextension=".html",
                initialfile=Path(source.get()).stem + ".html",
                filetypes=[("HTML", "*.html")],
            )
        else:
            default_name = default_xlsx_output_path(source.get()).name
            out = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel XLSX", "*.xlsx")],
            )
        if not out:
            return

        try:
            progress.set("Обработка файлов…")
            root.update_idletasks()
            if source_mode.get() == "folder":
                def update_progress(index, total, path):
                    progress.set(f"Обработка {index} из {total}: {path.name}")
                    root.update_idletasks()

        converted, failures = convert_directory(
                    source.get(),
                    out,
                    calc.get(),
                    chart.get(),
                    progress_callback=update_progress,
                    output_format=fmt,
                )
                details = ""
                if failures:
                    preview = "\n".join(f"• {path.name}: {error}" for path, error in failures[:5])
                    details = f"\n\nОшибок: {len(failures)}\n{preview}" + ("…" if len(failures) > 5 else "")
                label = "HTML" if fmt == "html" else "XLSX"
                messagebox.showinfo(
                    "Пакетная обработка завершена",
                    f"Создано {label}-файлов: {len(converted)}\nПапка результатов:\n{out}{details}",
                )
                return

            diag = ParseDiagnostics()
            model = convert(
                source.get(),
                out,
                calc.get(),
                chart.get(),
                diagnostics=diag,
                output_format=fmt,
            )
            extra = ""
            if diag.warnings:
                extra = f"\n\nПредупреждения: {len(diag.warnings)} (см. *.diagnostics.json)"
            label = "HTML" if fmt == "html" else "XLSX"
            messagebox.showinfo(
                "Готово",
                f"Создан {label}-файл:\n{out}\n\n"
                f"Ремонтных схем: {len(model.schemes)}\n"
                f"Режимных параметров: {len(model.mode_params)}\n"
                f"Влияющих факторов: {len(model.factors)}{extra}",
            )
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))
        finally:
            progress.set("")

    tk.Button(
        root,
        text="Сформировать результат",
        font=("Segoe UI", 11, "bold"),
        height=2,
        command=run,
    ).pack(pady=(14, 5))
    tk.Label(root, textvariable=progress, fg="#1f5bb5").pack()

    root.mainloop()


if __name__ == "__main__":
    main()
