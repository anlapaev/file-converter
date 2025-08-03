"""Простое GUI-приложение для конвертации файлов.

Поддерживаются форматы DOCX, XLSX и PDF. Программа позволяет выбрать
исходный файл, формат назначения и папку для сохранения результата.
Для некоторых пар форматов используется сторонние библиотеки:

- ``docx2pdf`` — DOCX → PDF
- ``xlsx2pdf`` — XLSX → PDF
- ``pdf2docx`` — PDF → DOCX

Остальные сочетания форматов в данный момент не поддерживаются.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

try:  # DOCX → PDF
    from docx2pdf import convert as docx2pdf_convert
except Exception:  # noqa: BLE001
    docx2pdf_convert = None

try:  # XLSX → PDF
    from xlsx2pdf import xlsx2pdf as xlsx2pdf_convert
except Exception:  # noqa: BLE001
    xlsx2pdf_convert = None

try:  # PDF → DOCX
    from pdf2docx import Converter as PDF2DocxConverter
except Exception:  # noqa: BLE001
    PDF2DocxConverter = None


FORMATS = ["docx", "xlsx", "pdf"]


def open_file(path: str) -> None:
    """Открыть файл стандартным приложением."""

    try:
        if sys.platform.startswith("darwin"):
            subprocess.call(["open", path])
        elif os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.call(["xdg-open", path])
    except Exception:  # noqa: BLE001
        pass


def browse_input() -> None:
    """Выбор исходного файла."""

    ext = input_var.get()
    filetypes = [(f"{ext.upper()} files", f"*.{ext}")]
    file_path = filedialog.askopenfilename(filetypes=filetypes)
    if file_path:
        entry_input.delete(0, tk.END)
        entry_input.insert(0, file_path)


def browse_output() -> None:
    """Выбор файла для сохранения результата."""

    ext = output_var.get()
    filetypes = [(f"{ext.upper()} files", f"*.{ext}")]
    file_path = filedialog.asksaveasfilename(
        defaultextension=f".{ext}", filetypes=filetypes
    )
    if file_path:
        entry_output.delete(0, tk.END)
        entry_output.insert(0, file_path)


def convert_file() -> None:
    """Запустить конвертацию по выбранным параметрам."""

    src = entry_input.get().strip()
    dst = entry_output.get().strip()
    in_ext = input_var.get()
    out_ext = output_var.get()

    if not src:
        messagebox.showerror("Ошибка", "Выберите исходный файл.")
        return
    if not dst:
        messagebox.showerror("Ошибка", "Укажите путь сохранения.")
        return

    try:
        if in_ext == out_ext:
            shutil.copyfile(src, dst)
        elif in_ext == "docx" and out_ext == "pdf":
            if docx2pdf_convert is None:
                raise RuntimeError("Библиотека docx2pdf не установлена")
            docx2pdf_convert(src, dst)
        elif in_ext == "xlsx" and out_ext == "pdf":
            if xlsx2pdf_convert is None:
                raise RuntimeError("Библиотека xlsx2pdf не установлена")
            xlsx2pdf_convert(src, dst)
        elif in_ext == "pdf" and out_ext == "docx":
            if PDF2DocxConverter is None:
                raise RuntimeError("Библиотека pdf2docx не установлена")
            with PDF2DocxConverter(src) as cv:
                cv.convert(dst)
        else:
            messagebox.showerror(
                "Ошибка",
                f"Конвертация {in_ext} → {out_ext} не поддерживается.",
            )
            return

        messagebox.showinfo("Готово", f"Файл сохранён:\n{dst}")
        open_file(dst)
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Ошибка", f"Не удалось конвертировать:\n{exc}")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("File Converter")

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack()

    # Входной файл и формат
    input_var = tk.StringVar(value=FORMATS[0])
    option_in = tk.OptionMenu(frame, input_var, *FORMATS)
    option_in.grid(row=0, column=0, padx=(0, 5))

    entry_input = tk.Entry(frame, width=40)
    entry_input.grid(row=0, column=1, padx=(0, 5))
    btn_browse_in = tk.Button(frame, text="Файл", command=browse_input)
    btn_browse_in.grid(row=0, column=2)

    # Файл назначения и формат
    output_var = tk.StringVar(value=FORMATS[2])
    option_out = tk.OptionMenu(frame, output_var, *FORMATS)
    option_out.grid(row=1, column=0, padx=(0, 5))

    entry_output = tk.Entry(frame, width=40)
    entry_output.grid(row=1, column=1, padx=(0, 5))
    btn_browse_out = tk.Button(frame, text="Сохранить как", command=browse_output)
    btn_browse_out.grid(row=1, column=2)

    btn_convert = tk.Button(root, text="Конвертировать", command=convert_file, padx=20, pady=5)
    btn_convert.pack(pady=(10, 0))

    root.mainloop()