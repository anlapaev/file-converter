import os
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    from docx2pdf import convert
except ImportError:
    convert = None


def select_file():
    """Выбор исходного файла DOCX."""
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if file_path:
        entry_input.delete(0, tk.END)
        entry_input.insert(0, file_path)


def open_file(path: str) -> None:
    """Открыть файл стандартным приложением."""
    if sys.platform.startswith("darwin"):
        subprocess.call(["open", path])
    elif os.name == "nt":
        os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.call(["xdg-open", path])


def convert_file():
    """Конвертация DOCX → PDF и открытие PDF."""
    if convert is None:
        messagebox.showerror("Ошибка", "Библиотека docx2pdf не установлена.")
        return

    input_path = entry_input.get().strip()
    if not input_path:
        messagebox.showerror("Ошибка", "Выберите DOCX-файл.")
        return

    output_path = os.path.splitext(input_path)[0] + ".pdf"
    try:
        convert(input_path, output_path)
        messagebox.showinfo("Готово", f"Файл сохранён:\n{output_path}")
        open_file(output_path)
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Ошибка", f"Не удалось конвертировать:\n{exc}")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("DOCX → PDF Converter")

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack()

    entry_input = tk.Entry(frame, width=50)
    entry_input.grid(row=0, column=0, padx=(0, 5))
    btn_browse = tk.Button(frame, text="Выбрать файл", command=select_file)
    btn_browse.grid(row=0, column=1, sticky="ew")

    btn_convert = tk.Button(root, text="Конвертировать", command=convert_file, padx=20, pady=5)
    btn_convert.pack(pady=(10, 0))

    root.mainloop()
