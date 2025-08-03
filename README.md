# file-converter

Простое графическое приложение для конвертации файлов между форматом
DOCX, XLSX и PDF.

## Требования

- Python 3.8+
- [docx2pdf](https://pypi.org/project/docx2pdf/)
- [xlsx2pdf](https://pypi.org/project/xlsx2pdf/)
- [pdf2docx](https://pypi.org/project/pdf2docx/)

Установить зависимости можно командой:

```bash
pip install -r requirements.txt
```

## Запуск

```bash
python converter.py
```

Программа откроет окно, где можно выбрать исходный файл, указать формат
назначения и место сохранения. Поддерживаются конвертации DOCX → PDF,
XLSX → PDF и PDF → DOCX. Полученный файл автоматически откроется для
проверки.
