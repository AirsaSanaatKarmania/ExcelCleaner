# Excel Cleaner

[![Tests](https://github.com/AirsaSanaatKarmania/ExcelCleaner/actions/workflows/tests.yml/badge.svg)](https://github.com/AirsaSanaatKarmania/ExcelCleaner/actions/workflows/tests.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

[فارسی](README.md)

A small Windows (Tkinter) tool for cleaning Excel data that contains Persian/Arabic digits and ID-like numbers such as card numbers, personnel codes and national IDs, without corrupting them. It can also strip all formatting and formulas down to raw values.

## Features

- Persian/Arabic digits to ASCII: `۱۲۳` -> `123`
- Numeric-looking text with separators to a plain string: `5892-1041-0085-9016` -> `5892104100859016`
- Keeps leading zeros and 16-digit card numbers (stored as text)
- Choose cells your way: drag, Ctrl+click, click a column/row header, or type an address like `B2:B500, D:D, 5:8`
- **Raw mode** for the selection, the current sheet or all sheets: formulas to values, clear cell styles and number formats, unmerge cells, remove hyperlinks/comments/conditional formatting/data validation/tables/filters, trim whitespace
- The original file is never modified; output is saved as `*_fixed.xlsx`

## Download

Grab `ExcelCleaner.exe` from the [latest release](../../releases/latest).

## Run from source

Python 3.9+ with tcl/tk.

```bash
pip install -r requirements.txt
python excel_cleaner.py
```

## Build the exe

Run `build_exe.bat` on Windows (output: `dist\ExcelCleaner.exe`), or push a tag such as `v1.0.0` and let GitHub Actions build and attach it to a release.

## Why do card numbers break in Excel?

Excel's Number type keeps only 15 significant digits, so a 16-digit card number becomes `5.8921E+15` and the last digit turns into `0` permanently. Card numbers are identifiers, not numbers: store them as Text. This tool defaults to text output and warns about numeric cells that were probably corrupted already.

## Limitations

- `.xlsx` / `.xlsm` only.
- Files are saved with openpyxl, so charts, images and pivot tables may not survive. Test on a copy first.
- Formula values come from the cached results stored by Excel; formulas without a cached value are left untouched (a warning is shown).

## Development

```bash
pip install -r requirements.txt pytest
pytest -q
```

## License

[MIT](LICENSE)
