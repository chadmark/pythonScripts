# Publisher to PDF Batch Converter

## Overview

`pub_to_pdf_batch.py` is a Python script designed to recursively locate Microsoft Publisher (`.pub`) files and convert them to PDF using LibreOffice in headless mode.

The script is designed for large batch conversions and includes:

- Parallel processing
- Manifest logging
- Simulation mode (WhatIf)
- Skip existing PDF detection
- Cross-platform compatibility

Converted PDFs are written to the **same directory as the original `.pub` file**.

---

# Features

## Recursive scanning
The script scans a root folder and all subfolders for `.pub` files.

## Parallel processing
Multiple CPU cores can be used to speed up conversions.

## Simulation mode (`--whatif`)
Runs the entire process without generating PDFs.

Useful for:

- testing automation pipelines
- previewing which files will convert
- generating a manifest for downstream scripts

## Skip existing PDFs
If a PDF already exists for a `.pub` file, the script can skip conversion.

## Conversion manifest output
Two log files are produced:

| File | Description |
|-----|-------------|
| `conversion_manifest.csv` | Excel-friendly conversion report |
| `conversion_manifest.jsonl` | Structured logging format |

Each entry records:

- timestamp
- pub file path
- expected pdf path
- conversion status
- return code
- error details

---

# Requirements

## Python

Python 3.8 or newer.

Download:
https://www.python.org/downloads/

## LibreOffice

LibreOffice is used as the conversion engine.

Download:
https://www.libreoffice.org/download/download/

Documentation:
https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html

Typical Windows installation path:
