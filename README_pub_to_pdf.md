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

```
C:\Program Files\LibreOffice\program\soffice.exe
```

---

# Installation

1. Install Python
2. Install LibreOffice
3. Place `pub_to_pdf_batch.py` in your scripts folder.

---

# Basic Usage

```
python pub_to_pdf_batch.py <root_folder>
```

Example:

```
python pub_to_pdf_batch.py "C:\PublisherFiles"
```

This will:

- scan all subdirectories
- convert `.pub` files to `.pdf`
- place PDFs next to their source files
- generate a conversion manifest

---

# Simulation Mode

Run a simulated conversion without creating PDFs.

```
python pub_to_pdf_batch.py "C:\PublisherFiles" --whatif
```

This will:

- scan files
- compute expected output paths
- record simulated conversions
- not call LibreOffice

---

# Skip Existing PDFs

Avoid converting files that already have a valid PDF.

```
python pub_to_pdf_batch.py "C:\PublisherFiles" --skip-existing
```

---

# Recommended Pre-Run Test

Before running a large batch conversion:

```
python pub_to_pdf_batch.py "C:\PublisherFiles" --whatif --skip-existing
```

This lets you verify exactly what will happen.

---

# Performance Options

## Set worker threads

```
--workers 6
```

Example:

```
python pub_to_pdf_batch.py "C:\PublisherFiles" --workers 6
```

## Set conversion timeout

```
--timeout 240
```

Example:

```
python pub_to_pdf_batch.py "C:\PublisherFiles" --timeout 240
```

---

# Example Manifest Output

```
timestamp_utc,status,pub_path,pdf_path,returncode,whatif,error
2026-03-03T19:02:10Z,success,C:\docs\file.pub,C:\docs\file.pdf,0,False,
2026-03-03T19:02:11Z,success_simulated,C:\docs\test.pub,C:\docs\test.pdf,0,True,
```

Status values:

| Status | Meaning |
|------|------|
| success | Conversion completed |
| success_simulated | Simulated in WhatIf mode |
| skipped_existing | Existing PDF detected |
| fail | Conversion failed |

---

# Troubleshooting

## LibreOffice not found

Specify the path manually:

```
--soffice "C:\Program Files\LibreOffice\program\soffice.exe"
```

---

## Conversion failures

Check the `error` column in the manifest.

Common causes include:

- unsupported Publisher features
- corrupted `.pub` files
- password-protected documents

---

# Safety Notes

- `.pub` source files are never modified
- PDFs are written only alongside source files
- simulation mode creates no files

---

# License

Internal automation script. Modify as needed.
