# README - Processing Contribution Certificates on macOS

> **Created by Guillermo Malagon on April 8, 2026.**

This document explains how to run the `procesar_certificados_aportes.py` script on a MacBook to process multiple PDF files from a folder and generate a single consolidated Excel file.

## 1) Check Python on macOS

On macOS, the correct command is usually `python3`, not `python`.

Run this in Terminal:

```bash
python3 --version
```

If you also want to verify `pip`:

```bash
pip3 --version
```

If `pip3` does not respond, use:

```bash
python3 -m pip --version
```

## 2) Install dependencies

Install the required libraries:

```bash
python3 -m pip install pdfplumber openpyxl
```

## 3) Locate the script file

Save the Python file with this name:

```text
procesar_certificados_aportes.py
```

You can place it, for example, here:

```text
/Users/guillermomalagon/Downloads/procesar_certificados_aportes.py
```

## 4) Expected structure

You should have:

- a folder containing all the PDFs
- the Python script
- a path where the output Excel file will be saved

Example:

```text
/Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/pdfs
```

## 5) Command to run the script

### Option A: single line

```bash
python3 "/Users/guillermomalagon/Downloads/procesar_certificados_aportes.py" --input "/Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/pdfs" --output "/Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/salida.xlsx"
```

### Option B: multiple lines in zsh

```bash
python3 "/Users/guillermomalagon/Downloads/procesar_certificados_aportes.py" \
  --input "/Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/pdfs" \
  --output "/Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/salida.xlsx"
```

## 6) What the script does

The script:

- scans all `.pdf` files in the specified folder
- takes only the last person/ID found in the PDF header
- consolidates the information into a single Excel file
- creates these sheets:
  - `Liquidaciones Pagadas`
  - `Seguridad Social`
  - `Aportes Parafiscales`
  - `Novedades`
- stores:
  - numbers as numbers
  - dates as real Excel dates
  - rates as numbers
  - novedades with one row per `X`
- continues processing the remaining files even if one fails

## 7) Expected console messages

During execution, you will see messages like these:

```text
Processing 4 PDF files...
OK  - 648590.pdf
OK  - 3954297.pdf
ERROR - 1180890.pdf: Could not extract the header identity.
OK  - 5503152.pdf
```

## 8) Example of the final summary

At the end, the console will show a summary like this:

```text
================================================================================
FINAL SUMMARY
================================================================================
Total PDFs found         : 4
Processed successfully   : 3
With errors              : 1
Excel generated at       : /Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/salida.xlsx

Successful files:
  - 648590.pdf
  - 3954297.pdf
  - 5503152.pdf

Files with errors:
  - 1180890.pdf: Could not extract the header identity.
```

## 9) What to do if `zsh: command not found: python` appears

On macOS, use:

```bash
python3
```

Do not use:

```bash
python
```

Correct example:

```bash
python3 "/Users/guillermomalagon/Downloads/procesar_certificados_aportes.py" --input "/path/to/pdfs" --output "/path/to/output.xlsx"
```

## 10) What to do if a PDF fails

If a file returns an error, check:

- the `DEBUG TEXTO file_name.pdf` block
- the exact error message
- whether the PDF structure differs from the others

The script is designed to keep processing the remaining PDFs even if one fails. That prevents the whole process from crashing because of a single problematic file.

## 11) Practical recommendation

Before running 100 files at once, test first with 3 or 4 PDFs.  
That helps you verify that:

- the header is being extracted correctly
- the tables are being interpreted correctly
- the Excel output has the expected format

Confidence first, scale second.
