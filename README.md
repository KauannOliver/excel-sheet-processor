<div align="center">

# Excel Sheet Processor

A desktop utility for grouping, validating, and formatting spreadsheet data into a cleaner workbook.

**Python · Flet · Pandas · OpenPyXL**

</div>

## What it does

The interface accepts spreadsheet files, processes rows in chunks, applies grouping and number formats, and writes a structured output workbook. The source code includes checks for missing or malformed inputs.

## Run locally

Requires Python and the packages listed in `requirements.txt`. Run `pip install -r requirements.txt`, then `python main.py`. The desktop file-picker flow requires a graphical session.

## Project structure

`main.py` contains the Flet interface and spreadsheet processing flow.

No license is defined in this repository.
