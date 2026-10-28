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




## Snapshot review

Setup and capabilities are based on the included code. External services and operating-system-specific integrations were not exercised during this publication review. Use synthetic data in a separate local environment.

Git-triggered Vercel deployment is disabled in the included web configurations. This update publishes source and documentation only.

## License

No new distribution license is granted for this corporate source edition. Public visibility does not grant reuse rights; obtain permission from the relevant rights holder. Existing third-party notices remain applicable.

## Verification record

See [snapshot checks and remaining limitations](docs/VERIFICATION.md).
