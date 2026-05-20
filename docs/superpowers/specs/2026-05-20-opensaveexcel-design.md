# opensaveexcel.py Design

## Goal

Make `eunwol1991/projects/opensaveexcel.py` a reliable Windows utility for batch-opening Excel workbooks with Microsoft Excel and saving them again.

## Runtime Scope

- Primary environment: Windows with Microsoft Excel installed.
- Python dependency: `xlwings`.
- Supported workbook extensions: `.xlsx`, `.xlsm`, `.xlsb`, `.xls`.
- Only workbook filenames containing `xx26` are processed; this match is case-insensitive.
- Temporary Excel lock files beginning with `~$` are skipped.

## User Flow

1. User runs the script.
2. The script prompts the user to input the folder path.
3. The script validates the folder and dependency availability before starting Excel.
4. Each supported workbook whose filename contains `xx26` is opened, saved, and closed.
5. The script prints a summary with success, failure, and skipped counts.

## Error Handling

- Missing `xlwings` produces a clear installation message.
- Invalid or missing folder paths produce a clear error and non-zero exit code.
- Excel startup failure produces a clear error and non-zero exit code.
- A failure in one workbook is reported but does not stop remaining files.
- Workbooks and the Excel application are closed in `finally` paths to reduce leftover Excel processes.

## Verification

- Python syntax must compile successfully.
- Script help must run without starting Excel.
- Runtime paths should avoid requiring Excel during automated validation in non-Windows CI.
