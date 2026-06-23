# Report Generator Runtime Options Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `reportgenerate.py` easier to run daily by selecting the source file at runtime, generating monthly sheets by month, weekly sheets by date range, and leaving weekly value/GST fields blank.

**Architecture:** Keep the single-file script structure. Add small helpers for source-file discovery, interactive/default inputs, date parsing, and row-level period filtering; then route `Weekly_Report` through week-filtered rows while monthly sheets keep month-filtered rows.

**Tech Stack:** Python 3, openpyxl, unittest.

---

### Task 1: Runtime source-file and period configuration

**Files:**
- Modify: `eunwol1991/projects/function/reportgenerate.py`
- Test: `eunwol1991/projects/function/test_reportgenerate.py`

- [ ] **Step 1: Add failing tests**

Add tests for newest source discovery and date parsing in `test_reportgenerate.py`:

```python
def test_find_latest_source_file_picks_newest_matching_file(self):
    module = _load_module()
    with tempfile.TemporaryDirectory() as temp_dir:
        folder = Path(temp_dir)
        old_file = folder / "Savori Sales Summary old.xlsx"
        new_file = folder / "Savori Sales Summary new.xlsx"
        old_file.write_text("old")
        new_file.write_text("new")
        os.utime(old_file, (1, 1))
        os.utime(new_file, (2, 2))
        self.assertEqual(module.find_latest_source_file(folder), new_file)

def test_parse_report_date_accepts_iso_date(self):
    module = _load_module()
    self.assertEqual(module.parse_report_date("2026-05-07").isoformat(), "2026-05-07")
```

- [ ] **Step 2: Run tests to verify failure**

Run: `python3 -m unittest eunwol1991.projects.function.test_reportgenerate`
Expected: fails because `find_latest_source_file` and `parse_report_date` do not exist.

- [ ] **Step 3: Implement helpers**

Add `import os`, `from datetime import datetime`, and `SOURCE_FOLDER = Path("/mnt/c/Users/jhunj/Dropbox/DO & INV")`. Implement:

```python
def find_latest_source_file(folder=SOURCE_FOLDER):
    files = list(Path(folder).glob("Savori Sales Summary*.xlsx"))
    if not files:
        raise FileNotFoundError(f"No Savori Sales Summary*.xlsx found in {folder}")
    return max(files, key=lambda path: path.stat().st_mtime)

def parse_report_date(value):
    return datetime.strptime(str(value).strip(), "%Y-%m-%d").date()
```

- [ ] **Step 4: Verify tests pass**

Run: `python3 -m unittest eunwol1991.projects.function.test_reportgenerate`
Expected: all tests pass.

### Task 2: Weekly date range filtering and blank weekly value columns

**Files:**
- Modify: `eunwol1991/projects/function/reportgenerate.py`
- Test: `eunwol1991/projects/function/test_reportgenerate.py`

- [ ] **Step 1: Add failing weekly tests**

Add a test creating source rows across two weeks. Call `build_weekly(ws_data, wb_out, start_date=date(2026, 5, 1), end_date=date(2026, 5, 7))`. Assert only rows in the range appear and columns 9, 10, 11 are `None`.

- [ ] **Step 2: Run tests to verify failure**

Run: `python3 -m unittest eunwol1991.projects.function.test_reportgenerate`
Expected: fails because `build_weekly` does not accept date range args and still writes formulas into columns 9-11.

- [ ] **Step 3: Implement weekly filtering**

Update `build_weekly(ws_data, wb_out, start_date=None, end_date=None)`. When scanning rows, include only rows where the source `Date` cell is within the inclusive range. Keep formulas for quantity fields only: `Qty in Pcs`, `Qty in Ctns`, `Total Qty in Pcs`, `Total Qty in Ctns`. Leave `Total Value`, `GST`, `Total Value Inclusive GST` blank.

- [ ] **Step 4: Verify tests pass**

Run: `python3 -m unittest eunwol1991.projects.function.test_reportgenerate`
Expected: all tests pass.

### Task 3: Interactive runtime defaults

**Files:**
- Modify: `eunwol1991/projects/function/reportgenerate.py`

- [ ] **Step 1: Implement runtime config prompts**

Add a `get_runtime_config()` helper that prompts:

```text
Source file [Enter = latest Savori Sales Summary*.xlsx]:
Year [2026]:
Month [May]:
Weekly start date YYYY-MM-DD [skip weekly filter]:
Weekly end date YYYY-MM-DD [same as start]:
```

Blank source uses `find_latest_source_file()`. Blank year/month use existing `FILTERS`. Blank weekly start leaves weekly unfiltered.

- [ ] **Step 2: Wire config into `main()`**

Use returned source path for `load_workbook`, update `FILTERS["Year"]` and `FILTERS["Month"]`, and pass weekly dates to `build_weekly`.

- [ ] **Step 3: Verify manually with non-interactive import path**

Run unit tests and `python3 -m py_compile eunwol1991/projects/function/reportgenerate.py eunwol1991/projects/function/test_reportgenerate.py`.

---

## Self-Review

- Spec coverage: source selection, month filters, weekly date filtering, and blank weekly value/GST fields are covered.
- Placeholder scan: no TBD/TODO placeholders remain.
- Type consistency: helpers use `Path` and `date` objects consistently; `build_weekly` accepts optional dates.
