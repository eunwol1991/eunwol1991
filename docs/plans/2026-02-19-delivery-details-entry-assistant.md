# Delivery Details Entry Assistant Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a local plugin-like Python desktop tool that opens an Excel workbook, finds `Delivery details` (case-insensitive), suggests newest matching history, and inserts one validated row while preserving formulas and formatting.

**Architecture:** Use `tkinter + openpyxl + rapidfuzz` with a modular package (`delivery_assistant`) under `eunwol1991/projects/function/`. Split concerns into sheet/header detection, history indexing/matching, row insertion, and UI orchestration. Insert workflow uses template-row snapshot + protected formula columns to avoid breaking totals/GST formulas.

**Tech Stack:** Python 3, tkinter/ttk, openpyxl, rapidfuzz, pathlib, datetime.

---

### Task 1: Create Package Skeleton

**Files:**
- Create: `eunwol1991/projects/function/delivery_assistant/__init__.py`
- Create: `eunwol1991/projects/function/delivery_assistant/__main__.py`
- Create: `eunwol1991/projects/function/delivery_assistant/ui_tk.py`
- Create: `eunwol1991/projects/function/delivery_assistant/excel_io.py`
- Create: `eunwol1991/projects/function/delivery_assistant/sheet_locator.py`
- Create: `eunwol1991/projects/function/delivery_assistant/history_index.py`
- Create: `eunwol1991/projects/function/delivery_assistant/matching.py`
- Create: `eunwol1991/projects/function/delivery_assistant/row_insert.py`
- Create: `eunwol1991/projects/function/delivery_assistant/schema.py`

**Step 1: Write the failing import smoke test**

```python
def test_delivery_assistant_importable():
    import eunwol1991.projects.function.delivery_assistant as pkg
    assert pkg is not None
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_delivery_assistant_smoke.py -v`
Expected: FAIL with import/module not found.

**Step 3: Add minimal package files and entrypoint**

```python
def main() -> int:
    from .ui_tk import launch
    launch()
    return 0
```

**Step 4: Run test to verify it passes**

Run: `pytest tests/test_delivery_assistant_smoke.py -v`
Expected: PASS.

**Step 5: Commit**

```bash
git add eunwol1991/projects/function/delivery_assistant tests/test_delivery_assistant_smoke.py
git commit -m "add delivery assistant package skeleton"
```

### Task 2: Sheet Detection + Header Mapping

**Files:**
- Modify: `eunwol1991/projects/function/delivery_assistant/sheet_locator.py`
- Modify: `eunwol1991/projects/function/delivery_assistant/schema.py`
- Test: `tests/test_sheet_locator.py`

**Step 1: Write failing tests for sheet/header detection**

```python
def test_find_delivery_details_case_insensitive(tmp_path):
    ...

def test_detect_header_row_from_expected_columns(tmp_path):
    ...
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_sheet_locator.py -v`
Expected: FAIL (functions missing).

**Step 3: Implement detection**

```python
def find_sheet_ci(wb, target: str):
    key = target.strip().casefold()
    ...

def detect_header_row(ws, expected_cols: set[str], scan_rows: int = 20) -> int:
    ...
```

**Step 4: Run test to verify it passes**

Run: `pytest tests/test_sheet_locator.py -v`
Expected: PASS.

**Step 5: Commit**

```bash
git add eunwol1991/projects/function/delivery_assistant/sheet_locator.py eunwol1991/projects/function/delivery_assistant/schema.py tests/test_sheet_locator.py
git commit -m "add case-insensitive Delivery details sheet and header detection"
```

### Task 3: History Index + Recency-first Matching

**Files:**
- Modify: `eunwol1991/projects/function/delivery_assistant/history_index.py`
- Modify: `eunwol1991/projects/function/delivery_assistant/matching.py`
- Test: `tests/test_matching_recency.py`

**Step 1: Write failing tests for scoring and recency tie-break**

```python
def test_recency_bias_prefers_newer_row_when_scores_equal():
    ...

def test_confidence_thresholds():
    ...
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_matching_recency.py -v`
Expected: FAIL.

**Step 3: Implement matching engine**

```python
def score_candidate(query, record) -> float:
    ...

def rank_candidates(query, records):
    ...
```

**Step 4: Run test to verify it passes**

Run: `pytest tests/test_matching_recency.py -v`
Expected: PASS.

**Step 5: Commit**

```bash
git add eunwol1991/projects/function/delivery_assistant/history_index.py eunwol1991/projects/function/delivery_assistant/matching.py tests/test_matching_recency.py
git commit -m "add recency-first fuzzy ranking for customer and outlet suggestions"
```

### Task 4: Safe Row Insert With Formula/Style Preservation

**Files:**
- Modify: `eunwol1991/projects/function/delivery_assistant/row_insert.py`
- Test: `tests/test_row_insert_preserve_formula_style.py`

**Step 1: Write failing tests for insert + protected formulas**

```python
def test_insert_row_copies_style_and_formula_from_template(tmp_path):
    ...

def test_user_input_does_not_overwrite_formula_columns(tmp_path):
    ...
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_row_insert_preserve_formula_style.py -v`
Expected: FAIL.

**Step 3: Implement insertion strategy**

```python
def insert_with_template(ws, insert_row: int, template_row: int, ...):
    # snapshot style/formulas
    # ws.insert_rows(...)
    # translate formulas to new row
    # write user columns only
```

**Step 4: Run test to verify it passes**

Run: `pytest tests/test_row_insert_preserve_formula_style.py -v`
Expected: PASS.

**Step 5: Commit**

```bash
git add eunwol1991/projects/function/delivery_assistant/row_insert.py tests/test_row_insert_preserve_formula_style.py
git commit -m "insert Delivery details row with preserved styles and formulas"
```

### Task 5: UI Form + Suggest + Preview + Confirm Insert

**Files:**
- Modify: `eunwol1991/projects/function/delivery_assistant/ui_tk.py`
- Modify: `eunwol1991/projects/function/delivery_assistant/excel_io.py`
- Test: `tests/test_ui_insert_plan.py`

**Step 1: Write failing tests for insert plan generation**

```python
def test_build_insert_plan_contains_user_fields_and_target_row():
    ...
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_ui_insert_plan.py -v`
Expected: FAIL.

**Step 3: Implement UI orchestration**

```python
class DeliveryAssistantApp(tk.Tk):
    ...
```

**Step 4: Run test to verify it passes**

Run: `pytest tests/test_ui_insert_plan.py -v`
Expected: PASS.

**Step 5: Commit**

```bash
git add eunwol1991/projects/function/delivery_assistant/ui_tk.py eunwol1991/projects/function/delivery_assistant/excel_io.py tests/test_ui_insert_plan.py
git commit -m "add Delivery details data-entry UI with suggestion and confirm-insert flow"
```

### Task 6: End-to-end Validation and User Runbook

**Files:**
- Create: `eunwol1991/projects/function/delivery_assistant/README.md`

**Step 1: Write failing/manual checklist**

```python
def test_manual_runbook_exists():
    ...
```

**Step 2: Run project checks**

Run: `pytest -q`
Expected: PASS for new tests.

**Step 3: Run syntax checks**

Run: `python -m compileall "eunwol1991/projects/function/delivery_assistant"`
Expected: no compile errors.

**Step 4: Document runbook**

```markdown
python -m eunwol1991.projects.function.delivery_assistant
```

**Step 5: Commit**

```bash
git add eunwol1991/projects/function/delivery_assistant/README.md
git commit -m "document delivery assistant usage and validation checklist"
```
