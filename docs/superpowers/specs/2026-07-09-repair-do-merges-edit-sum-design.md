# RepairDoMerges Edit Sum Formula Repair Design

## Goal

Extend `eunwol1991/projects/copypasterfile/vba/RepairDoMerges.bas` so the existing dry-run/apply workflow also repairs `Invoice` sheet `Subtotal`, `GST 9%`, and `Total` formulas. The VBA behavior mirrors the useful part of `editsum.py` without its temporary `xx25` filename filter.

## Formula Behavior

- Process every target workbook already handled by `RepairDoMergesWithOptions`; do not filter by `xx25` or any other filename substring.
- Only edit the `Invoice` sheet when it exists.
- Find label rows by scanning string cells in the used range and matching normalized text containing `Subtotal`, `GST 9%`, and `Total`.
- Write formulas in amount column `I`.
- Set subtotal to `=SUM(I24:I{subtotalRow - 1})`.
- Set GST to `=I{subtotalRow}*0.09`.
- Set total to `=SUM(I{subtotalRow}:I{gstRow})`.
- If any required label row is missing, do not change formulas.

## Workflow Integration

Formula repair runs inside the existing `Invoice` branch of `RepairDoMerges_RepairWorkbookWithReferenceRanges`. It respects `applyChanges`: dry-run returns that formula repair is needed without mutating cells, while apply mode writes the formulas. The batch result includes `invoiceFormula=True/False`, and formula-only changes are save-worthy in apply mode.

## Tests

`RepairDoMergesTests.bas` covers apply mode writing the expected formulas and dry-run reporting pending formula changes while leaving the original formulas unchanged. These tests are VBA self-tests intended to run from Excel with `RunRepairDoMergesTests` after importing the modules into a test macro workbook.
