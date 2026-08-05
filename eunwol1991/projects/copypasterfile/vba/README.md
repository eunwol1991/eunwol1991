# RepairDoMerges VBA Modules

These source-controlled VBA files can be imported into the existing Excel macro workbook, such as `for excel printing.xlsm`, without editing the binary workbook in this repository.

## Files

- `RepairDoMerges.bas`: production macros for repairing missing `DO` sheet merge ranges.
- `RepairDoMergesTests.bas`: self-test macros that can be imported into a test copy of the macro workbook.

## Import

1. Open a copied `.xlsm` macro workbook in Excel.
2. Press `Alt+F11` to open the VBA editor.
3. Choose `File > Import File...`.
4. Import `RepairDoMerges.bas`.
5. Optionally import `RepairDoMergesTests.bas` into a test copy and run `RunRepairDoMergesTests` from the Immediate window or Macro dialog.

## Recommended Workflow

1. Work on copied target/reference files first. Do not start with production workbooks.
2. Run `RepairDoMergesDryRun` first. Review the Immediate window output for `printArea=[DO=..., Invoice=...]`, `invoiceFormula=`, `add=`, `conflict=`, `format=`, `SKIP`, and `saved=False` messages.
3. After verifying the dry-run output, run `RepairDoMergesApply` to repair target `DO` and `Invoice` print areas, `Invoice` subtotal/GST/total formulas, `DO` missing merge ranges, and the format recalculation hook in one pass.

## Custom Paths

Use `RepairDoMergesWithOptions` when testing alternate copied folders:

```vb
RepairDoMergesWithOptions "C:\path\to\copied targets", "C:\path\to\copied references", False
RepairDoMergesWithOptions "C:\path\to\copied targets", "C:\path\to\copied references", True
```

The target folder scan is non-recursive unless `recursive:=True` is passed. The reference folder scan is always recursive to match the Python workflow.

Target workbooks must be `.xlsx` or `.xlsm` and must contain `DO & INV` in the filename. If the filename has a final parenthesized outlet name, matching uses that outlet. If it has no outlet parentheses, matching falls back to the first filename token, such as `AK` or `GPTG`, and prefers a same-prefix `xx26` reference workbook.

## Safety Notes

- Merge repair edits only the sheet named `DO`; print area repair and format recalculation also process `Invoice` when it exists.
- It imports exact missing merge ranges from the selected reference `DO` sheet.
- During one batch run, each selected reference workbook is opened only once; its merge ranges are cached and reused for matching target workbooks.
- Workbooks are opened with Excel automation macros temporarily disabled, including `.xlsm` files, then the previous automation security setting is restored.
- Print area repair runs before merge repair on both `DO` and `Invoice` when those sheets exist.
- Invoice formula repair updates the amount cells to the right of rows labelled `Subtotal`, `GST 9%`, and `Total`. For legacy layouts this is usually column `I`; for merged layouts such as ABR it resolves the amount merge block to the right of the label, such as `K:M`, and writes to that block's anchor cell. If no safe amount cell can be inferred, it skips formula repair instead of overwriting label text. This runs for every processed workbook with an `Invoice` sheet; there is no `xx25` filename filter.
- The matching reference sheet supplies the print-area right edge, while the target sheet supplies the final row from current content so deleted rows, such as MOS files with removed lines, do not regain extra blank printable space.
- The repaired print-area last row is capped at the matching reference print-area last row. This means delete-row files can become smaller than the original, but runaway target print areas will not grow larger than the reference.
- Print area repair disables manual zoom (`PageSetup.Zoom = False`) and sets `Scale to Fit > Width` and `Height` to `1 page` (`PageSetup.FitToPagesWide = 1` and `PageSetup.FitToPagesTall = 1`), so Page Break Preview should show the repaired print area fitting one page wide and high.
- Apply mode also switches processed `DO` and `Invoice` sheets to Page Break Preview (`ActiveWindow.View = xlPageBreakPreview`) so reopening/checking the workbook shows the page-break layout directly. A view-only change is treated as a save-worthy change.
- Print area status is logged per sheet as `repaired`, `unchanged`, or `skipped-empty`.
- Exact reference merges are not extended to column `K`; for example, a reference merge `A22:G23` stays `A22:G23`.
- It skips a proposed merge if any non-anchor target cell in that range already contains a value, and logs the skipped range as a conflict.
- Signature labels `Received In Good Order` and `Authorised Signature & Stamp` merge from the label cell to the target sheet print-area right edge. If the print area is missing or invalid, the only fallback is column `K`.
- Merge repair still only edits `DO`; it does not merge cells on `Invoice`.
- Apply mode saves a target workbook when print area repair, invoice formula repair, merge repair, or format recalculation changes the workbook.
- The batch macros temporarily disable screen updating, events, alerts, and automatic calculation, then restore the original Excel settings even if an error occurs.

## Self-Tests

`RepairDoMergesTests.bas` includes checks for the risky boundaries:

- outlet extraction uses the last parenthesized part of the filename;
- reference selection prefers same prefix and `xx26` format files;
- target files without outlet parentheses can match references by filename prefix;
- print area repair uses reference columns and target content rows for `DO` and `Invoice`;
- invoice formula repair updates subtotal, GST, and total formulas for both legacy column layouts and merged right-side amount blocks while dry-run reports pending changes without mutating cells;
- print area repair caps target rows at the reference print-area last row;
- print area repair disables zoom and sets scale-to-fit width and height to one page;
- apply mode switches processed sheets to Page Break Preview;
- empty sheets keep their existing print area and report `skipped-empty`;
- dry-run reports print-area changes without mutating the sheet;
- exact reference ranges do not expand to arbitrary columns;
- signature label ranges stop at the print-area right edge;
- missing print area falls back to column `K`;
- conflicting non-anchor values prevent merging;
- dry-run reports ranges without merging.
