# RepairDoMerges Print Area Repair Design

## Goal

Extend `eunwol1991/projects/copypasterfile/vba/RepairDoMerges.bas` so one apply run can repair the `DO` and `Invoice` sheet print areas, restore missing `DO` merge cells, and then trigger the existing format recalculation flow for both sheets. This removes the need to run a separate PowerShell/manual step after AI keyed delivery orders change the printable area.

## Target Workflow

1. User runs `RepairDoMergesDryRun` to preview changes.
2. User runs `RepairDoMergesApply` after checking the Immediate window output.
3. For each target workbook, the macro opens the matching reference workbook as it does today.
4. On the target `DO` and `Invoice` sheets, the macro repairs print areas first when those sheets exist.
5. The macro repairs missing reference merge ranges and signature-label merges on `DO` only.
6. The macro runs the format recalculation hook last for `DO` and `Invoice`.
7. The target workbook is saved only if apply mode made a print-area, merge, or format change.

## Print Area Strategy

Use a hybrid reference/target strategy:

- The matching reference sheet supplies the print area start column and end column.
- The target sheet supplies the ending row based on its current actual content, so deleted rows do not force the old reference page height back onto the target.
- The target ending row is capped at the matching reference print-area ending row, so repaired files may be shorter than the reference after row deletion but will not become taller than the reference when print areas run away.
- Print area repair sets `PageSetup.Zoom = False`, `PageSetup.FitToPagesWide = 1`, and `PageSetup.FitToPagesTall = 1`, so Page Break Preview and printing fit the repaired area to one page wide and high.
- Apply mode switches processed `DO` and `Invoice` sheets to Page Break Preview with `ActiveWindow.View = xlPageBreakPreview`; a view-only change counts as save-worthy.
- If the reference print area is missing or invalid, fall back to the existing right-edge behavior: column `K` is the print-area right boundary.
- If the target has no useful content, leave its existing print area unchanged and report a conflict/skip rather than guessing.

Example: if the reference print area is `A1:K54` and the target currently has content through row `48`, set the target print area to `A1:K48`.

## Merge Repair Interaction

The existing merge repair order should stay intact after print area repair:

- Exact reference merge ranges are still copied as exact addresses.
- Signature labels still merge from the label cell to the current target `DO` print-area right edge.
- Because print area is repaired first, signature labels use the corrected right edge instead of stale or missing page setup data.

## Format Recalculation Hook

The repository does not currently contain the user's existing `recalculate format` VBA macro. The implementation should therefore expose a small internal hook instead of hard-coding an unknown macro name:

- If the format logic is added to this module, call it after merge repair for `DO` and `Invoice`.
- If the real macro already exists in the macro workbook, call it through a clearly named wrapper that can be edited to delegate to the external macro name.
- Dry-run mode must report that format recalculation would run, but must not mutate the workbook.

## Logging

Immediate window output should make the batch easy to audit:

- Report whether each sheet print area was `unchanged`, `repaired`, or `skipped-empty`.
- Keep the existing `add=`, `conflict=`, and `saved=` merge output.
- Include whether format recalculation ran or was skipped.

## Safety

- Merge repair continues editing only the sheet named `DO`.
- Print area repair and format recalculation may edit `DO` and `Invoice`.
- Continue opening target workbooks read-only in dry-run mode.
- Continue disabling macros when opening target/reference workbooks.
- Do not overwrite a non-empty target print area with a guessed area if neither reference print area nor target content can identify a safe boundary.

## Tests

Update `RepairDoMergesTests.bas` with self-tests for:

- Reference `A1:K54` plus target content through row `48` produces `A1:K48`.
- Missing reference print area falls back to column `K` while using the target last content row.
- Print area repair runs before signature merge, so signature labels merge to the repaired right edge.
- Dry-run reports a print-area change without changing `PageSetup.PrintArea`.
- Empty sheets keep their existing print area and report `skipped-empty`.
- Manual zoom is disabled and scale-to-fit width and height are set to one page when print area repair applies.
- Processed sheets open/check in Page Break Preview after apply mode.
- Filenames must contain `DO & INV`. Targets with a final parenthesized outlet key match by outlet; targets without outlet parentheses fall back to same-prefix reference matching.

## Open Configuration Point

Before final import into the production macro workbook, confirm the real `recalculate format` macro name if it should call an existing procedure outside this source-controlled module.
