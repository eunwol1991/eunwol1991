# RepairDoMerges Print Area Repair Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `RepairDoMergesApply` repair target `DO` and `Invoice` print areas before `DO` merge repair, then run a format recalculation hook last for both sheets.

**Architecture:** Keep the current source-controlled VBA module structure. Add small public helper functions to parse and repair print areas so they can be covered by `RepairDoMergesTests.bas`, then thread print-area status and format status through the existing workbook repair result string.

**Tech Stack:** Excel VBA, `PageSetup.PrintArea`, existing self-test module `RepairDoMergesTests.bas`.

---

## File Structure

- Modify `eunwol1991/projects/copypasterfile/vba/RepairDoMerges.bas`: add print-area parsing/repair helpers, call `DO` and `Invoice` print-area repair before `DO` merge repair, add a format recalculation hook after merge repair, and save when any apply-mode step changes the workbook.
- Modify `eunwol1991/projects/copypasterfile/vba/RepairDoMergesTests.bas`: add focused self-tests for print-area repair, dry-run behavior, and signature merge order.
- Modify `eunwol1991/projects/copypasterfile/vba/README.md`: update workflow and safety notes for the new print-area/format sequence.

### Task 1: Add Print Area Repair Tests

**Files:**
- Modify: `eunwol1991/projects/copypasterfile/vba/RepairDoMergesTests.bas`

- [ ] **Step 1: Add tests to the runner**

Modify `RunRepairDoMergesTests` so the print-area tests run before existing signature merge tests:

```vb
Public Sub RunRepairDoMergesTests()
    TestExtractOutletNameUsesLastParenthesizedPart
    TestNormalizeTextCollapsesWhitespace
    TestSelectReferencePrefersSamePrefixThenXx26
    TestPrintAreaRepairUsesReferenceColumnsAndTargetLastRow
    TestPrintAreaRepairFallsBackToKWithTargetLastRow
    TestPrintAreaRepairDryRunDoesNotChangeSheet
    TestPrintAreaRepairRunsBeforeSignatureMerge
    TestExactReferenceMergeUsesOnlyProvidedAddress
    TestPrintAreaBoundaryStopsAtRightEdge
    TestSignatureFindKeepsNormalizedLabelMatching
    TestPrintAreaFallbackStopsAtK
    TestConflictSkipLeavesWorkbookUnmerged
    TestDryRunDoesNotSaveOrMerge
    Debug.Print "RepairDoMergesTests: all tests passed"
End Sub
```

- [ ] **Step 2: Add hybrid reference/target print-area test**

Append this test before `TestExactReferenceMergeUsesOnlyProvidedAddress`:

```vb
Public Sub TestPrintAreaRepairUsesReferenceColumnsAndTargetLastRow()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"
    With wb.Worksheets("DO")
        .Range("B48").Value = "last target row"
        .PageSetup.PrintArea = "A1:X99"
    End With

    Dim changed As Boolean
    changed = RepairDoMerges_RepairPrintArea(wb.Worksheets("DO"), "A1:K54", True)

    AssertTrue changed, "print area should change"
    AssertEquals "$A$1:$K$48", wb.Worksheets("DO").PageSetup.PrintArea, "reference columns plus target last row"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 710, "TestPrintAreaRepairUsesReferenceColumnsAndTargetLastRow", message
End Sub
```

- [ ] **Step 3: Add missing reference print-area fallback test**

Append this test after the hybrid test:

```vb
Public Sub TestPrintAreaRepairFallsBackToKWithTargetLastRow()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"
    wb.Worksheets("DO").Range("C33").Value = "last target row"

    Dim changed As Boolean
    changed = RepairDoMerges_RepairPrintArea(wb.Worksheets("DO"), "", True)

    AssertTrue changed, "fallback print area should change"
    AssertEquals "$A$1:$K$33", wb.Worksheets("DO").PageSetup.PrintArea, "fallback K plus target last row"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 711, "TestPrintAreaRepairFallsBackToKWithTargetLastRow", message
End Sub
```

- [ ] **Step 4: Add dry-run test**

Append this test after the fallback test:

```vb
Public Sub TestPrintAreaRepairDryRunDoesNotChangeSheet()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"
    With wb.Worksheets("DO")
        .Range("A20").Value = "last target row"
        .PageSetup.PrintArea = "A1:J10"
    End With

    Dim changed As Boolean
    changed = RepairDoMerges_RepairPrintArea(wb.Worksheets("DO"), "A1:K54", False)

    AssertTrue changed, "dry-run should report pending print-area change"
    AssertEquals "$A$1:$J$10", wb.Worksheets("DO").PageSetup.PrintArea, "dry-run must not mutate print area"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 712, "TestPrintAreaRepairDryRunDoesNotChangeSheet", message
End Sub
```

- [ ] **Step 5: Add ordering test for signature merge**

Append this test after the dry-run test:

```vb
Public Sub TestPrintAreaRepairRunsBeforeSignatureMerge()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"
    With wb.Worksheets("DO")
        .Range("I30").Value = "Received In Good Order"
        .Range("A30").Value = "content"
        .PageSetup.PrintArea = "A1:X30"
    End With

    Dim changed As Boolean
    changed = RepairDoMerges_RepairPrintArea(wb.Worksheets("DO"), "A1:K54", True)

    Dim added As Collection
    Dim conflicts As Collection
    Set added = New Collection
    Set conflicts = New Collection
    RepairDoMerges_RepairSignatureLabels wb.Worksheets("DO"), True, added, conflicts

    AssertTrue changed, "print area should be repaired first"
    AssertEquals "$A$1:$K$30", wb.Worksheets("DO").PageSetup.PrintArea, "repaired print area"
    AssertTrue RepairDoMerges_IsExactMerged(wb.Worksheets("DO"), "I30:K30"), "signature should merge to repaired K edge"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 713, "TestPrintAreaRepairRunsBeforeSignatureMerge", message
End Sub
```

- [ ] **Step 6: Run tests and verify they fail**

Import the updated `RepairDoMergesTests.bas` into a test macro workbook and run:

```vb
RunRepairDoMergesTests
```

Expected: compile failure or runtime failure because `RepairDoMerges_RepairPrintArea` is not defined yet.

### Task 2: Implement Print Area Helpers

**Files:**
- Modify: `eunwol1991/projects/copypasterfile/vba/RepairDoMerges.bas`

- [ ] **Step 1: Add constants near existing constants**

Add this below `SIGNATURE_END_COLUMN_FALLBACK`:

```vb
Private Const PRINT_AREA_START_ROW As Long = 1
Private Const PRINT_AREA_START_COLUMN As Long = 1
```

- [ ] **Step 2: Add public print-area repair function**

Add this after `RepairDoMerges_RepairSignatureLabels`:

```vb
Public Function RepairDoMerges_RepairPrintArea(ByVal targetSheet As Worksheet, ByVal referencePrintArea As String, ByVal applyChanges As Boolean) As Boolean
    Dim endColumn As Long
    Dim endRow As Long
    Dim desiredAddress As String
    Dim currentAddress As String

    endColumn = RepairDoMerges_PrintAreaEndColumnFromAddress(targetSheet, referencePrintArea)
    endRow = RepairDoMerges_LastContentRow(targetSheet)
    If endRow = 0 Then Exit Function

    desiredAddress = targetSheet.Range(targetSheet.Cells(PRINT_AREA_START_ROW, PRINT_AREA_START_COLUMN), targetSheet.Cells(endRow, endColumn)).Address
    currentAddress = targetSheet.PageSetup.PrintArea

    If RepairDoMerges_NormalizePrintAreaAddress(currentAddress) = RepairDoMerges_NormalizePrintAreaAddress(desiredAddress) Then Exit Function

    RepairDoMerges_RepairPrintArea = True
    If applyChanges Then targetSheet.PageSetup.PrintArea = desiredAddress
End Function
```

- [ ] **Step 3: Add print-area right-edge helper**

Add this near existing `RepairDoMerges_PrintAreaEndColumn`:

```vb
Public Function RepairDoMerges_PrintAreaEndColumnFromAddress(ByVal sheet As Worksheet, ByVal printArea As String) As Long
    Dim firstArea As String
    Dim bangPos As Long
    Dim rangeAddress As String

    firstArea = Trim$(printArea)
    If Len(firstArea) = 0 Then
        RepairDoMerges_PrintAreaEndColumnFromAddress = SIGNATURE_END_COLUMN_FALLBACK
        Exit Function
    End If

    firstArea = Split(firstArea, ",")(0)
    bangPos = InStrRev(firstArea, "!")
    If bangPos > 0 Then firstArea = Mid$(firstArea, bangPos + 1)
    rangeAddress = Replace(Replace(firstArea, "'", ""), "$", "")
    On Error GoTo UseFallback
    RepairDoMerges_PrintAreaEndColumnFromAddress = sheet.Range(rangeAddress).Columns(sheet.Range(rangeAddress).Columns.Count).Column
    Exit Function
UseFallback:
    RepairDoMerges_PrintAreaEndColumnFromAddress = SIGNATURE_END_COLUMN_FALLBACK
End Function
```

- [ ] **Step 4: Simplify existing print-area edge function to reuse helper**

Replace `RepairDoMerges_PrintAreaEndColumn` with:

```vb
Public Function RepairDoMerges_PrintAreaEndColumn(ByVal sheet As Worksheet) As Long
    RepairDoMerges_PrintAreaEndColumn = RepairDoMerges_PrintAreaEndColumnFromAddress(sheet, sheet.PageSetup.PrintArea)
End Function
```

- [ ] **Step 5: Add target last-content-row helper**

Add this near the print-area helpers:

```vb
Public Function RepairDoMerges_LastContentRow(ByVal sheet As Worksheet) As Long
    Dim found As Range

    Set found = sheet.Cells.Find(What:="*", After:=sheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not found Is Nothing Then RepairDoMerges_LastContentRow = found.Row
End Function
```

- [ ] **Step 6: Add print-area normalization helper**

Add this near the print-area helpers:

```vb
Private Function RepairDoMerges_NormalizePrintAreaAddress(ByVal printArea As String) As String
    RepairDoMerges_NormalizePrintAreaAddress = UCase$(Replace(Trim$(printArea), "'", ""))
End Function
```

- [ ] **Step 7: Run print-area tests**

Import both VBA modules into a test macro workbook and run:

```vb
RunRepairDoMergesTests
```

Expected: the new print-area tests pass. Existing workbook-level tests may still need later tasks for batch integration.

### Task 3: Integrate Print Area Repair Into Workbook Flow

**Files:**
- Modify: `eunwol1991/projects/copypasterfile/vba/RepairDoMerges.bas`
- Modify: `eunwol1991/projects/copypasterfile/vba/RepairDoMergesTests.bas`

- [ ] **Step 1: Cache reference print areas**

In `RepairDoMergesWithOptions`, add a new dictionary beside `referenceMergeCache`:

```vb
Dim referencePrintAreaCache As Object
```

Initialize it after `referenceMergeCache`:

```vb
Set referencePrintAreaCache = CreateObject("Scripting.Dictionary")
```

- [ ] **Step 2: Pass reference print area into workbook repair**

In the target loop, before calling `RepairDoMerges_RepairWorkbookWithReferenceRanges`, add:

```vb
Dim referencePrintAreas As Object
Set referencePrintAreas = RepairDoMerges_GetCachedReferencePrintAreas(referencePath, referencePrintAreaCache)
```

Change the call to:

```vb
resultLine = RepairDoMerges_RepairWorkbookWithReferenceRanges(CStr(targetPath), referencePath, referenceMergeRanges, referencePrintAreas, applyChanges)
```

- [ ] **Step 3: Update direct repair entry point**

In `RepairDoMerges_RepairWorkbook`, add:

```vb
Dim referencePrintAreas As Object
Set referencePrintAreas = RepairDoMerges_LoadReferencePrintAreas(referencePath)
```

Change the return assignment to:

```vb
RepairDoMerges_RepairWorkbook = RepairDoMerges_RepairWorkbookWithReferenceRanges(targetPath, referencePath, referenceMergeRanges, referencePrintAreas, applyChanges)
```

- [ ] **Step 4: Update private repair signature and save condition**

Change the function signature to:

```vb
Private Function RepairDoMerges_RepairWorkbookWithReferenceRanges(ByVal targetPath As String, ByVal referencePath As String, ByVal referenceMergeRanges As Collection, ByVal referencePrintAreas As Object, ByVal applyChanges As Boolean) As String
```

Add local variables with the existing declarations:

```vb
Dim printAreaChanged As Boolean
Dim printAreaStatuses As Collection
Dim formatChanged As Boolean
```

Before the merge loop, add:

```vb
printAreaChanged = RepairDoMerges_RepairPrintArea(targetSheet, RepairDoMerges_ReferenceSheetPrintArea(referencePrintAreas, "DO"), applyChanges, doPrintAreaStatus)
If WorkbookHasSheet(targetWb, "Invoice") Then
    Set invoiceSheet = targetWb.Worksheets("Invoice")
    printAreaChanged = RepairDoMerges_RepairPrintArea(invoiceSheet, RepairDoMerges_ReferenceSheetPrintArea(referencePrintAreas, "Invoice"), applyChanges, invoicePrintAreaStatus) Or printAreaChanged
End If
```

After `RepairDoMerges_RepairSignatureLabels`, add:

```vb
formatChanged = RepairDoMerges_RecalculateFormat(targetSheet, applyChanges)
If Not invoiceSheet Is Nothing Then formatChanged = RepairDoMerges_RecalculateFormat(invoiceSheet, applyChanges) Or formatChanged
```

Replace the save calculation with:

```vb
saved = ((printAreaChanged Or added.Count > 0 Or formatChanged) And applyChanges)
```

Replace the result string with:

```vb
RepairDoMerges_RepairWorkbookWithReferenceRanges = FileNameOnly(targetPath) & ": reference=" & FileNameOnly(referencePath) & "; printArea=" & CollectionToText(printAreaStatuses) & "; add=" & added.Count & " " & CollectionToText(added) & "; conflict=" & conflicts.Count & " " & CollectionToText(conflicts) & "; format=" & CStr(formatChanged) & "; saved=" & CStr(saved)
```

- [ ] **Step 5: Add reference print-area loaders**

Add these near the reference merge cache/load helpers:

```vb
Private Function RepairDoMerges_GetCachedReferencePrintAreas(ByVal referencePath As String, ByVal referencePrintAreaCache As Object) As Object
    If Not referencePrintAreaCache.Exists(referencePath) Then referencePrintAreaCache.Add referencePath, RepairDoMerges_LoadReferencePrintAreas(referencePath)
    Set RepairDoMerges_GetCachedReferencePrintAreas = referencePrintAreaCache.Item(referencePath)
End Function

Private Function RepairDoMerges_LoadReferencePrintAreas(ByVal referencePath As String) As Object
    Dim referenceWb As Workbook
    Dim printAreas As Object

    Set printAreas = CreateObject("Scripting.Dictionary")

    On Error GoTo CleanFail
    Set referenceWb = RepairDoMerges_OpenWorkbook(referencePath, True)

    If WorkbookHasSheet(referenceWb, "DO") Then printAreas.Add "DO", referenceWb.Worksheets("DO").PageSetup.PrintArea
    If WorkbookHasSheet(referenceWb, "Invoice") Then printAreas.Add "Invoice", referenceWb.Worksheets("Invoice").PageSetup.PrintArea
    Set RepairDoMerges_LoadReferencePrintAreas = printAreas

CleanExit:
    On Error Resume Next
    referenceWb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    referenceWb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 606, "RepairDoMerges_LoadReferencePrintAreas", message
End Function
```

- [ ] **Step 6: Run all VBA tests**

Import both VBA modules into a test macro workbook and run:

```vb
RunRepairDoMergesTests
```

Expected: all tests pass.

### Task 4: Add Format Recalculation Hook

**Files:**
- Modify: `eunwol1991/projects/copypasterfile/vba/RepairDoMerges.bas`

- [ ] **Step 1: Add a no-op hook with clear edit point**

Add this after `RepairDoMerges_RepairPrintArea`:

```vb
Public Function RepairDoMerges_RecalculateFormat(ByVal targetSheet As Worksheet, ByVal applyChanges As Boolean) As Boolean
    Dim wasSaved As Boolean

    If Not applyChanges Then Exit Function

    wasSaved = targetSheet.Parent.Saved
    targetSheet.Calculate
    RepairDoMerges_RecalculateFormat = (wasSaved And Not targetSheet.Parent.Saved)
End Function
```

- [ ] **Step 2: Optional production wrapper for an existing macro**

When importing into a production macro workbook that already exposes a procedure named `RecalculateFormat`, use this body instead of the `targetSheet.Calculate` hook:

```vb
Public Function RepairDoMerges_RecalculateFormat(ByVal targetSheet As Worksheet, ByVal applyChanges As Boolean) As Boolean
    Dim wasSaved As Boolean

    If Not applyChanges Then Exit Function

    wasSaved = targetSheet.Parent.Saved
    targetSheet.Parent.Activate
    targetSheet.Activate
    Application.Run "RecalculateFormat"
    RepairDoMerges_RecalculateFormat = (wasSaved And Not targetSheet.Parent.Saved)
End Function
```

- [ ] **Step 3: Run all VBA tests**

Import both VBA modules into a test macro workbook and run:

```vb
RunRepairDoMergesTests
```

Expected: all tests pass. If the test workbook does not contain an external `RecalculateFormat` macro, keep the no-op `targetSheet.Calculate` hook until the production macro name is confirmed.

### Task 5: Update Documentation

**Files:**
- Modify: `eunwol1991/projects/copypasterfile/vba/README.md`

- [ ] **Step 1: Update recommended workflow**

Replace lines 20-22 with:

```md
1. Work on copied target/reference files first. Do not start with production workbooks.
2. Run `RepairDoMergesDryRun` first. Review the Immediate window output for `printArea=[DO=..., Invoice=...]`, `add=`, `conflict=`, `format=`, and `saved=False` messages.
3. After verifying the dry-run output, run `RepairDoMergesApply` to repair target `DO` and `Invoice` print areas, `DO` missing merge ranges, and the format recalculation hook in one pass.
```

- [ ] **Step 2: Update safety notes**

Add these bullets before the existing signature-label note:

```md
- Print area repair runs before merge repair.
- The matching reference sheet supplies the print-area right edge, while the target sheet supplies the final row from current content so deleted rows stay deleted.
- The target ending row is capped at the matching reference print-area ending row, so delete-row files can shrink but runaway print areas cannot grow past the reference.
- Print area repair sets `PageSetup.Zoom = False`, `PageSetup.FitToPagesWide = 1`, and `PageSetup.FitToPagesTall = 1`, so Page Break Preview should show the repaired print area fitting one page wide and high.
- Apply mode switches processed `DO` and `Invoice` sheets to Page Break Preview with `ActiveWindow.View = xlPageBreakPreview`; a view-only change counts as save-worthy.
- Print area status is logged per sheet as `repaired`, `unchanged`, or `skipped-empty`.
```

Replace the apply-mode save bullet with:

```md
- Apply mode saves a target workbook when print area repair, merge repair, or format recalculation changes the workbook.
```

- [ ] **Step 3: Update self-test list**

Add these bullets before the existing signature-label bullet:

```md
- print area repair uses reference columns and target content rows;
- dry-run reports print-area changes without mutating the sheet;
```

### Task 6: Final Verification

**Files:**
- Verify: `eunwol1991/projects/copypasterfile/vba/RepairDoMerges.bas`
- Verify: `eunwol1991/projects/copypasterfile/vba/RepairDoMergesTests.bas`
- Verify: `eunwol1991/projects/copypasterfile/vba/README.md`

- [ ] **Step 1: Search for unresolved placeholders**

Run:

```bash
rg -n "TBD|TODO|implement later|fill in details" "eunwol1991/projects/copypasterfile/vba"
```

Expected: no output for files changed by this plan.

- [ ] **Step 2: Run VBA self-tests in Excel**

Import `RepairDoMerges.bas` and `RepairDoMergesTests.bas` into a copied macro workbook and run:

```vb
RunRepairDoMergesTests
```

Expected Immediate window output:

```text
RepairDoMergesTests: all tests passed
```

- [ ] **Step 3: Manual dry-run on copied workbooks**

Run this from the Immediate window against copied folders:

```vb
RepairDoMergesWithOptions "C:\path\to\copied targets", "C:\path\to\copied references", False
```

Expected: output includes `printArea=True` or `printArea=False`, `format=False`, and `saved=False`; target files are not modified.

- [ ] **Step 4: Manual apply on copied workbooks**

Run this from the Immediate window against copied folders:

```vb
RepairDoMergesWithOptions "C:\path\to\copied targets", "C:\path\to\copied references", True
```

Expected: target `DO` and `Invoice` print areas are repaired before `DO` signature merges, format hook runs last, and changed workbooks are saved.

Targets without outlet parentheses, such as `AK 0726 - 002 - DO & INV.xlsx`, should still match references by filename prefix and prefer same-prefix `xx26` reference files.

- [ ] **Step 5: Review git diff**

Run:

```bash
git diff -- "eunwol1991/projects/copypasterfile/vba/RepairDoMerges.bas" "eunwol1991/projects/copypasterfile/vba/RepairDoMergesTests.bas" "eunwol1991/projects/copypasterfile/vba/README.md"
```

Expected: diff only contains print-area repair, format hook, tests, and README updates described in this plan.
