Attribute VB_Name = "RepairDoMergesTests"
Option Explicit

Public Sub RunRepairDoMergesTests()
    TestExtractOutletNameUsesLastParenthesizedPart
    TestNormalizeTextCollapsesWhitespace
    TestSelectReferencePrefersSamePrefixThenXx26
    TestSelectReferenceWorksForNoParenthesesPrefixOnlyTarget
    TestPrintAreaRepairUsesReferenceColumnsAndTargetLastRow
    TestPrintAreaRepairWorksForInvoiceSheet
    TestInvoiceFormulaRepairWritesExpectedFormulas
    TestInvoiceFormulaRepairDryRunReportsWithoutMutating
    TestInvoiceFormulaRepairUsesDynamicAmountCellForAbrLayout
    TestInvoiceFormulaRepairSkipsWhenNoAmountCellCanBeInferred
    TestPrintAreaRepairDoesNotExceedReferenceLastRow
    TestPrintAreaRepairFallsBackToKWithTargetLastRow
    TestPrintAreaRepairDryRunDoesNotChangeSheet
    TestPrintAreaRepairSkipsEmptySheet
    TestPrintAreaRepairRunsBeforeSignatureMerge
    TestExactReferenceMergeUsesOnlyProvidedAddress
    TestPrintAreaBoundaryStopsAtRightEdge
    TestSignatureFindKeepsNormalizedLabelMatching
    TestPrintAreaFallbackStopsAtK
    TestConflictSkipLeavesWorkbookUnmerged
    TestDryRunDoesNotSaveOrMerge
    Debug.Print "RepairDoMergesTests: all tests passed"
End Sub

Public Sub TestExtractOutletNameUsesLastParenthesizedPart()
    AssertEquals "Bedok Mall", RepairDoMerges_ExtractOutletName("MOS 0726 - 022 - DO & INV (Bedok Mall).xlsx"), "last outlet part"
    AssertEquals "West", RepairDoMerges_ExtractOutletName("KFP 0726 - 004 - DO & INV (Kebabs Concepts (West) - West Mall).xlsx"), "last simple parenthesized part"
    AssertEquals "", RepairDoMerges_ExtractOutletName("AK 0726 - 002.xlsx"), "AK without outlet parentheses has no match key"
    AssertEquals "Outlet", RepairDoMerges_ExtractOutletName("AK 0726 - 002 - DO & INV (Outlet).xlsx"), "AK with outlet parentheses can match reference"
End Sub

Public Sub TestNormalizeTextCollapsesWhitespace()
    AssertEquals "received in good order", RepairDoMerges_NormalizeText("  Received" & vbTab & "In  Good Order "), "normal text"
End Sub

Public Sub TestSelectReferencePrefersSamePrefixThenXx26()
    Dim candidates As Collection
    Set candidates = New Collection
    candidates.Add "C:\ref\deep\ABC 0726 - 001 - DO & INV (Outlet).xlsx"
    candidates.Add "C:\ref\ABC xx26 - 00x - DO & INV (Outlet).xlsx"
    candidates.Add "C:\ref\MOS xx26 - 00x - DO & INV (Outlet).xlsx"

    AssertEquals "C:\ref\ABC xx26 - 00x - DO & INV (Outlet).xlsx", RepairDoMerges_SelectReferencePath("C:\target\ABC 0726 - 001 - DO & INV (Outlet).xlsx", candidates), "same prefix and xx26"
End Sub

Public Sub TestSelectReferenceWorksForNoParenthesesPrefixOnlyTarget()
    Dim candidates As Collection
    Set candidates = New Collection
    candidates.Add "C:\ref\GPTG xx26 - 00x - DO & INV (Outlet).xlsx"
    candidates.Add "C:\ref\AK xx26 - 00x - DO & INV.xlsx"
    candidates.Add "C:\ref\AK 0726 - 001 - DO & INV.xlsx"

    AssertEquals "C:\ref\AK xx26 - 00x - DO & INV.xlsx", RepairDoMerges_SelectReferencePath("C:\target\AK 0726 - 002.xlsx", candidates), "no-parentheses target should use prefix and prefer xx26"
End Sub

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
    AssertEquals False, wb.Worksheets("DO").PageSetup.Zoom, "scale-to-fit should disable zoom"
    AssertEquals 1, wb.Worksheets("DO").PageSetup.FitToPagesWide, "print area should fit to one page wide"
    AssertEquals 1, wb.Worksheets("DO").PageSetup.FitToPagesTall, "print area should fit to one page high"

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

Public Sub TestPrintAreaRepairWorksForInvoiceSheet()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "Invoice"
    With wb.Worksheets("Invoice")
        .Range("D27").Value = "last invoice row"
        .PageSetup.PrintArea = "A1:X99"
    End With

    Dim changed As Boolean
    changed = RepairDoMerges_RepairPrintArea(wb.Worksheets("Invoice"), "A1:L60", True)

    AssertTrue changed, "invoice print area should change"
    AssertEquals "$A$1:$L$27", wb.Worksheets("Invoice").PageSetup.PrintArea, "invoice uses reference columns plus target last row"
    AssertEquals False, wb.Worksheets("Invoice").PageSetup.Zoom, "invoice scale-to-fit should disable zoom"
    AssertEquals 1, wb.Worksheets("Invoice").PageSetup.FitToPagesWide, "invoice should fit to one page wide"
    AssertEquals 1, wb.Worksheets("Invoice").PageSetup.FitToPagesTall, "invoice should fit to one page high"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 714, "TestPrintAreaRepairWorksForInvoiceSheet", message
End Sub

Public Sub TestInvoiceFormulaRepairWritesExpectedFormulas()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "Invoice"

    With wb.Worksheets("Invoice")
        .Range("A30").Value = "Subtotal"
        .Range("A31").Value = "GST 9%"
        .Range("A32").Value = "Total"
        .Range("I30").Formula = "=1"
        .Range("I31").Formula = "=1"
        .Range("I32").Formula = "=1"
    End With

    Dim changed As Boolean
    changed = RepairDoMerges_RepairInvoiceFormulas(wb.Worksheets("Invoice"), True)

    AssertTrue changed, "invoice formulas should report changed"
    AssertEquals "=SUM(I24:I29)", wb.Worksheets("Invoice").Range("I30").Formula, "subtotal formula"
    AssertEquals "=I30*0.09", wb.Worksheets("Invoice").Range("I31").Formula, "gst formula"
    AssertEquals "=SUM(I30:I31)", wb.Worksheets("Invoice").Range("I32").Formula, "total formula"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 717, "TestInvoiceFormulaRepairWritesExpectedFormulas", message
End Sub

Public Sub TestInvoiceFormulaRepairDryRunReportsWithoutMutating()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "Invoice"

    With wb.Worksheets("Invoice")
        .Range("A30").Value = "Subtotal"
        .Range("A31").Value = "GST 9%"
        .Range("A32").Value = "Total"
        .Range("I30").Formula = "=1"
        .Range("I31").Formula = "=2"
        .Range("I32").Formula = "=3"
    End With

    Dim changed As Boolean
    changed = RepairDoMerges_RepairInvoiceFormulas(wb.Worksheets("Invoice"), False)

    AssertTrue changed, "dry-run should report pending formula repair"
    AssertEquals "=1", wb.Worksheets("Invoice").Range("I30").Formula, "dry-run subtotal unchanged"
    AssertEquals "=2", wb.Worksheets("Invoice").Range("I31").Formula, "dry-run gst unchanged"
    AssertEquals "=3", wb.Worksheets("Invoice").Range("I32").Formula, "dry-run total unchanged"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 718, "TestInvoiceFormulaRepairDryRunReportsWithoutMutating", message
End Sub

Public Sub TestInvoiceFormulaRepairUsesDynamicAmountCellForAbrLayout()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "Invoice"

    With wb.Worksheets("Invoice")
        .Range("I34:J34").Merge
        .Range("I36:J36").Merge
        .Range("I38:J38").Merge
        .Range("K34:M34").Merge
        .Range("K36:M36").Merge
        .Range("K38:M38").Merge

        .Range("K26").Value = 10
        .Range("K32").Value = 20
        .Range("K33").Value = 0
        .Range("I34").Value = "Subtotal"
        .Range("I36").Value = "Add GST 9%"
        .Range("I38").Value = "Total"
        .Range("K34").Formula = "=1"
        .Range("K36").Formula = "=2"
        .Range("K38").Formula = "=3"
    End With

    Dim changed As Boolean
    changed = RepairDoMerges_RepairInvoiceFormulas(wb.Worksheets("Invoice"), True)

    AssertTrue changed, "ABR invoice formulas should report changed"
    AssertEquals "Subtotal", wb.Worksheets("Invoice").Range("I34").Value, "subtotal label should stay text"
    AssertEquals "Add GST 9%", wb.Worksheets("Invoice").Range("I36").Value, "gst label should stay text"
    AssertEquals "Total", wb.Worksheets("Invoice").Range("I38").Value, "total label should stay text"
    AssertEquals "=SUM(K26:M32)", wb.Worksheets("Invoice").Range("K34").Formula, "ABR subtotal formula"
    AssertEquals "=K34*0.09", wb.Worksheets("Invoice").Range("K36").Formula, "ABR gst formula"
    AssertEquals "=SUM(K33:M36)", wb.Worksheets("Invoice").Range("K38").Formula, "ABR total formula"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 719, "TestInvoiceFormulaRepairUsesDynamicAmountCellForAbrLayout", message
End Sub

Public Sub TestInvoiceFormulaRepairSkipsWhenNoAmountCellCanBeInferred()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "Invoice"

    With wb.Worksheets("Invoice")
        .Range("I34:J34").Merge
        .Range("I36:J36").Merge
        .Range("I38:J38").Merge
        .Range("I34").Value = "Subtotal"
        .Range("I36").Value = "Add GST 9%"
        .Range("I38").Value = "Total"
    End With

    Dim changed As Boolean
    changed = RepairDoMerges_RepairInvoiceFormulas(wb.Worksheets("Invoice"), True)

    AssertFalse changed, "missing amount cells should not report changed"
    AssertEquals "Subtotal", wb.Worksheets("Invoice").Range("I34").Value, "subtotal label should remain"
    AssertEquals "Add GST 9%", wb.Worksheets("Invoice").Range("I36").Value, "gst label should remain"
    AssertEquals "Total", wb.Worksheets("Invoice").Range("I38").Value, "total label should remain"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 720, "TestInvoiceFormulaRepairSkipsWhenNoAmountCellCanBeInferred", message
End Sub

Public Sub TestPrintAreaRepairDoesNotExceedReferenceLastRow()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"
    With wb.Worksheets("DO")
        .Range("A80").Value = "runaway content"
        .PageSetup.PrintArea = "A1:X99"
    End With

    Dim changed As Boolean
    changed = RepairDoMerges_RepairPrintArea(wb.Worksheets("DO"), "A1:K54", True)

    AssertTrue changed, "runaway print area should change"
    AssertEquals "$A$1:$K$54", wb.Worksheets("DO").PageSetup.PrintArea, "target print area must not exceed reference last row"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 716, "TestPrintAreaRepairDoesNotExceedReferenceLastRow", message
End Sub

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

Public Sub TestPrintAreaRepairSkipsEmptySheet()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"
    wb.Worksheets("DO").PageSetup.PrintArea = "A1:J10"

    Dim status As String
    Dim changed As Boolean
    changed = RepairDoMerges_RepairPrintArea(wb.Worksheets("DO"), "A1:K54", True, status)

    AssertFalse changed, "empty sheet should not change print area"
    AssertEquals "skipped-empty", status, "empty sheet status"
    AssertEquals "$A$1:$J$10", wb.Worksheets("DO").PageSetup.PrintArea, "empty sheet keeps existing print area"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 715, "TestPrintAreaRepairSkipsEmptySheet", message
End Sub

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

Public Sub TestPrintAreaBoundaryStopsAtRightEdge()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"
    With wb.Worksheets("DO")
        .PageSetup.PrintArea = "A1:J54"
        .Range("I50").Value = "Received In Good Order"
        .Range("H52").Value = "Authorised Signature & Stamp"
    End With

    Dim added As Collection
    Dim conflicts As Collection
    Set added = New Collection
    Set conflicts = New Collection

    RepairDoMerges_RepairSignatureLabels wb.Worksheets("DO"), True, added, conflicts

    AssertTrue RepairDoMerges_IsExactMerged(wb.Worksheets("DO"), "I50:J50"), "I50 should merge only through J50"
    AssertFalse RepairDoMerges_IsExactMerged(wb.Worksheets("DO"), "I50:K50"), "I50 must not merge through fallback K when print area ends at J"
    AssertTrue RepairDoMerges_IsExactMerged(wb.Worksheets("DO"), "H52:J52"), "H52 should merge through J52"
    AssertEquals "I50:J50", added.Item(1), "first signature range"
    AssertEquals "H52:J52", added.Item(2), "second signature range"
    AssertEquals xlCenter, wb.Worksheets("DO").Range("I50").HorizontalAlignment, "horizontal center"
    AssertEquals xlCenter, wb.Worksheets("DO").Range("I50").VerticalAlignment, "vertical center"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 701, "TestPrintAreaBoundaryStopsAtRightEdge", message
End Sub

Public Sub TestSignatureFindKeepsNormalizedLabelMatching()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"
    With wb.Worksheets("DO")
        .PageSetup.PrintArea = "A1:F20"
        .Range("C10").Value = "received" & vbTab & "in  good order"
        .Range("C12").Value = "received later by customer"
    End With

    Dim added As Collection
    Dim conflicts As Collection
    Set added = New Collection
    Set conflicts = New Collection

    RepairDoMerges_RepairSignatureLabels wb.Worksheets("DO"), True, added, conflicts

    AssertTrue RepairDoMerges_IsExactMerged(wb.Worksheets("DO"), "C10:F10"), "normalized received label should merge"
    AssertFalse wb.Worksheets("DO").Range("C12:F12").MergeCells, "non-label received text should not merge"
    AssertEquals 1, added.Count, "only exact normalized label should be added"
    AssertEquals "C10:F10", added.Item(1), "normalized signature range"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 709, "TestSignatureFindKeepsNormalizedLabelMatching", message
End Sub

Public Sub TestExactReferenceMergeUsesOnlyProvidedAddress()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"

    Dim added As Collection
    Dim conflicts As Collection
    Set added = New Collection
    Set conflicts = New Collection

    RepairDoMerges_AddMissingMerge wb.Worksheets("DO"), "A22:G23", True, added, conflicts

    AssertTrue RepairDoMerges_IsExactMerged(wb.Worksheets("DO"), "A22:G23"), "reference range should be merged exactly"
    AssertFalse RepairDoMerges_IsExactMerged(wb.Worksheets("DO"), "A22:K23"), "reference range must not extend to arbitrary K column"
    AssertEquals "A22:G23", added.Item(1), "exact reference range"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 707, "TestExactReferenceMergeUsesOnlyProvidedAddress", message
End Sub

Public Sub TestPrintAreaFallbackStopsAtK()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"
    wb.Worksheets("DO").Range("I50").Value = "Received In Good Order"

    Dim added As Collection
    Dim conflicts As Collection
    Set added = New Collection
    Set conflicts = New Collection

    RepairDoMerges_RepairSignatureLabels wb.Worksheets("DO"), True, added, conflicts

    AssertTrue RepairDoMerges_IsExactMerged(wb.Worksheets("DO"), "I50:K50"), "missing print area should fall back to K"
    AssertEquals "I50:K50", added.Item(1), "fallback signature range"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 708, "TestPrintAreaFallbackStopsAtK", message
End Sub

Public Sub TestConflictSkipLeavesWorkbookUnmerged()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"
    With wb.Worksheets("DO")
        .Range("H62").Value = "Received In Good Order"
        .Range("I62").Value = "AI keyed value"
    End With

    Dim added As Collection
    Dim conflicts As Collection
    Set added = New Collection
    Set conflicts = New Collection

    RepairDoMerges_AddMissingMerge wb.Worksheets("DO"), "H62:K62", True, added, conflicts

    AssertEquals 0, added.Count, "conflict should add nothing"
    AssertEquals 1, conflicts.Count, "conflict should be reported"
    AssertEquals "H62:K62", conflicts.Item(1), "conflict range"
    AssertFalse wb.Worksheets("DO").Range("H62:K62").MergeCells, "conflicting range should stay unmerged"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 702, "TestConflictSkipLeavesWorkbookUnmerged", message
End Sub

Public Sub TestDryRunDoesNotSaveOrMerge()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail

    wb.Worksheets(1).Name = "DO"

    Dim added As Collection
    Dim conflicts As Collection
    Set added = New Collection
    Set conflicts = New Collection

    RepairDoMerges_AddMissingMerge wb.Worksheets("DO"), "A7:K7", False, added, conflicts

    AssertEquals 1, added.Count, "dry run should report range"
    AssertEquals "A7:K7", added.Item(1), "dry run range"
    AssertFalse wb.Worksheets("DO").Range("A7:K7").MergeCells, "dry run should not merge"

CleanExit:
    wb.Close SaveChanges:=False
    Exit Sub
CleanFail:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 703, "TestDryRunDoesNotSaveOrMerge", message
End Sub

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal label As String)
    If CStr(expected) <> CStr(actual) Then
        Err.Raise vbObjectError + 704, "RepairDoMergesTests", label & ": expected [" & CStr(expected) & "], got [" & CStr(actual) & "]"
    End If
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then
        Err.Raise vbObjectError + 705, "RepairDoMergesTests", label
    End If
End Sub

Private Sub AssertFalse(ByVal condition As Boolean, ByVal label As String)
    If condition Then
        Err.Raise vbObjectError + 706, "RepairDoMergesTests", label
    End If
End Sub
