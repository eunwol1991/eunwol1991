Option Explicit

Private Const DEFAULT_TARGET_DIR As String = "C:\Users\jhunj\Dropbox\for jj\Doc to print - JJ"
Private Const DEFAULT_REFERENCE_DIR As String = "C:\Users\jhunj\Dropbox\DO & INV\DO & INV 2026"

Private Const SIGNATURE_END_COLUMN_FALLBACK As Long = 11
Private Const PRINT_AREA_START_ROW As Long = 1
Private Const PRINT_AREA_START_COLUMN As Long = 1
Private Const INVOICE_SUBTOTAL_START_ROW As Long = 24
Private Const AUTOMATION_SECURITY_FORCE_DISABLE As Long = 3

Private Const PAGE_MARGIN_INCHES As Double = 0.3
Private Const HEADER_FOOTER_MARGIN_INCHES As Double = 0.8

Public Sub RepairDoMergesDryRun()
    RepairDoMergesWithOptions DEFAULT_TARGET_DIR, DEFAULT_REFERENCE_DIR, False, False
End Sub

Public Sub RepairDoMergesApply()
    RepairDoMergesWithOptions DEFAULT_TARGET_DIR, DEFAULT_REFERENCE_DIR, True, False
End Sub

Public Sub RepairDoMergesWithOptions(ByVal targetDir As String, ByVal referenceDir As String, ByVal applyChanges As Boolean, Optional ByVal recursive As Boolean = False)
    Dim targetFiles As Collection
    Dim referenceIndex As Object
    Dim referencePrefixIndex As Object
    Dim referenceMergeCache As Object
    Dim referencePrintAreaCache As Object
    Dim missingReferenceDo As Object
    Dim targetPath As Variant
    Dim outletKey As String
    Dim prefixKey As String
    Dim referencePath As String
    Dim referenceMergeRanges As Collection
    Dim referencePrintAreas As Object
    Dim resultLine As String
    Dim index As Long
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim previousDisplayAlerts As Boolean
    Dim previousCalculation As XlCalculation

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousDisplayAlerts = Application.DisplayAlerts
    previousCalculation = Application.Calculation

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Debug.Print "Mode: " & IIf(applyChanges, "APPLY (saving changes)", "DRY RUN (preview only)")
    Debug.Print "Scanning target folder: " & FolderNameOnly(targetDir)

    Set targetFiles = RepairDoMerges_ListExcelFiles(targetDir, recursive)
    Debug.Print "Found " & targetFiles.Count & " target workbook(s)."

    Debug.Print "Indexing reference folder: " & FolderNameOnly(referenceDir)
    Set referenceIndex = RepairDoMerges_BuildReferenceIndex(referenceDir)
    Set referencePrefixIndex = RepairDoMerges_BuildReferencePrefixIndex(referenceDir)
    Debug.Print "Found " & ReferenceIndexCount(referenceIndex) & " reference workbook(s)."

    Set referenceMergeCache = CreateObject("Scripting.Dictionary")
    Set referencePrintAreaCache = CreateObject("Scripting.Dictionary")
    Set missingReferenceDo = CreateObject("Scripting.Dictionary")

    index = 0
    For Each targetPath In targetFiles
        index = index + 1
        Debug.Print "[" & index & "/" & targetFiles.Count & "] Checking " & FileNameOnly(CStr(targetPath))

        outletKey = RepairDoMerges_NormalizeText(RepairDoMerges_ExtractOutletName(CStr(targetPath)))
        prefixKey = RepairDoMerges_FilePrefixKey(CStr(targetPath))
        referencePath = ""

        If Len(outletKey) > 0 Then
            If referenceIndex.Exists(outletKey) Then
                referencePath = RepairDoMerges_SelectReferencePath(CStr(targetPath), referenceIndex.item(outletKey))
            End If
        End If

        If Len(referencePath) = 0 And Len(prefixKey) > 0 Then
            If referencePrefixIndex.Exists(prefixKey) Then
                referencePath = RepairDoMerges_SelectReferencePath(CStr(targetPath), referencePrefixIndex.item(prefixKey))
            End If
        End If

        Set referenceMergeRanges = Nothing
        Set referencePrintAreas = Nothing

        If Len(referencePath) = 0 Then
            Debug.Print "NO REFERENCE " & FileNameOnly(CStr(targetPath)) & ": page setup only"
        Else
            Set referenceMergeRanges = RepairDoMerges_GetCachedReferenceMergedRanges(referencePath, referenceMergeCache, missingReferenceDo)
            Set referencePrintAreas = RepairDoMerges_GetCachedReferencePrintAreas(referencePath, referencePrintAreaCache)
        End If

        resultLine = RepairDoMerges_RepairWorkbookWithReferenceRanges(CStr(targetPath), referencePath, referenceMergeRanges, referencePrintAreas, applyChanges)
        Debug.Print resultLine
    Next targetPath

CleanExit:
    Application.Calculation = previousCalculation
    Application.DisplayAlerts = previousDisplayAlerts
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    Dim message As String
    message = Err.Description

    On Error Resume Next
    Application.Calculation = previousCalculation
    Application.DisplayAlerts = previousDisplayAlerts
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0

    Err.Raise vbObjectError + 603, "RepairDoMergesWithOptions", message
End Sub

Public Function RepairDoMerges_RepairWorkbook(ByVal targetPath As String, ByVal referencePath As String, ByVal applyChanges As Boolean) As String
    Dim referenceMergeRanges As Collection
    Dim referencePrintAreas As Object

    If Len(referencePath) > 0 Then
        Set referenceMergeRanges = RepairDoMerges_LoadReferenceMergedRanges(referencePath)
        Set referencePrintAreas = RepairDoMerges_LoadReferencePrintAreas(referencePath)
    End If

    RepairDoMerges_RepairWorkbook = RepairDoMerges_RepairWorkbookWithReferenceRanges(targetPath, referencePath, referenceMergeRanges, referencePrintAreas, applyChanges)
End Function

Private Function RepairDoMerges_RepairWorkbookWithReferenceRanges(ByVal targetPath As String, ByVal referencePath As String, ByVal referenceMergeRanges As Collection, ByVal referencePrintAreas As Object, ByVal applyChanges As Boolean) As String
    Dim targetWb As Workbook
    Dim targetSheet As Worksheet
    Dim invoiceSheet As Worksheet
    Dim added As Collection
    Dim conflicts As Collection
    Dim mergeRange As Variant
    Dim printAreaChanged As Boolean
    Dim printAreaStatuses As Collection
    Dim doPrintAreaStatus As String
    Dim invoicePrintAreaStatus As String
    Dim invoiceFormulaChanged As Boolean
    Dim abrPoWrapTextChanged As Boolean
    Dim formatChanged As Boolean
    Dim viewChanged As Boolean
    Dim saved As Boolean
    Dim hasSupportedSheet As Boolean
    Dim referenceText As String

    Set added = New Collection
    Set conflicts = New Collection
    Set printAreaStatuses = New Collection

    On Error GoTo CleanFail

    Set targetWb = RepairDoMerges_OpenWorkbook(targetPath, Not applyChanges)

    hasSupportedSheet = WorkbookHasSheet(targetWb, "DO") Or WorkbookHasSheet(targetWb, "Invoice")
    If Not hasSupportedSheet Then
        RepairDoMerges_RepairWorkbookWithReferenceRanges = "SKIP " & FileNameOnly(targetPath) & ": missing DO and Invoice sheet"
        GoTo CleanExit
    End If

    If WorkbookHasSheet(targetWb, "DO") Then
        Set targetSheet = targetWb.Worksheets("DO")

        printAreaChanged = RepairDoMerges_RepairPrintArea(targetSheet, RepairDoMerges_ReferenceSheetPrintArea(referencePrintAreas, "DO"), applyChanges, doPrintAreaStatus)
        printAreaStatuses.Add "DO=" & doPrintAreaStatus

        If Not referenceMergeRanges Is Nothing Then
            For Each mergeRange In referenceMergeRanges
                RepairDoMerges_AddMissingMerge targetSheet, CStr(mergeRange), applyChanges, added, conflicts
            Next mergeRange

            RepairDoMerges_RepairSignatureLabels targetSheet, applyChanges, added, conflicts
        End If

        formatChanged = RepairDoMerges_RecalculateFormat(targetSheet, applyChanges)
        If applyChanges Then viewChanged = RepairDoMerges_SetPageBreakPreview(targetSheet)
    End If

    If WorkbookHasSheet(targetWb, "Invoice") Then
        Set invoiceSheet = targetWb.Worksheets("Invoice")

        invoiceFormulaChanged = RepairDoMerges_RepairInvoiceFormulas(invoiceSheet, applyChanges)
        printAreaChanged = RepairDoMerges_RepairPrintArea(invoiceSheet, RepairDoMerges_ReferenceSheetPrintArea(referencePrintAreas, "Invoice"), applyChanges, invoicePrintAreaStatus) Or printAreaChanged
        printAreaStatuses.Add "Invoice=" & invoicePrintAreaStatus

        formatChanged = RepairDoMerges_RecalculateFormat(invoiceSheet, applyChanges) Or formatChanged
        If applyChanges Then viewChanged = RepairDoMerges_SetPageBreakPreview(invoiceSheet) Or viewChanged
    End If

    abrPoWrapTextChanged = RepairDoMerges_RepairAbrPoWrapText(targetWb, targetPath, applyChanges)

    saved = ((printAreaChanged Or invoiceFormulaChanged Or abrPoWrapTextChanged Or added.Count > 0 Or formatChanged Or viewChanged) And applyChanges)
    If saved Then targetWb.Save

    If Len(referencePath) > 0 Then
        referenceText = FileNameOnly(referencePath)
    Else
        referenceText = "none"
    End If

    RepairDoMerges_RepairWorkbookWithReferenceRanges = FileNameOnly(targetPath) & _
        ": reference=" & referenceText & _
        "; printArea=" & CollectionToText(printAreaStatuses) & _
        "; invoiceFormula=" & CStr(invoiceFormulaChanged) & _
        "; abrPoWrapText=" & CStr(abrPoWrapTextChanged) & _
        "; add=" & added.Count & " " & CollectionToText(added) & _
        "; conflict=" & conflicts.Count & " " & CollectionToText(conflicts) & _
        "; format=" & CStr(formatChanged) & _
        "; view=" & CStr(viewChanged) & _
        "; saved=" & CStr(saved)

CleanExit:
    On Error Resume Next
    targetWb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

CleanFail:
    Dim failMessage As String
    failMessage = Err.Description

    On Error Resume Next
    targetWb.Close SaveChanges:=False
    On Error GoTo 0

    Err.Raise vbObjectError + 601, "RepairDoMerges_RepairWorkbook", failMessage
End Function

Private Function RepairDoMerges_RepairAbrPoWrapText(ByVal targetWb As Workbook, ByVal targetPath As String, ByVal applyChanges As Boolean) As Boolean
    Dim isAbr As Boolean
    Dim doRange As Range
    Dim invoiceRange As Range

    isAbr = RepairDoMerges_IsAbrWorkbook(targetWb, targetPath)
    If Not isAbr Then Exit Function

    If WorkbookHasSheet(targetWb, "DO") Then
        Set doRange = targetWb.Worksheets("DO").Range("J11:J12")
        If RepairDoMerges_SetWrapTextIfNeeded(doRange, applyChanges) Then
            RepairDoMerges_RepairAbrPoWrapText = True
        End If
    End If

    If WorkbookHasSheet(targetWb, "Invoice") Then
        Set invoiceRange = targetWb.Worksheets("Invoice").Range("K11:L12")
        If RepairDoMerges_SetWrapTextIfNeeded(invoiceRange, applyChanges) Then
            RepairDoMerges_RepairAbrPoWrapText = True
        End If
    End If
End Function

Private Function RepairDoMerges_IsAbrWorkbook(ByVal targetWb As Workbook, ByVal targetPath As String) As Boolean
    Dim fileName As String

    fileName = UCase$(FileNameOnly(targetPath))

    If Left$(fileName, 3) = "ABR" Then
        RepairDoMerges_IsAbrWorkbook = True
        Exit Function
    End If

    If WorkbookHasSheet(targetWb, "DO") Then
        If RepairDoMerges_RangeContainsAbrPo(targetWb.Worksheets("DO").Range("J11:J12")) Then
            RepairDoMerges_IsAbrWorkbook = True
            Exit Function
        End If
    End If

    If WorkbookHasSheet(targetWb, "Invoice") Then
        If RepairDoMerges_RangeContainsAbrPo(targetWb.Worksheets("Invoice").Range("K11:L12")) Then
            RepairDoMerges_IsAbrWorkbook = True
        End If
    End If
End Function

Private Function RepairDoMerges_RangeContainsAbrPo(ByVal targetRange As Range) As Boolean
    Dim cell As Range
    Dim valueText As String

    For Each cell In targetRange.Cells
        valueText = UCase$(Replace(Trim$(CStr(cell.Value2)), " ", ""))

        If Left$(valueText, 3) = "ABR" Then
            RepairDoMerges_RangeContainsAbrPo = True
            Exit Function
        End If
    Next cell
End Function

Private Function RepairDoMerges_SetWrapTextIfNeeded(ByVal targetRange As Range, ByVal applyChanges As Boolean) As Boolean
    Dim currentState As Variant
    Dim needsWrapText As Boolean

    currentState = targetRange.WrapText

    If IsNull(currentState) Then
        needsWrapText = True
    ElseIf CBool(currentState) = False Then
        needsWrapText = True
    End If

    If Not needsWrapText Then Exit Function

    RepairDoMerges_SetWrapTextIfNeeded = True

    If applyChanges Then
        targetRange.WrapText = True
    End If
End Function

Public Sub RepairDoMerges_AddMissingMerge(ByVal sheet As Worksheet, ByVal mergeAddress As String, ByVal applyChanges As Boolean, ByRef added As Collection, ByRef conflicts As Collection)
    If RepairDoMerges_IsExactMerged(sheet, mergeAddress) Then Exit Sub

    If RepairDoMerges_RangeHasNonAnchorValue(sheet.Range(mergeAddress)) Then
        RepairDoMerges_AddUnique conflicts, NormalizeAddress(sheet.Range(mergeAddress))
        Exit Sub
    End If

    RepairDoMerges_AddUnique added, NormalizeAddress(sheet.Range(mergeAddress))

    If applyChanges Then sheet.Range(mergeAddress).Merge
End Sub

Public Sub RepairDoMerges_RepairSignatureLabels(ByVal sheet As Worksheet, ByVal applyChanges As Boolean, ByRef added As Collection, ByRef conflicts As Collection)
    Dim endColumn As Long

    endColumn = RepairDoMerges_PrintAreaEndColumn(sheet)

    RepairDoMerges_FindAndRepairSignatureLabel sheet, "Received", "received in good order", endColumn, applyChanges, added, conflicts
    RepairDoMerges_FindAndRepairSignatureLabel sheet, "Authorised", "authorised signature & stamp", endColumn, applyChanges, added, conflicts
End Sub

Public Function RepairDoMerges_RepairPrintArea(ByVal targetSheet As Worksheet, ByVal referencePrintArea As String, ByVal applyChanges As Boolean, Optional ByRef status As String = "") As Boolean
    Dim endColumn As Long
    Dim endRow As Long
    Dim referenceEndRow As Long
    Dim desiredAddress As String
    Dim currentAddress As String
    Dim needsRepair As Boolean

    status = "unchanged"

    endColumn = RepairDoMerges_PrintAreaEndColumnForRepair(targetSheet, referencePrintArea)
    endRow = RepairDoMerges_LastContentRow(targetSheet)

    If endRow = 0 Then
        status = "skipped-empty"
        Exit Function
    End If

    referenceEndRow = RepairDoMerges_PrintAreaEndRowFromAddress(targetSheet, referencePrintArea)
    If referenceEndRow > 0 And endRow > referenceEndRow Then endRow = referenceEndRow

    desiredAddress = targetSheet.Range(targetSheet.Cells(PRINT_AREA_START_ROW, PRINT_AREA_START_COLUMN), targetSheet.Cells(endRow, endColumn)).address
    currentAddress = targetSheet.PageSetup.printArea

    needsRepair = False

    If RepairDoMerges_NormalizePrintAreaAddress(currentAddress) <> RepairDoMerges_NormalizePrintAreaAddress(desiredAddress) Then
        needsRepair = True
    End If

    If RepairDoMerges_PageSetupNeedsRepair(targetSheet) Then
        needsRepair = True
    End If

    If Not needsRepair Then Exit Function

    status = "repaired"
    RepairDoMerges_RepairPrintArea = True

    If applyChanges Then
        RepairDoMerges_ApplyRequiredPageSetup targetSheet, desiredAddress
    End If
End Function

Public Function RepairDoMerges_RepairInvoiceFormulas(ByVal invoiceSheet As Worksheet, ByVal applyChanges As Boolean) As Boolean
    Dim cell As Range
    Dim normalizedValue As String
    Dim subtotalLabel As Range
    Dim gstLabel As Range
    Dim totalLabel As Range
    Dim subtotalAmount As Range
    Dim gstAmount As Range
    Dim totalAmount As Range
    Dim subtotalStartRow As Long
    Dim totalStartRow As Long
    Dim subtotalFormula As String
    Dim gstFormula As String
    Dim totalFormula As String

    For Each cell In invoiceSheet.UsedRange.Cells
        If VarType(cell.value) = vbString Then
            normalizedValue = RepairDoMerges_NormalizeText(CStr(cell.value))

            If InStr(1, normalizedValue, "subtotal", vbTextCompare) > 0 Then
                Set subtotalLabel = cell
            ElseIf InStr(1, normalizedValue, "gst 9%", vbTextCompare) > 0 Then
                Set gstLabel = cell
            ElseIf InStr(1, normalizedValue, "total", vbTextCompare) > 0 Then
                Set totalLabel = cell
            End If
        End If
    Next cell

    If subtotalLabel Is Nothing Or gstLabel Is Nothing Or totalLabel Is Nothing Then Exit Function

    Set subtotalAmount = RepairDoMerges_InvoiceAmountCellForLabel(subtotalLabel)
    Set gstAmount = RepairDoMerges_InvoiceAmountCellForLabel(gstLabel)
    Set totalAmount = RepairDoMerges_InvoiceAmountCellForLabel(totalLabel)

    If subtotalAmount Is Nothing Or gstAmount Is Nothing Or totalAmount Is Nothing Then Exit Function

    subtotalStartRow = RepairDoMerges_FirstAmountRowAbove(subtotalAmount, subtotalLabel.Row)
    totalStartRow = RepairDoMerges_FirstAmountRowInBand(subtotalAmount, subtotalLabel.Row - 1, gstLabel.Row)
    If totalStartRow = 0 Then totalStartRow = subtotalLabel.Row

    subtotalFormula = "=SUM(" & RepairDoMerges_InvoiceAmountRangeAddress(subtotalAmount, subtotalStartRow, subtotalLabel.Row - 1) & ")"
    gstFormula = "=" & subtotalAmount.address(False, False) & "*0.09"
    totalFormula = "=SUM(" & RepairDoMerges_InvoiceAmountRangeAddress(totalAmount, totalStartRow, gstLabel.Row) & ")"

    If subtotalAmount.Formula <> subtotalFormula Then
        RepairDoMerges_RepairInvoiceFormulas = True
        If applyChanges Then subtotalAmount.Formula = subtotalFormula
    End If

    If gstAmount.Formula <> gstFormula Then
        RepairDoMerges_RepairInvoiceFormulas = True
        If applyChanges Then gstAmount.Formula = gstFormula
    End If

    If totalAmount.Formula <> totalFormula Then
        RepairDoMerges_RepairInvoiceFormulas = True
        If applyChanges Then totalAmount.Formula = totalFormula
    End If
End Function

Private Function RepairDoMerges_InvoiceAmountCellForLabel(ByVal labelCell As Range) As Range
    Dim sheet As Worksheet
    Dim labelArea As Range
    Dim used As Range
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim candidate As Range
    Dim candidateArea As Range

    Set sheet = labelCell.Worksheet
    Set labelArea = labelCell
    If labelCell.MergeCells Then Set labelArea = labelCell.MergeArea

    Set used = sheet.UsedRange
    lastColumn = used.Columns(used.Columns.Count).Column

    For columnIndex = labelArea.Column + labelArea.Columns.Count To lastColumn
        Set candidate = sheet.Cells(labelArea.Row, columnIndex)
        Set candidateArea = candidate
        If candidate.MergeCells Then Set candidateArea = candidate.MergeArea

        If candidate.Column = candidateArea.Column Then
            If RepairDoMerges_IsAmountCell(candidateArea.Cells(1, 1)) Then
                Set RepairDoMerges_InvoiceAmountCellForLabel = candidateArea.Cells(1, 1)
                Exit Function
            End If
        End If
    Next columnIndex
End Function

Private Function RepairDoMerges_IsAmountCell(ByVal cell As Range) As Boolean
    If Len(CStr(cell.Formula)) = 0 Then Exit Function
    If Left$(CStr(cell.Formula), 1) = "=" Then
        RepairDoMerges_IsAmountCell = True
    ElseIf VarType(cell.value) <> vbString Then
        RepairDoMerges_IsAmountCell = True
    End If
End Function

Private Function RepairDoMerges_FirstAmountRowAbove(ByVal amountCell As Range, ByVal labelRow As Long) As Long
    Dim foundRow As Long

    foundRow = RepairDoMerges_FirstAmountRowInBand(amountCell, INVOICE_SUBTOTAL_START_ROW, labelRow - 1)
    If foundRow > 0 Then
        RepairDoMerges_FirstAmountRowAbove = foundRow
    Else
        RepairDoMerges_FirstAmountRowAbove = INVOICE_SUBTOTAL_START_ROW
    End If
End Function

Private Function RepairDoMerges_FirstAmountRowInBand(ByVal amountCell As Range, ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim amountArea As Range
    Dim rowIndex As Long
    Dim columnIndex As Long

    If startRow < 1 Then startRow = 1
    If endRow < startRow Then Exit Function

    Set amountArea = amountCell
    If amountCell.MergeCells Then Set amountArea = amountCell.MergeArea

    For rowIndex = startRow To endRow
        For columnIndex = amountArea.Column To amountArea.Column + amountArea.Columns.Count - 1
            If RepairDoMerges_IsAmountCell(amountCell.Worksheet.Cells(rowIndex, columnIndex)) Then
                RepairDoMerges_FirstAmountRowInBand = rowIndex
                Exit Function
            End If
        Next columnIndex
    Next rowIndex
End Function

Private Function RepairDoMerges_InvoiceAmountRangeAddress(ByVal amountCell As Range, ByVal startRow As Long, ByVal endRow As Long) As String
    Dim amountArea As Range
    Dim firstCell As Range
    Dim lastCell As Range

    Set amountArea = amountCell
    If amountCell.MergeCells Then Set amountArea = amountCell.MergeArea

    Set firstCell = amountCell.Worksheet.Cells(startRow, amountArea.Column)
    Set lastCell = amountCell.Worksheet.Cells(endRow, amountArea.Column + amountArea.Columns.Count - 1)

    RepairDoMerges_InvoiceAmountRangeAddress = amountCell.Worksheet.Range(firstCell, lastCell).address(False, False)
End Function

Private Function RepairDoMerges_PageSetupNeedsRepair(ByVal targetSheet As Worksheet) As Boolean
    With targetSheet.PageSetup
        If .Orientation <> xlPortrait Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If

        If .Zoom <> False Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If

        If CLng(.FitToPagesWide) <> 1 Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If

        If CLng(.FitToPagesTall) <> 1 Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If

        If Abs(.TopMargin - Application.InchesToPoints(PAGE_MARGIN_INCHES)) > 0.1 Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If

        If Abs(.BottomMargin - Application.InchesToPoints(PAGE_MARGIN_INCHES)) > 0.1 Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If

        If Abs(.LeftMargin - Application.InchesToPoints(PAGE_MARGIN_INCHES)) > 0.1 Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If

        If Abs(.RightMargin - Application.InchesToPoints(PAGE_MARGIN_INCHES)) > 0.1 Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If

        If Abs(.HeaderMargin - Application.InchesToPoints(HEADER_FOOTER_MARGIN_INCHES)) > 0.1 Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If

        If Abs(.FooterMargin - Application.InchesToPoints(HEADER_FOOTER_MARGIN_INCHES)) > 0.1 Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If

        If .CenterHorizontally <> True Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If

        If .CenterVertically <> False Then
            RepairDoMerges_PageSetupNeedsRepair = True
            Exit Function
        End If
    End With
End Function

Private Sub RepairDoMerges_ApplyRequiredPageSetup(ByVal targetSheet As Worksheet, ByVal desiredPrintArea As String)
    With targetSheet.PageSetup
        .printArea = desiredPrintArea

        .Orientation = xlPortrait

        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1

        .TopMargin = Application.InchesToPoints(PAGE_MARGIN_INCHES)
        .BottomMargin = Application.InchesToPoints(PAGE_MARGIN_INCHES)
        .LeftMargin = Application.InchesToPoints(PAGE_MARGIN_INCHES)
        .RightMargin = Application.InchesToPoints(PAGE_MARGIN_INCHES)

        .HeaderMargin = Application.InchesToPoints(HEADER_FOOTER_MARGIN_INCHES)
        .FooterMargin = Application.InchesToPoints(HEADER_FOOTER_MARGIN_INCHES)

        .CenterHorizontally = True
        .CenterVertically = False
    End With
End Sub

Public Function RepairDoMerges_RecalculateFormat(ByVal targetSheet As Worksheet, ByVal applyChanges As Boolean) As Boolean
    Dim wasSaved As Boolean

    If Not applyChanges Then Exit Function

    wasSaved = targetSheet.Parent.saved
    targetSheet.Calculate

    RepairDoMerges_RecalculateFormat = (wasSaved And Not targetSheet.Parent.saved)
End Function

Public Function RepairDoMerges_SetPageBreakPreview(ByVal targetSheet As Worksheet) As Boolean
    targetSheet.Parent.Activate
    targetSheet.Activate

    If ActiveWindow.View <> xlPageBreakPreview Then
        ActiveWindow.View = xlPageBreakPreview
        RepairDoMerges_SetPageBreakPreview = True
    End If
End Function

Private Sub RepairDoMerges_FindAndRepairSignatureLabel(ByVal sheet As Worksheet, ByVal searchText As String, ByVal expectedLabel As String, ByVal endColumn As Long, ByVal applyChanges As Boolean, ByRef added As Collection, ByRef conflicts As Collection)
    Dim used As Range
    Dim found As Range
    Dim firstAddress As String

    Set used = sheet.UsedRange
    Set found = used.Find(What:=searchText, After:=used.Cells(used.Cells.Count), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

    If found Is Nothing Then Exit Sub

    firstAddress = found.address(False, False)

    Do
        If VarType(found.value) = vbString Then
            If RepairDoMerges_NormalizeText(CStr(found.value)) = expectedLabel Then
                RepairDoMerges_RepairSignatureCell sheet, found, endColumn, applyChanges, added, conflicts
            End If
        End If

        Set found = used.FindNext(found)
    Loop While Not found Is Nothing And found.address(False, False) <> firstAddress
End Sub

Private Sub RepairDoMerges_RepairSignatureCell(ByVal sheet As Worksheet, ByVal cell As Range, ByVal endColumn As Long, ByVal applyChanges As Boolean, ByRef added As Collection, ByRef conflicts As Collection)
    Dim mergeAddress As String
    Dim mergeRange As Range

    If cell.Column >= endColumn Then Exit Sub

    mergeAddress = sheet.Range(cell, sheet.Cells(cell.Row, endColumn)).address(False, False)
    Set mergeRange = sheet.Range(mergeAddress)

    If RepairDoMerges_IsExactMerged(sheet, mergeAddress) Then Exit Sub

    If RepairDoMerges_RangeHasNonAnchorValue(mergeRange) Then
        RepairDoMerges_AddUnique conflicts, NormalizeAddress(mergeRange)
    Else
        RepairDoMerges_AddUnique added, NormalizeAddress(mergeRange)

        If applyChanges Then
            mergeRange.Merge
            With cell
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End If
    End If
End Sub

Public Function RepairDoMerges_ExtractOutletName(ByVal fileNameOrPath As String) As String
    Dim nameOnly As String
    Dim scanIndex As Long
    Dim closePos As Long
    Dim openPos As Long
    Dim candidate As String

    nameOnly = FileNameOnly(fileNameOrPath)
    scanIndex = Len(nameOnly)

    Do While scanIndex > 0
        closePos = InStrRev(nameOnly, ")", scanIndex, vbBinaryCompare)
        If closePos = 0 Then Exit Do

        openPos = InStrRev(Left$(nameOnly, closePos - 1), "(", -1, vbBinaryCompare)
        If openPos = 0 Then Exit Do

        candidate = Mid$(nameOnly, openPos + 1, closePos - openPos - 1)

        If InStr(1, candidate, "(", vbBinaryCompare) = 0 And InStr(1, candidate, ")", vbBinaryCompare) = 0 Then
            RepairDoMerges_ExtractOutletName = Trim$(candidate)
            Exit Function
        End If

        scanIndex = closePos - 1
    Loop

    RepairDoMerges_ExtractOutletName = ""
End Function

Public Function RepairDoMerges_NormalizeText(ByVal text As String) As String
    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\s+"

    RepairDoMerges_NormalizeText = LCase$(re.Replace(Trim$(text), " "))
End Function

Public Function RepairDoMerges_SelectReferencePath(ByVal targetPath As String, ByVal candidates As Collection) As String
    Dim filtered As Collection
    Dim candidate As Variant
    Dim targetPrefix As String

    If candidates.Count = 0 Then Exit Function

    targetPrefix = LCase$(FirstFilenameToken(FileNameOnly(targetPath)))

    Set filtered = New Collection

    For Each candidate In candidates
        If LCase$(Left$(FileNameOnly(CStr(candidate)), Len(targetPrefix))) = targetPrefix Then
            filtered.Add CStr(candidate)
        End If
    Next candidate

    If filtered.Count = 0 Then Set filtered = candidates

    Set filtered = PreferContaining(filtered, "xx26")
    RepairDoMerges_SelectReferencePath = StableShortestReference(filtered)
End Function

Public Function RepairDoMerges_ListExcelFiles(ByVal folderPath As String, ByVal recursive As Boolean) As Collection
    Dim files As Collection

    Set files = New Collection
    AddExcelFiles files, folderPath, recursive

    Set RepairDoMerges_ListExcelFiles = SortPathCollection(files)
End Function

Public Function RepairDoMerges_BuildReferenceIndex(ByVal referenceDir As String) As Object
    Dim index As Object
    Dim files As Collection
    Dim path As Variant
    Dim outletKey As String

    Set index = CreateObject("Scripting.Dictionary")
    Set files = RepairDoMerges_ListExcelFiles(referenceDir, True)

    For Each path In files
        outletKey = RepairDoMerges_NormalizeText(RepairDoMerges_ExtractOutletName(CStr(path)))

        If Len(outletKey) > 0 Then
            If Not index.Exists(outletKey) Then index.Add outletKey, New Collection
            index.item(outletKey).Add CStr(path)
        End If
    Next path

    Set RepairDoMerges_BuildReferenceIndex = index
End Function

Public Function RepairDoMerges_BuildReferencePrefixIndex(ByVal referenceDir As String) As Object
    Dim index As Object
    Dim files As Collection
    Dim path As Variant
    Dim prefixKey As String

    Set index = CreateObject("Scripting.Dictionary")
    Set files = RepairDoMerges_ListExcelFiles(referenceDir, True)

    For Each path In files
        prefixKey = RepairDoMerges_FilePrefixKey(CStr(path))

        If Len(prefixKey) > 0 Then
            If Not index.Exists(prefixKey) Then index.Add prefixKey, New Collection
            index.item(prefixKey).Add CStr(path)
        End If
    Next path

    Set RepairDoMerges_BuildReferencePrefixIndex = index
End Function

Public Function RepairDoMerges_FilePrefixKey(ByVal fileNameOrPath As String) As String
    RepairDoMerges_FilePrefixKey = RepairDoMerges_NormalizeText(FirstFilenameToken(FileNameOnly(fileNameOrPath)))
End Function

Private Function RepairDoMerges_GetCachedReferenceMergedRanges(ByVal referencePath As String, ByVal referenceMergeCache As Object, ByVal missingReferenceDo As Object) As Collection
    If missingReferenceDo.Exists(referencePath) Then Exit Function

    If Not referenceMergeCache.Exists(referencePath) Then
        Dim mergeRanges As Collection
        Set mergeRanges = RepairDoMerges_LoadReferenceMergedRanges(referencePath)

        If mergeRanges Is Nothing Then
            missingReferenceDo.Add referencePath, True
        Else
            referenceMergeCache.Add referencePath, mergeRanges
        End If
    End If

    If referenceMergeCache.Exists(referencePath) Then
        Set RepairDoMerges_GetCachedReferenceMergedRanges = referenceMergeCache.item(referencePath)
    End If
End Function

Private Function RepairDoMerges_GetCachedReferencePrintAreas(ByVal referencePath As String, ByVal referencePrintAreaCache As Object) As Object
    If Not referencePrintAreaCache.Exists(referencePath) Then
        referencePrintAreaCache.Add referencePath, RepairDoMerges_LoadReferencePrintAreas(referencePath)
    End If

    Set RepairDoMerges_GetCachedReferencePrintAreas = referencePrintAreaCache.item(referencePath)
End Function

Private Function RepairDoMerges_LoadReferenceMergedRanges(ByVal referencePath As String) As Collection
    Dim referenceWb As Workbook
    Dim referenceSheet As Worksheet

    On Error GoTo CleanFail

    Set referenceWb = RepairDoMerges_OpenWorkbook(referencePath, True)

    If WorkbookHasSheet(referenceWb, "DO") Then
        Set referenceSheet = referenceWb.Worksheets("DO")
        Set RepairDoMerges_LoadReferenceMergedRanges = RepairDoMerges_GetMergedRanges(referenceSheet)
    End If

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

    Err.Raise vbObjectError + 604, "RepairDoMerges_LoadReferenceMergedRanges", message
End Function

Private Function RepairDoMerges_LoadReferencePrintAreas(ByVal referencePath As String) As Object
    Dim referenceWb As Workbook
    Dim printAreas As Object

    Set printAreas = CreateObject("Scripting.Dictionary")

    On Error GoTo CleanFail

    Set referenceWb = RepairDoMerges_OpenWorkbook(referencePath, True)

    If WorkbookHasSheet(referenceWb, "DO") Then
        printAreas.Add "DO", referenceWb.Worksheets("DO").PageSetup.printArea
    End If

    If WorkbookHasSheet(referenceWb, "Invoice") Then
        printAreas.Add "Invoice", referenceWb.Worksheets("Invoice").PageSetup.printArea
    End If

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

Private Function RepairDoMerges_ReferenceSheetPrintArea(ByVal referencePrintAreas As Object, ByVal sheetName As String) As String
    If referencePrintAreas Is Nothing Then Exit Function

    If referencePrintAreas.Exists(sheetName) Then
        RepairDoMerges_ReferenceSheetPrintArea = CStr(referencePrintAreas.item(sheetName))
    End If
End Function

Private Function RepairDoMerges_OpenWorkbook(ByVal workbookPath As String, ByVal readOnlyMode As Boolean) As Workbook
    Dim previousAutomationSecurity As Long

    previousAutomationSecurity = Application.AutomationSecurity

    On Error GoTo CleanFail

    Application.AutomationSecurity = AUTOMATION_SECURITY_FORCE_DISABLE
    Set RepairDoMerges_OpenWorkbook = Workbooks.Open(fileName:=workbookPath, UpdateLinks:=0, ReadOnly:=readOnlyMode)

CleanExit:
    Application.AutomationSecurity = previousAutomationSecurity
    Exit Function

CleanFail:
    Dim message As String
    message = Err.Description

    On Error Resume Next
    Application.AutomationSecurity = previousAutomationSecurity
    On Error GoTo 0

    Err.Raise vbObjectError + 605, "RepairDoMerges_OpenWorkbook", message
End Function

Public Function RepairDoMerges_GetMergedRanges(ByVal sheet As Worksheet) As Collection
    Dim ranges As Collection
    Dim seen As Object
    Dim cell As Range
    Dim address As String

    Set ranges = New Collection
    Set seen = CreateObject("Scripting.Dictionary")

    For Each cell In sheet.UsedRange.Cells
        If cell.MergeCells Then
            address = NormalizeAddress(cell.MergeArea)

            If Not seen.Exists(address) Then
                seen.Add address, True
                ranges.Add address
            End If
        End If
    Next cell

    Set RepairDoMerges_GetMergedRanges = SortTextCollection(ranges)
End Function

Public Function RepairDoMerges_IsExactMerged(ByVal sheet As Worksheet, ByVal mergeAddress As String) As Boolean
    Dim target As Range
    Dim mergeState As Variant

    Set target = sheet.Range(mergeAddress)
    mergeState = target.MergeCells

    If IsNull(mergeState) Then Exit Function
    If Not CBool(mergeState) Then Exit Function

    RepairDoMerges_IsExactMerged = (NormalizeAddress(target.Cells(1, 1).MergeArea) = NormalizeAddress(target))
End Function

Public Function RepairDoMerges_RangeHasNonAnchorValue(ByVal targetRange As Range) As Boolean
    Dim cell As Range

    For Each cell In targetRange.Cells
        If cell.Row <> targetRange.Row Or cell.Column <> targetRange.Column Then
            If Len(CStr(cell.value)) > 0 Then
                RepairDoMerges_RangeHasNonAnchorValue = True
                Exit Function
            End If
        End If
    Next cell
End Function

Public Function RepairDoMerges_PrintAreaEndColumn(ByVal sheet As Worksheet) As Long
    RepairDoMerges_PrintAreaEndColumn = RepairDoMerges_PrintAreaEndColumnFromAddress(sheet, sheet.PageSetup.printArea)
End Function

Private Function RepairDoMerges_PrintAreaEndColumnForRepair(ByVal sheet As Worksheet, ByVal referencePrintArea As String) As Long
    Dim endColumn As Long

    endColumn = RepairDoMerges_PrintAreaEndColumnFromAddress(sheet, referencePrintArea)

    If Len(Trim$(referencePrintArea)) = 0 Then
        If Len(Trim$(sheet.PageSetup.printArea)) > 0 Then
            endColumn = RepairDoMerges_PrintAreaEndColumnFromAddress(sheet, sheet.PageSetup.printArea)
        End If
    End If

    If endColumn <= 0 Then endColumn = SIGNATURE_END_COLUMN_FALLBACK

    RepairDoMerges_PrintAreaEndColumnForRepair = endColumn
End Function

Public Function RepairDoMerges_PrintAreaEndColumnFromAddress(ByVal sheet As Worksheet, ByVal printArea As String) As Long
    Dim firstArea As String
    Dim bangPos As Long
    Dim rangeAddress As String

    printArea = Trim$(printArea)

    If Len(printArea) = 0 Then
        RepairDoMerges_PrintAreaEndColumnFromAddress = SIGNATURE_END_COLUMN_FALLBACK
        Exit Function
    End If

    firstArea = Split(printArea, ",")(0)
    bangPos = InStrRev(firstArea, "!")

    If bangPos > 0 Then firstArea = Mid$(firstArea, bangPos + 1)

    rangeAddress = Replace(Replace(firstArea, "'", ""), "$", "")

    On Error GoTo UseFallback

    RepairDoMerges_PrintAreaEndColumnFromAddress = sheet.Range(rangeAddress).Columns(sheet.Range(rangeAddress).Columns.Count).Column
    Exit Function

UseFallback:
    RepairDoMerges_PrintAreaEndColumnFromAddress = SIGNATURE_END_COLUMN_FALLBACK
End Function

Public Function RepairDoMerges_PrintAreaEndRowFromAddress(ByVal sheet As Worksheet, ByVal printArea As String) As Long
    Dim firstArea As String
    Dim bangPos As Long
    Dim rangeAddress As String

    printArea = Trim$(printArea)

    If Len(printArea) = 0 Then Exit Function

    firstArea = Split(printArea, ",")(0)
    bangPos = InStrRev(firstArea, "!")

    If bangPos > 0 Then firstArea = Mid$(firstArea, bangPos + 1)

    rangeAddress = Replace(Replace(firstArea, "'", ""), "$", "")

    On Error GoTo UseFallback

    RepairDoMerges_PrintAreaEndRowFromAddress = sheet.Range(rangeAddress).Rows(sheet.Range(rangeAddress).Rows.Count).Row
    Exit Function

UseFallback:
    RepairDoMerges_PrintAreaEndRowFromAddress = 0
End Function

Public Function RepairDoMerges_LastContentRow(ByVal sheet As Worksheet) As Long
    Dim found As Range

    Set found = sheet.Cells.Find(What:="*", After:=sheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)

    If Not found Is Nothing Then
        RepairDoMerges_LastContentRow = found.Row
    End If
End Function

Private Function RepairDoMerges_NormalizePrintAreaAddress(ByVal printArea As String) As String
    RepairDoMerges_NormalizePrintAreaAddress = UCase$(Replace(Replace(Trim$(printArea), "'", ""), "$", ""))
End Function

Public Sub RepairDoMerges_AddUnique(ByRef values As Collection, ByVal value As String)
    Dim item As Variant

    For Each item In values
        If CStr(item) = value Then Exit Sub
    Next item

    values.Add value
End Sub

Private Sub AddExcelFiles(ByRef files As Collection, ByVal folderPath As String, ByVal recursive As Boolean)
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim subFolder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        Err.Raise vbObjectError + 602, "RepairDoMerges", "Folder not found: " & folderPath
    End If

    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.files
        If IsExcelWorkbookName(file.Name) Then files.Add file.path
    Next file

    If recursive Then
        For Each subFolder In folder.SubFolders
            AddExcelFiles files, subFolder.path, True
        Next subFolder
    End If
End Sub

Private Function IsExcelWorkbookName(ByVal fileName As String) As Boolean
    Dim lowerName As String

    lowerName = LCase$(fileName)

    If Left$(fileName, 2) = "~$" Then Exit Function

    If Right$(lowerName, 5) = ".xlsx" Then
        IsExcelWorkbookName = True
    ElseIf Right$(lowerName, 5) = ".xlsm" Then
        IsExcelWorkbookName = True
    ElseIf Right$(lowerName, 4) = ".xls" Then
        IsExcelWorkbookName = True
    End If
End Function

Private Function PreferContaining(ByVal candidates As Collection, ByVal needle As String) As Collection
    Dim matches As Collection
    Dim candidate As Variant

    Set matches = New Collection

    For Each candidate In candidates
        If InStr(1, LCase$(FileNameOnly(CStr(candidate))), LCase$(needle), vbTextCompare) > 0 Then
            matches.Add CStr(candidate)
        End If
    Next candidate

    If matches.Count > 0 Then
        Set PreferContaining = matches
    Else
        Set PreferContaining = candidates
    End If
End Function

Private Function StableShortestReference(ByVal candidates As Collection) As String
    Dim candidate As Variant
    Dim best As String

    For Each candidate In candidates
        If Len(best) = 0 Then
            best = CStr(candidate)
        ElseIf PathDepth(CStr(candidate)) < PathDepth(best) Then
            best = CStr(candidate)
        ElseIf PathDepth(CStr(candidate)) = PathDepth(best) Then
            If LCase$(FileNameOnly(CStr(candidate))) < LCase$(FileNameOnly(best)) Then
                best = CStr(candidate)
            End If
        End If
    Next candidate

    StableShortestReference = best
End Function

Private Function SortPathCollection(ByVal values As Collection) As Collection
    Dim sorted As Collection
    Dim i As Long
    Dim j As Long
    Dim temp() As String

    ReDim temp(1 To IIf(values.Count = 0, 1, values.Count))

    For i = 1 To values.Count
        temp(i) = CStr(values.item(i))
    Next i

    For i = 1 To values.Count - 1
        For j = i + 1 To values.Count
            If LCase$(temp(j)) < LCase$(temp(i)) Then
                SwapStrings temp(i), temp(j)
            End If
        Next j
    Next i

    Set sorted = New Collection

    For i = 1 To values.Count
        sorted.Add temp(i)
    Next i

    Set SortPathCollection = sorted
End Function

Private Function SortTextCollection(ByVal values As Collection) As Collection
    Set SortTextCollection = SortPathCollection(values)
End Function

Private Sub SwapStrings(ByRef leftValue As String, ByRef rightValue As String)
    Dim temp As String

    temp = leftValue
    leftValue = rightValue
    rightValue = temp
End Sub

Private Function WorkbookHasSheet(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet

    For Each sheet In wb.Worksheets
        If sheet.Name = sheetName Then
            WorkbookHasSheet = True
            Exit Function
        End If
    Next sheet
End Function

Private Function CollectionToText(ByVal values As Collection) As String
    Dim item As Variant
    Dim result As String

    result = "["

    For Each item In values
        If Len(result) > 1 Then result = result & ", "
        result = result & CStr(item)
    Next item

    CollectionToText = result & "]"
End Function

Private Function ReferenceIndexCount(ByVal index As Object) As Long
    Dim key As Variant

    For Each key In index.Keys
        ReferenceIndexCount = ReferenceIndexCount + index.item(key).Count
    Next key
End Function

Private Function FirstFilenameToken(ByVal fileName As String) As String
    FirstFilenameToken = Split(Trim$(fileName), " ")(0)
End Function

Private Function FileNameOnly(ByVal fullPath As String) As String
    Dim slashPos As Long

    slashPos = InStrRev(fullPath, "\")

    If slashPos = 0 Then slashPos = InStrRev(fullPath, "/")

    If slashPos = 0 Then
        FileNameOnly = fullPath
    Else
        FileNameOnly = Mid$(fullPath, slashPos + 1)
    End If
End Function

Private Function FolderNameOnly(ByVal fullPath As String) As String
    Dim trimmed As String
    Dim slashPos As Long

    trimmed = fullPath

    Do While Right$(trimmed, 1) = "\" Or Right$(trimmed, 1) = "/"
        trimmed = Left$(trimmed, Len(trimmed) - 1)
    Loop

    slashPos = InStrRev(trimmed, "\")

    If slashPos = 0 Then slashPos = InStrRev(trimmed, "/")

    If slashPos = 0 Then
        FolderNameOnly = trimmed
    Else
        FolderNameOnly = Mid$(trimmed, slashPos + 1)
    End If
End Function

Private Function PathDepth(ByVal fullPath As String) As Long
    Dim normalized As String
    Dim index As Long

    normalized = Replace(fullPath, "/", "\")

    For index = 1 To Len(normalized)
        If Mid$(normalized, index, 1) = "\" Then
            PathDepth = PathDepth + 1
        End If
    Next index
End Function

Private Function NormalizeAddress(ByVal targetRange As Range) As String
    NormalizeAddress = targetRange.address(False, False)
End Function



