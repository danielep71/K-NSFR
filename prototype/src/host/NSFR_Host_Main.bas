Attribute VB_Name = "NSFR_Host_Main"
'==============================================================================
' MODULE: NSFR_Host_Main
'------------------------------------------------------------------------------
' PURPOSE
'   Workbook glue of the prototype: the macros of the sheet buttons. It
'   reads Params, Positions and BankCells into arrays, calls NSFR_API_Run
'   and writes Results, C80, C81, PositionCalc, Reconciliation and Checks;
'   it imports a bank's figures from its reporting workbook or CSV export.
'
' PUBLIC SURFACE
'   Buttons: RunNSFR, ImportBankWorkbook, ImportBankCsv, ClearOutputs,
'   ClearBankCells. Silent variants for the regression harness:
'   RunNSFR_Silent, ImportBankWorkbookFrom, ImportBankCsvFrom.
'
' DEPENDENCIES
'   NSFR_API_Functions; NSFR_Core_Config for the layout; NSFR_Core_Util and
'   NSFR_Core_Import for conversion and parsing. Workbook names created by
'   tools/build_template.py locate every table (hdr*) and setting.
'
' STATE OWNERSHIP
'   Changes ScreenUpdating, Calculation and EnableEvents during an
'   operation and restores the captured values on success and failure;
'   leaves the output sheets unprotected (PROTOTYPE.md, Departures).
'
' ERROR POLICY
'   A narrow On Error handler per public macro captures the primary error,
'   restores the application state, then reports it (MsgBox, or the
'   returned text of a silent variant).
'
' COMPATIBILITY
'   Excel VBA7, 64-bit, Windows. No references beyond the defaults.
'
' UPDATED
'   2026-11-05
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const APP_TITLE         As String = "K-NSFR"
        Private Const ERR_HOST          As Long = vbObjectError + 2101
        Private Const SH_RESULTS        As String = "Results"
        Private Const SH_C80            As String = "C80"
        Private Const SH_C81            As String = "C81"
        Private Const SH_POSCALC        As String = "PositionCalc"
        Private Const SH_RECON          As String = "Reconciliation"
        Private Const SH_CHECKS         As String = "Checks"
        Private Const SH_BANK           As String = "BankCells"
        Private Const MAX_ROWS          As Long = 20000  'Rows cleared below an output header

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Application settings captured by BeginOperation for EndOperation.
        Private mSavedScreen     As Boolean
        Private mSavedEvents     As Boolean
        Private mSavedCalc       As Long
        Private mStateCaptured   As Boolean

'------------------------------------------------------------------------------
'                                   BUTTONS
'------------------------------------------------------------------------------
Public Sub RunNSFR()
'
'   PURPOSE    Button "Run benchmark": run and report the outcome.
'   UPDATED    2026-11-05
'
    Dim Outcome         As String      'Result text

        Outcome = RunNSFR_Silent()
        If Left$(Outcome, 3) = "OK:" Then
            MsgBox Mid$(Outcome, 5), vbInformation, APP_TITLE
        Else
            MsgBox Outcome, vbExclamation, APP_TITLE
        End If

End Sub


Public Function RunNSFR_Silent() As String
'
'   PURPOSE    Run the benchmark and the reconciliation, write every output
'              sheet and return the outcome as text.
'   RETURNS    "OK: <summary>" when the run completed (its status may still
'              be INVALID or STOPPED, as Results shows); "FAIL: <error>"
'              when a host error stopped it.
'   UPDATED    2026-11-05
'
    Dim Factors()       As Variant     'Params factor table
    Dim Categories()    As Variant     'Params category table
    Dim Summary()       As Variant     'Params summary map
    Dim Positions()     As Variant     'Positions
    Dim BankCells()     As Variant     'Bank figures
    Dim BankSettings(1 To 4) As Variant 'AmountScale and tolerances
    Dim FactorRows      As Long
    Dim CategoryRows    As Long
    Dim SummaryRows     As Long
    Dim PositionRows    As Long
    Dim BankRows        As Long
    Dim RowAmounts()    As Double      'Benchmark amount per row and bucket
    Dim RowStable()     As Double      'Benchmark stable funding per row and bucket
    Dim PositionOut()   As Variant     'PositionCalc rows
    Dim SummaryOut()    As Variant     'C 84.00 values
    Dim Totals()        As Variant     'ASF, RSF, NSFR
    Dim Counts()        As Long        'Accepted, rejected, overrides
    Dim RunClass        As String      'REFERENCE or WHAT-IF
    Dim Lines           As Variant     'Reconciliation lines
    Dim LineCount       As Long
    Dim Headline        As Variant     'Attribution
    Dim HasBank         As Boolean     'Bank figures usable
    Dim Diag()          As Variant     'Diagnostics list
    Dim DiagCount       As Long
    Dim Status          As String      'Run status
    Dim ErrNumber       As Long
    Dim ErrDescription  As String

        On Error GoTo Failed
        BeginOperation
        FactorRows = ReadTable("hdrFactors", FC_COLUMNS, Factors, True)
        CategoryRows = ReadTable("hdrCategories", CC_COLUMNS, Categories, True)
        SummaryRows = ReadTable("hdrSummary", SC_COLUMNS, Summary, True)
        PositionRows = ReadTable("hdrPositions", PC_COLUMNS, Positions, False)
        BankRows = ReadTable("hdrBankCells", BC_COLUMNS, BankCells, False)
        BankSettings(1) = NamedCell("AmountScale").Value
        BankSettings(2) = NamedCell("AbsTolerance").Value
        BankSettings(3) = NamedCell("RelTolerance").Value
        BankSettings(4) = NamedCell("FactorTolerance").Value

        Status = NSFR_API_Run(TextOf(NamedCell("ProfileID").Value), NamedCell("ReportingDate").Value, _
                              Factors, FactorRows, Categories, CategoryRows, Summary, SummaryRows, _
                              Positions, PositionRows, BankCells, BankRows, BankSettings, _
                              RowAmounts, RowStable, PositionOut, SummaryOut, Totals, RunClass, Counts, _
                              Lines, LineCount, Headline, HasBank, Diag, DiagCount)

        ClearOutputSheets
        WriteResults Status, RunClass, Totals, Counts, SummaryOut, SummaryRows, DiagCount, Diag
        If Left$(Status, Len(STATUS_STOPPED)) <> STATUS_STOPPED Then
            WriteTemplates Factors, FactorRows, RowAmounts, RowStable, (Status = STATUS_VALID)
            If PositionRows > 0 Then WriteBlock SH_POSCALC, "hdrPositionCalc", PositionOut, PositionRows, PO_COLUMNS
            If HasBank Then WriteReconciliation Lines, LineCount, Headline
        End If
        If DiagCount > 0 Then WriteBlock SH_CHECKS, "hdrChecks", ListToMatrix(Diag, DiagCount, DG_COLUMNS), _
                                         DiagCount, DG_COLUMNS
        RunNSFR_Silent = "OK: " & Status & " | " & RunClass & " | errors " & CountSeverity(Diag, DiagCount, SEV_ERROR) _
                         & ", warnings " & CountSeverity(Diag, DiagCount, SEV_WARNING) _
                         & IIf(HasBank, " | reconciliation lines " & LineCount, " | no bank figures")

CleanUp:
        EndOperation
        If ErrNumber <> 0 Then RunNSFR_Silent = "FAIL: " & ErrDescription & " (" & ErrNumber & ")"
        Exit Function
Failed:
        ErrNumber = Err.Number
        ErrDescription = Err.Description
        Resume CleanUp
End Function


Public Sub ClearOutputs()
'
'   PURPOSE    Button "Clear outputs": empty every output sheet.
'   UPDATED    2026-11-05
'
    Dim ErrNumber       As Long
    Dim ErrDescription  As String

        On Error GoTo Failed
        BeginOperation
        ClearOutputSheets
CleanUp:
        EndOperation
        If ErrNumber <> 0 Then MsgBox "Clear stopped: " & ErrDescription, vbExclamation, APP_TITLE
        Exit Sub
Failed:
        ErrNumber = Err.Number
        ErrDescription = Err.Description
        Resume CleanUp
End Sub


Public Sub ClearBankCells()
'
'   PURPOSE    Button "Clear bank figures": empty the BankCells table.
'   UPDATED    2026-11-05
'
        If MsgBox("Clear every bank figure on BankCells?", vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub
        ClearBelow NamedCell("hdrBankCells"), BC_COLUMNS

End Sub

'------------------------------------------------------------------------------
'                                    IMPORT
'------------------------------------------------------------------------------
Public Sub ImportBankWorkbook()
'
'   PURPOSE    Button "Import bank workbook": ask for the bank's reporting
'              workbook and import its C 80.00, C 81.00 and C 84.00 sheets.
'   UPDATED    2026-11-05
'
    Dim Path            As Variant     'Chosen file, or False

        Path = Application.GetOpenFilename("Excel workbooks (*.xls*),*.xls*", , "Bank NSFR reporting workbook")
        If VarType(Path) = vbBoolean Then Exit Sub
        ReportImport ImportBankWorkbookFrom(CStr(Path))

End Sub


Public Sub ImportBankCsv()
'
'   PURPOSE    Button "Import bank CSV": ask for a CSV export
'              (RowColumn_Index,Suffix,Amount) and import it.
'   UPDATED    2026-11-05
'
    Dim Path            As Variant     'Chosen file, or False

        Path = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Bank NSFR CSV export")
        If VarType(Path) = vbBoolean Then Exit Sub
        ReportImport ImportBankCsvFrom(CStr(Path))

End Sub


Public Function ImportBankWorkbookFrom( _
    ByVal Path As String) _
    As String
'
'   PURPOSE    Open a workbook read-only, read every visible sheet titled
'              C 80.00, C 81.00 or C 84.00 and replace BankCells with its
'              figures.
'   RETURNS    "OK: ..." with the count per template, or "FAIL: ...".
'   STATE OWNERSHIP
'              Opens and closes the bank workbook without saving; restores
'              the application settings.
'   UPDATED    2026-11-05
'
    Dim Book            As Workbook    'Bank workbook
    Dim Sheet           As Worksheet   'One of its sheets
    Dim Values()        As Variant     'Sheet values from A1
    Dim Cells()         As Variant     'Imported rows
    Dim CellCount       As Long
    Dim TemplateCode    As String      'Template of a sheet
    Dim Problem         As String      'Why a sheet could not be read
    Dim Report          As String      'Per-sheet counts
    Dim CellsRead       As Long        'Cells read from a sheet
    Dim Seen            As String      'Templates already imported
    Dim ErrNumber       As Long
    Dim ErrDescription  As String

        On Error GoTo Failed
        If StrComp(Path, ThisWorkbook.FullName, vbTextCompare) = 0 Then
            Err.Raise ERR_HOST, "ImportBankWorkbookFrom", "Choose the bank's workbook, not this one."
        End If
        BeginOperation
        Set Book = Workbooks.Open(Filename:=Path, ReadOnly:=True, UpdateLinks:=0)
        For Each Sheet In Book.Worksheets
            If Sheet.Visible = xlSheetVisible Then
                SheetValues Sheet, Values
                TemplateCode = TemplateOfSheet(Values)
                Select Case TemplateCode
                    Case TEMPLATE_RSF, TEMPLATE_ASF, TEMPLATE_SUMMARY
                        If InStr(Seen, "|" & TemplateCode) > 0 Then
                            Report = Report & vbLf & Sheet.Name & ": " & TemplateCode & " already read, skipped"
                        Else
                            CellsRead = ScanTemplateSheet(Values, TemplateCode, Book.Name & "!" & Sheet.Name, _
                                                     Cells, CellCount, Problem)
                            If Len(Problem) > 0 Then
                                Report = Report & vbLf & Sheet.Name & ": " & Problem
                            Else
                                Report = Report & vbLf & Sheet.Name & ": " & TemplateCode & ", " & CellsRead & " cells"
                                Seen = Seen & "|" & TemplateCode
                            End If
                        End If
                    Case "C 82.00", "C 83.00"
                        Report = Report & vbLf & Sheet.Name & ": " & TemplateCode _
                                 & " (simplified NSFR) not imported"
                End Select
            End If
        Next Sheet
        Book.Close SaveChanges:=False
        Set Book = Nothing
        If CellCount = 0 Then
            Err.Raise ERR_HOST, "ImportBankWorkbookFrom", "No C 80.00, C 81.00 or C 84.00 figures found." & Report
        End If
        ReplaceBankCells Cells, CellCount
        ImportBankWorkbookFrom = "OK: " & CellCount & " bank figures imported from " & Path & Report

CleanUp:
        If Not Book Is Nothing Then
            On Error Resume Next      'closing a read-only copy; the primary error is kept
            Book.Close SaveChanges:=False
            On Error GoTo 0
        End If
        EndOperation
        If ErrNumber <> 0 Then ImportBankWorkbookFrom = "FAIL: " & ErrDescription
        Exit Function
Failed:
        ErrNumber = Err.Number
        ErrDescription = Err.Description
        Resume CleanUp
End Function


Public Function ImportBankCsvFrom( _
    ByVal Path As String) _
    As String
'
'   PURPOSE    Read a CSV export (RowColumn_Index,Suffix,Amount) and
'              replace BankCells with its figures.
'   RETURNS    "OK: ..." or "FAIL: ..."; malformed lines are counted and
'              the first five named.
'   UPDATED    2026-11-05
'
    Dim FileNumber      As Integer     'Open file
    Dim LineText        As String      'One line
    Dim LineNumber      As Long
    Dim Cells()         As Variant     'Imported rows
    Dim CellCount       As Long
    Dim RowItem         As Variant     'BankCells row
    Dim TemplateCode    As String
    Dim RowCode         As String
    Dim ColumnCode      As String
    Dim Value           As Double
    Dim Problem         As String
    Dim Bad             As Long        'Malformed lines
    Dim BadText         As String      'First malformed lines
    Dim ErrNumber       As Long
    Dim ErrDescription  As String

        On Error GoTo Failed
        BeginOperation
        FileNumber = FreeFile
        Open Path For Input As #FileNumber
        Do While Not EOF(FileNumber)
            Line Input #FileNumber, LineText
            LineNumber = LineNumber + 1
            If ParseCsvLine(LineText, TemplateCode, RowCode, ColumnCode, Value, Problem) Then
                RowItem = NewRow(BC_COLUMNS)
                RowItem(BC_TEMPLATE) = TemplateCode
                RowItem(BC_ROW) = RowCode
                RowItem(BC_COLUMN) = ColumnCode
                RowItem(BC_VALUE) = Value
                RowItem(BC_SOURCE) = Dir$(Path) & " line " & LineNumber
                ListAdd Cells, CellCount, RowItem
            ElseIf Len(Problem) > 0 Then
                Bad = Bad + 1
                If Bad <= 5 Then BadText = BadText & vbLf & "line " & LineNumber & ": " & Problem
            End If
        Loop
        Close #FileNumber
        FileNumber = 0
        If CellCount = 0 Then Err.Raise ERR_HOST, "ImportBankCsvFrom", "No figures found in " & Path & BadText
        ReplaceBankCells Cells, CellCount
        ImportBankCsvFrom = "OK: " & CellCount & " bank figures imported from " & Path _
                            & IIf(Bad > 0, vbLf & Bad & " line(s) skipped:" & BadText, "")

CleanUp:
        If FileNumber <> 0 Then Close #FileNumber
        EndOperation
        If ErrNumber <> 0 Then ImportBankCsvFrom = "FAIL: " & ErrDescription
        Exit Function
Failed:
        ErrNumber = Err.Number
        ErrDescription = Err.Description
        Resume CleanUp
End Function


Private Sub ReportImport( _
    ByVal Outcome As String)
'
'   PURPOSE    Show the outcome of an import.
'   UPDATED    2026-11-05
'
        If Left$(Outcome, 3) = "OK:" Then
            MsgBox Mid$(Outcome, 5) & vbLf & vbLf & "Press Run benchmark to reconcile.", vbInformation, APP_TITLE
        Else
            MsgBox Outcome, vbExclamation, APP_TITLE
        End If

End Sub


Private Sub ReplaceBankCells( _
    ByRef Cells() As Variant, _
    ByVal CellCount As Long)
'
'   PURPOSE    Replace the rows of the BankCells table.
'   UPDATED    2026-11-05
'
    Dim Header          As Range       'Header cell of the table

        Set Header = NamedCell("hdrBankCells")
        ClearBelow Header, BC_COLUMNS
        Header.Offset(1, 0).Resize(CellCount, BC_COLUMNS).Value = ListToMatrix(Cells, CellCount, BC_COLUMNS)

End Sub


Private Sub SheetValues( _
    ByVal Sheet As Worksheet, _
    ByRef Values() As Variant)
'
'   PURPOSE    Values of a sheet from A1 to the end of its used range, as a
'              1-based 2-D array (at least 1 x 1).
'   UPDATED    2026-11-05
'
    Dim LastRow         As Long
    Dim LastColumn      As Long
    Dim Block           As Variant     'Range values

        LastRow = Sheet.UsedRange.Row + Sheet.UsedRange.Rows.Count - 1
        LastColumn = Sheet.UsedRange.Column + Sheet.UsedRange.Columns.Count - 1
        If LastRow < 2 Then LastRow = 2
        If LastColumn < 2 Then LastColumn = 2
        Block = Sheet.Range(Sheet.Cells(1, 1), Sheet.Cells(LastRow, LastColumn)).Value
        Values = Block

End Sub

'------------------------------------------------------------------------------
'                                    READING
'------------------------------------------------------------------------------
Private Function NamedCell( _
    ByVal NameText As String) _
    As Range
'
'   PURPOSE    The cell a workbook name refers to.
'   ERROR POLICY
'              Raises ERR_HOST naming the missing name.
'   UPDATED    2026-11-05
'
    Dim Found           As Range       'Referenced cell

        On Error Resume Next          'a missing name is reported below
        Set Found = ThisWorkbook.Names(NameText).RefersToRange
        On Error GoTo 0
        If Found Is Nothing Then
            Err.Raise ERR_HOST, "NamedCell", "Workbook name '" & NameText & "' is missing: rebuild the template."
        End If
        Set NamedCell = Found

End Function


Private Function ReadTable( _
    ByVal HeaderName As String, _
    ByVal Columns As Long, _
    ByRef Data() As Variant, _
    ByVal StopAtBlank As Boolean) _
    As Long
'
'   PURPOSE    Rows of the table below a header cell. StopAtBlank: the
'              table ends at its first entirely blank row (Params, where
'              tables are stacked); otherwise blank rows are dropped and
'              reading goes on to the end of the used range.
'   RETURNS    Row count; Data (1 To Max(count,1), 1 To Columns).
'   UPDATED    2026-11-05
'
    Dim Header          As Range       'Header cell (first column)
    Dim Sheet           As Worksheet   'Table sheet
    Dim LastRow         As Long        'Last used row of the sheet
    Dim Block           As Variant     'Raw values
    Dim Count           As Long        'Rows kept
    Dim LastUsed        As Long        'Last row of Block that belongs to the table
    Dim Blank           As Boolean     'Row entirely blank
    Dim r               As Long
    Dim c               As Long

        Set Header = NamedCell(HeaderName)
        Set Sheet = Header.Worksheet
        LastRow = Sheet.UsedRange.Row + Sheet.UsedRange.Rows.Count - 1
        ReDim Data(1 To 1, 1 To Columns)
        If LastRow <= Header.Row Then Exit Function
        Block = Sheet.Range(Header.Offset(1, 0), Sheet.Cells(LastRow, Header.Column + Columns - 1)).Value
        If Not IsArray(Block) Then
            Exit Function
        End If
        For r = 1 To UBound(Block, 1)
            Blank = True
            For c = 1 To Columns
                If Not IsBlankValue(Block(r, c)) Then
                    Blank = False
                    Exit For
                End If
            Next c
            If Blank And StopAtBlank Then
                LastUsed = r - 1
                Exit For
            End If
            If Not Blank Then Count = Count + 1
            LastUsed = r
        Next r
        If Count = 0 Then Exit Function
        ReDim Data(1 To Count, 1 To Columns)
        Count = 0
        For r = 1 To LastUsed
            Blank = True
            For c = 1 To Columns
                If Not IsBlankValue(Block(r, c)) Then
                    Blank = False
                    Exit For
                End If
            Next c
            If Not Blank Then
                Count = Count + 1
                For c = 1 To Columns
                    Data(Count, c) = Block(r, c)
                Next c
            End If
        Next r
        ReadTable = Count

End Function

'------------------------------------------------------------------------------
'                                    WRITING
'------------------------------------------------------------------------------
Private Sub ClearOutputSheets()
'
'   PURPOSE    Empty the output areas of every output sheet.
'   UPDATED    2026-11-05
'
        ClearBelow NamedCell("hdrPositionCalc"), PO_COLUMNS
        ClearBelow NamedCell("hdrRecon"), RL_COLUMNS
        ClearBelow NamedCell("hdrChecks"), DG_COLUMNS
        NamedCell("resValues").ClearContents
        NamedCell("resSummaryValues").ClearContents
        NamedCell("c80Values").ClearContents
        NamedCell("c81Values").ClearContents
        NamedCell("recHeadline").ClearContents
        NamedCell("recStatus").Value = "No reconciliation yet: enter or import bank figures on BankCells, then run."
        NamedCell("resRunInfo").Value = "No results yet: press Run benchmark."

End Sub


Private Function WithSheet( _
    ByVal SheetName As String) _
    As Worksheet
'
'   PURPOSE    An output sheet of this workbook.
'   UPDATED    2026-11-05
'
        Set WithSheet = ThisWorkbook.Worksheets(SheetName)

End Function


Private Sub ClearBelow( _
    ByVal Header As Range, _
    ByVal Columns As Long)
'
'   PURPOSE    Clear the contents of a table below its header.
'   UPDATED    2026-11-05
'
    Dim Sheet           As Worksheet

        Set Sheet = Header.Worksheet
        Sheet.Range(Header.Offset(1, 0), Header.Offset(MAX_ROWS, Columns - 1)).ClearContents

End Sub


Private Sub WriteBlock( _
    ByVal SheetName As String, _
    ByVal HeaderName As String, _
    ByVal Block As Variant, _
    ByVal Rows As Long, _
    ByVal Columns As Long)
'
'   PURPOSE    Write a 2-D array below a table header.
'   UPDATED    2026-11-05
'
    Dim Header          As Range

        If Rows < 1 Then Exit Sub
        Set Header = NamedCell(HeaderName)
        WithSheet(SheetName).Range(Header.Offset(1, 0), Header.Offset(Rows, Columns - 1)).Value = Block

End Sub


Private Sub WriteResults( _
    ByVal Status As String, _
    ByVal RunClass As String, _
    ByRef Totals() As Variant, _
    ByRef Counts() As Long, _
    ByRef SummaryOut() As Variant, _
    ByVal SummaryRows As Long, _
    ByVal DiagCount As Long, _
    ByRef Diag() As Variant)
'
'   PURPOSE    Headline cells and the C 84.00 values of Results.
'   UPDATED    2026-11-05
'
    Dim Sheet           As Worksheet
    Dim Values(1 To 9, 1 To 1) As Variant 'Headline column
    Dim SummaryBlock()  As Variant     'C 84.00 values
    Dim s               As Long
    Dim c               As Long

        Set Sheet = WithSheet(SH_RESULTS)
        Values(1, 1) = Status
        Values(2, 1) = RunClass
        Values(3, 1) = NamedCell("ReportingDate").Value
        Values(4, 1) = NamedCell("Perimeter").Value
        If Status = STATUS_VALID Then
            Values(5, 1) = Totals(1)
            Values(6, 1) = Totals(2)
            Values(7, 1) = Totals(3)
        Else
            Values(5, 1) = "WITHHELD"
            Values(6, 1) = "WITHHELD"
            Values(7, 1) = "WITHHELD"
        End If
        If UBound(Counts) >= 2 Then
            Values(8, 1) = Counts(1)
            Values(9, 1) = Counts(2)
        End If
        NamedCell("resValues").Value = Values
        If Status = STATUS_VALID And SummaryRows > 0 Then
            ReDim SummaryBlock(1 To SummaryRows, 1 To 4)
            For s = 1 To SummaryRows
                For c = 1 To 4
                    SummaryBlock(s, c) = SummaryOut(s, c)
                Next c
            Next s
            NamedCell("resSummaryValues").Resize(SummaryRows, 4).Value = SummaryBlock
        End If
        NamedCell("resRunInfo").Value = "Run " & Format$(Now, "yyyy-mm-dd hh:nn") & ": " & Status & ", " & RunClass _
                                       & "; " & CountSeverity(Diag, DiagCount, SEV_ERROR) & " error(s), " _
                                       & CountSeverity(Diag, DiagCount, SEV_WARNING) & " warning(s) on Checks."

End Sub


Private Sub WriteTemplates( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef RowAmounts() As Double, _
    ByRef RowStable() As Double, _
    ByVal WithFigures As Boolean)
'
'   PURPOSE    Fill the C80 and C81 replicas: amounts, regulatory and
'              applied factors (effective where the cell holds an amount)
'              and stable funding, matched by row code.
'   INPUTS     WithFigures: False writes the factors only (totals withheld).
'   UPDATED    2026-11-05
'
    Dim Sheet           As Worksheet
    Dim Area            As Range       'Values block of the replica
    Dim Codes           As Variant     'Row codes of the replica
    Dim Block()         As Variant     'Values written
    Dim TemplateCode    As String
    Dim Buckets         As Long        'Amount buckets of the template
    Dim FactorRow       As Long        'Row of the factor table
    Dim Stable          As Double      'Row stable funding
    Dim Present         As Boolean
    Dim Factor          As Double
    Dim t               As Long
    Dim r               As Long
    Dim b               As Long

        For t = 1 To 2
            If t = 1 Then
                Set Sheet = WithSheet(SH_C80)
                Set Area = NamedCell("c80Values")
                TemplateCode = TEMPLATE_RSF
                Buckets = 4
            Else
                Set Sheet = WithSheet(SH_C81)
                Set Area = NamedCell("c81Values")
                TemplateCode = TEMPLATE_ASF
                Buckets = 3
            End If
            Codes = Area.Offset(0, -3).Resize(Area.Rows.Count, 1).Value
            ReDim Block(1 To Area.Rows.Count, 1 To 3 * Buckets + 1)
            For r = 1 To Area.Rows.Count
                FactorRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ROW, TextOf(Codes(r, 1)))
                If FactorRow > 0 Then
                    Stable = 0
                    For b = 1 To Buckets
                        Factor = FactorValue(Factors, FactorRow, FC_REG, b, Present)
                        If Present Then Block(r, Buckets + b) = Factor
                        Factor = FactorValue(Factors, FactorRow, FC_APPLIED, b, Present)
                        If Present Then Block(r, 2 * Buckets + b) = Factor
                        If WithFigures Then
                            If HasBucket(TextOf(Factors(FactorRow, FC_AMOUNT_CELLS)), b) Then
                                Block(r, b) = RowAmounts(FactorRow, b)
                            End If
                            If RowAmounts(FactorRow, b) <> 0 And TextOf(Factors(FactorRow, FC_KIND)) <> KIND_AGG Then
                                Block(r, 2 * Buckets + b) = RowStable(FactorRow, b) / RowAmounts(FactorRow, b)
                            End If
                            Stable = Stable + RowStable(FactorRow, b)
                        End If
                    Next b
                    If WithFigures Then Block(r, 3 * Buckets + 1) = Stable
                End If
            Next r
            Area.Value = Block
        Next t

End Sub


Private Sub WriteReconciliation( _
    ByVal Lines As Variant, _
    ByVal LineCount As Long, _
    ByVal Headline As Variant)
'
'   PURPOSE    Attribution block and lines of Reconciliation.
'   UPDATED    2026-11-05
'
    Dim Sheet           As Worksheet

        Set Sheet = WithSheet(SH_RECON)
        NamedCell("recHeadline").Value = Headline
        NamedCell("recStatus").Value = "Reconciliation " & Format$(Now, "yyyy-mm-dd hh:nn") & ": " & LineCount _
                                      & " line(s); filter Result for DIFF and MISSING."
        If LineCount > 0 Then WriteBlock SH_RECON, "hdrRecon", Lines, LineCount, RL_COLUMNS

End Sub

'------------------------------------------------------------------------------
'                                APPLICATION STATE
'------------------------------------------------------------------------------
Private Sub BeginOperation()
'
'   PURPOSE    Capture and switch off screen updating, events and automatic
'              calculation.
'   UPDATED    2026-11-05
'
        mSavedScreen = Application.ScreenUpdating
        mSavedEvents = Application.EnableEvents
        mSavedCalc = Application.Calculation
        mStateCaptured = True
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual

End Sub


Private Sub EndOperation()
'
'   PURPOSE    Restore the captured settings, each independently, and
'              recalculate.
'   UPDATED    2026-11-05
'
        If Not mStateCaptured Then Exit Sub
        On Error Resume Next          'each restore is attempted even if another fails
        Application.Calculation = mSavedCalc
        Application.EnableEvents = mSavedEvents
        Application.ScreenUpdating = mSavedScreen
        Application.Calculate
        On Error GoTo 0
        mStateCaptured = False

End Sub
