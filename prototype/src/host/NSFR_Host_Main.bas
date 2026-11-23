Attribute VB_Name = "NSFR_Host_Main"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: NSFR_Host_Main
'------------------------------------------------------------------------------
' PURPOSE
'   Workbook glue of the prototype: the macros of the sheet buttons. It
'   reads Params, Positions and BankCells into arrays, calls NSFR_API_Run
'   and writes Results, C80, C81, PositionCalc, Reconciliation and Checks;
'   it imports a bank's figures from its reporting workbook or CSV export
'
' WHY THIS EXISTS
'   NSFR_API_Run and the core work on arrays only; this module is the host
'   adapter that moves values between the workbook and those arrays
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - the button macros: RunNSFR, ImportBankWorkbook, ImportBankCsv,
'       ClearOutputs, ClearBankCells
'     - their silent variants for the regression harness: RunNSFR_Silent,
'       ImportBankWorkbookFrom, ImportBankCsvFrom
'
' ERROR POLICY
'   A narrow On Error handler per public macro captures the primary error,
'   restores the application state, then reports it (MsgBox, or the
'   returned text of a silent variant)
'
' DEPENDENCIES
'   NSFR_API_Functions; NSFR_Core_Config for the layout; NSFR_Core_Util and
'   NSFR_Core_Import for conversion and parsing. Workbook names created by
'   tools/build_template.py locate every table (hdr*) and setting
'
' NOTES
'   Workbook buttons and tools/lo_regression.py call the public procedures
'   by name
'
' PUBLIC SURFACE
'   Buttons: RunNSFR, ImportBankWorkbook, ImportBankCsv, ClearOutputs,
'   ClearBankCells. Silent variants for the regression harness:
'   RunNSFR_Silent, ImportBankWorkbookFrom, ImportBankCsvFrom
'
' STATE OWNERSHIP
'   Changes ScreenUpdating, Calculation and EnableEvents during an
'   operation and restores the captured values on success and failure;
'   leaves the output sheets unprotected (PROTOTYPE.md, Departures)
'
' COMPATIBILITY
'   Excel VBA7, 64-bit, Windows. No references beyond the defaults
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const APP_TITLE             As String = "K-NSFR"                'Title of every message box

    'Error numbers - every host failure
    Private Const ERR_HOST              As Long = vbObjectError + 2101      'A host operation failed

    'Sheet names
    Private Const SH_RESULTS            As String = "Results"               'Headline and C 84.00 values
    Private Const SH_C80                As String = "C80"                   'C 80.00 replica
    Private Const SH_C81                As String = "C81"                   'C 81.00 replica
    Private Const SH_POSCALC            As String = "PositionCalc"          'Per-position calculation
    Private Const SH_RECON              As String = "Reconciliation"        'Reconciliation lines
    Private Const SH_CHECKS             As String = "Checks"                'Diagnostics list
    Private Const SH_BANK               As String = "BankCells"             'Bank figures

    'Output tables
    Private Const MAX_ROWS              As Long = 20000                     'Rows cleared below an output header

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    'Application settings captured by BeginOperation for EndOperation
    Private mSavedScreen                As Boolean                          'Captured ScreenUpdating
    Private mSavedEvents                As Boolean                          'Captured EnableEvents
    Private mSavedCalc                  As Long                             'Captured Calculation
    Private mStateCaptured              As Boolean                          'Settings captured and not yet restored


'
'------------------------------------------------------------------------------
'
'                                   BUTTONS
'
'------------------------------------------------------------------------------
'

Public Sub RunNSFR()

'
'------------------------------------------------------------------------------
'                       RUN THE BENCHMARK AND REPORT IT
'------------------------------------------------------------------------------
' PURPOSE
'   Button "Run benchmark": run and report the outcome
'
' WHY THIS EXISTS
'   The button needs a message box; the run itself is RunNSFR_Silent, which
'   the regression harness calls without one
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunNSFR_Silent, then shows its summary as information when it
'   starts with "OK:", or the whole text as a warning otherwise
'
' ERROR POLICY
'   No handler: RunNSFR_Silent returns host errors as text, which this
'   shows
'
' DEPENDENCIES
'   RunNSFR_Silent
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.RunNSFR"

    Dim Outcome                 As String       'Result text

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Run the benchmark without messages
        Outcome = RunNSFR_Silent()

'------------------------------------------------------------------------------
' REPORT THE OUTCOME
'------------------------------------------------------------------------------
    'Show the summary of a completed run, or the failure
        If Left$(Outcome, 3) = "OK:" Then
            MsgBox Mid$(Outcome, 5), vbInformation, APP_TITLE
        Else
            MsgBox Outcome, vbExclamation, APP_TITLE
        End If

End Sub

Public Function RunNSFR_Silent() As String

'
'------------------------------------------------------------------------------
'                   RUN THE BENCHMARK AND WRITE THE OUTPUTS
'------------------------------------------------------------------------------
' PURPOSE
'   Run the benchmark and the reconciliation, write every output sheet and
'   return the outcome as text
'
' WHY THIS EXISTS
'   The button and the regression harness share one run; this variant
'   shows no message box
'
' INPUTS
'   None; reads the Params tables, Positions, BankCells and the named
'   settings of the workbook
'
' RETURNS
'   "OK: <summary>" when the run completed (its status may still be INVALID
'   or STOPPED, as Results shows); "FAIL: <error>" when a host error
'   stopped it
'
' BEHAVIOR
'   Reads the input tables and the bank settings, calls NSFR_API_Run,
'   clears the output sheets and writes Results. Unless the status starts
'   with STOPPED, also writes C80 and C81 (figures only when VALID),
'   PositionCalc when there are positions and Reconciliation when the bank
'   figures are usable. Checks is written whenever there are diagnostics
'
' ERROR POLICY
'   Contained: a host error is captured, the application settings are
'   restored, and the error is returned as "FAIL: <description> (<number>)"
'
' DEPENDENCIES
'   BeginOperation, EndOperation, ReadTable, NamedCell, NSFR_API_Run,
'   ClearOutputSheets, WriteResults, WriteTemplates, WriteBlock,
'   WriteReconciliation, ListToMatrix, CountSeverity, TextOf
'
' NOTES
'   Called by tools/lo_regression.py as NSFR_Host_Main.RunNSFR_Silent
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.RunNSFR_Silent"

    Dim Factors()               As Variant      'Params factor table
    Dim Categories()            As Variant      'Params category table
    Dim Summary()               As Variant      'Params summary map
    Dim Positions()             As Variant      'Positions
    Dim BankCells()             As Variant      'Bank figures
    Dim BankSettings(1 To 4)    As Variant      'AmountScale and tolerances
    Dim FactorRows              As Long         'Rows of Factors
    Dim CategoryRows            As Long         'Rows of Categories
    Dim SummaryRows             As Long         'Rows of Summary
    Dim PositionRows            As Long         'Rows of Positions
    Dim BankRows                As Long         'Rows of BankCells
    Dim RowAmounts()            As Double       'Benchmark amount per row and bucket
    Dim RowStable()             As Double       'Benchmark stable funding per row and bucket
    Dim PositionOut()           As Variant      'PositionCalc rows
    Dim SummaryOut()            As Variant      'C 84.00 values
    Dim Totals()                As Variant      'ASF, RSF, NSFR
    Dim Counts()                As Long         'Accepted, rejected, overrides
    Dim RunClass                As String       'REFERENCE or WHAT-IF
    Dim Lines                   As Variant      'Reconciliation lines
    Dim LineCount               As Long         'Number of reconciliation lines
    Dim Headline                As Variant      'Attribution
    Dim HasBank                 As Boolean      'Bank figures usable
    Dim Diag()                  As Variant      'Diagnostics list
    Dim DiagCount               As Long         'Number of diagnostics
    Dim Status                  As String       'Run status
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain every host error in the returned text
        On Error GoTo ErrorHandler
    'Switch off screen updating, events and automatic calculation
        BeginOperation

'------------------------------------------------------------------------------
' READ THE INPUTS
'------------------------------------------------------------------------------
    'Read the Params tables, which end at their first blank row
        FactorRows = ReadTable("hdrFactors", FC_COLUMNS, Factors, True)
        CategoryRows = ReadTable("hdrCategories", CC_COLUMNS, Categories, True)
        SummaryRows = ReadTable("hdrSummary", SC_COLUMNS, Summary, True)
    'Read the positions and the bank figures
        PositionRows = ReadTable("hdrPositions", PC_COLUMNS, Positions, False)
        BankRows = ReadTable("hdrBankCells", BC_COLUMNS, BankCells, False)
    'Read the amount scale and the tolerances
        BankSettings(1) = NamedCell("AmountScale").Value
        BankSettings(2) = NamedCell("AbsTolerance").Value
        BankSettings(3) = NamedCell("RelTolerance").Value
        BankSettings(4) = NamedCell("FactorTolerance").Value

'------------------------------------------------------------------------------
' RUN THE ENGINE
'------------------------------------------------------------------------------
    'Validate, calculate and reconcile
        Status = NSFR_API_Run(TextOf(NamedCell("ProfileID").Value), NamedCell("ReportingDate").Value, _
                              Factors, FactorRows, Categories, CategoryRows, Summary, SummaryRows, _
                              Positions, PositionRows, BankCells, BankRows, BankSettings, _
                              RowAmounts, RowStable, PositionOut, SummaryOut, Totals, RunClass, Counts, _
                              Lines, LineCount, Headline, HasBank, Diag, DiagCount)

'------------------------------------------------------------------------------
' WRITE THE OUTPUTS
'------------------------------------------------------------------------------
    'Replace the previous outputs with the headline of this run
        ClearOutputSheets
        WriteResults Status, RunClass, Totals, Counts, SummaryOut, SummaryRows, DiagCount, Diag
    'Write the replicas, the positions and the reconciliation unless the run stopped
        If Left$(Status, Len(STATUS_STOPPED)) <> STATUS_STOPPED Then
            WriteTemplates Factors, FactorRows, RowAmounts, RowStable, (Status = STATUS_VALID)
            If PositionRows > 0 Then WriteBlock SH_POSCALC, "hdrPositionCalc", PositionOut, PositionRows, _
                                                PO_COLUMNS
            If HasBank Then WriteReconciliation Lines, LineCount, Headline
        End If
    'Write the diagnostics
        If DiagCount > 0 Then WriteBlock SH_CHECKS, "hdrChecks", ListToMatrix(Diag, DiagCount, DG_COLUMNS), _
                                         DiagCount, DG_COLUMNS

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Summarize the run
        RunNSFR_Silent = "OK: " & Status & " | " & RunClass & " | errors " _
                         & CountSeverity(Diag, DiagCount, SEV_ERROR) _
                         & ", warnings " & CountSeverity(Diag, DiagCount, SEV_WARNING) _
                         & IIf(HasBank, " | reconciliation lines " & LineCount, " | no bank figures")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Restore the application settings
        EndOperation
    'Return a captured error instead of the summary
        If ErrorNumber <> 0 Then RunNSFR_Silent = "FAIL: " & ErrorDescription & " (" & ErrorNumber & ")"
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Restore the settings and return the error on the common exit path
        Resume CleanUp

End Function

Public Sub ClearOutputs()

'
'------------------------------------------------------------------------------
'                           CLEAR THE OUTPUT SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Button "Clear outputs": empty every output sheet
'
' WHY THIS EXISTS
'   Lets the user remove the results of a previous run before the next one
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Switches off screen updating, events and automatic calculation, clears
'   the output areas, then restores the settings
'
' ERROR POLICY
'   Contained: a host error is captured, the settings are restored and the
'   error is shown in a message box
'
' DEPENDENCIES
'   BeginOperation, ClearOutputSheets, EndOperation
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.ClearOutputs"

    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain every host error
        On Error GoTo ErrorHandler
    'Switch off screen updating, events and automatic calculation
        BeginOperation

'------------------------------------------------------------------------------
' CLEAR THE OUTPUTS
'------------------------------------------------------------------------------
    'Empty every output area
        ClearOutputSheets

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Restore the application settings
        EndOperation
    'Show a captured error
        If ErrorNumber <> 0 Then MsgBox "Clear stopped: " & ErrorDescription, vbExclamation, APP_TITLE
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Restore the settings and show the error on the common exit path
        Resume CleanUp

End Sub

Public Sub ClearBankCells()

'
'------------------------------------------------------------------------------
'                            CLEAR THE BANK FIGURES
'------------------------------------------------------------------------------
' PURPOSE
'   Button "Clear bank figures": empty the BankCells table
'
' WHY THIS EXISTS
'   Lets the user remove a bank's figures before entering or importing
'   others
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Asks for confirmation, then clears the rows below the BankCells header
'
' ERROR POLICY
'   No handler: errors, such as a missing workbook name, propagate to Excel
'
' DEPENDENCIES
'   NamedCell, ClearBelow
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CONFIRM
'------------------------------------------------------------------------------
    'Stop unless the user confirms
        If MsgBox("Clear every bank figure on BankCells?", vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub

'------------------------------------------------------------------------------
' CLEAR THE TABLE
'------------------------------------------------------------------------------
    'Clear the rows below the header
        ClearBelow NamedCell("hdrBankCells"), BC_COLUMNS

End Sub

'
'------------------------------------------------------------------------------
'
'                                    IMPORT
'
'------------------------------------------------------------------------------
'

Public Sub ImportBankWorkbook()

'
'------------------------------------------------------------------------------
'                       IMPORT A BANK REPORTING WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Button "Import bank workbook": ask for the bank's reporting workbook and
'   import its C 80.00, C 81.00 and C 84.00 sheets
'
' WHY THIS EXISTS
'   The button needs a file dialog and a message box; the import itself is
'   ImportBankWorkbookFrom, which the regression harness calls without them
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Shows the file dialog; does nothing when it is cancelled; otherwise
'   imports the chosen workbook and shows the outcome
'
' ERROR POLICY
'   No handler: ImportBankWorkbookFrom returns its errors as text, which
'   ReportImport shows
'
' DEPENDENCIES
'   Application.GetOpenFilename, ImportBankWorkbookFrom, ReportImport
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.ImportBankWorkbook"

    Dim Path                    As Variant      'Chosen file, or False

'------------------------------------------------------------------------------
' CHOOSE THE FILE
'------------------------------------------------------------------------------
    'Ask for the workbook; stop when the dialog is cancelled
        Path = Application.GetOpenFilename("Excel workbooks (*.xls*),*.xls*", , "Bank NSFR reporting workbook")
        If VarType(Path) = vbBoolean Then Exit Sub

'------------------------------------------------------------------------------
' IMPORT
'------------------------------------------------------------------------------
    'Import the workbook and show the outcome
        ReportImport ImportBankWorkbookFrom(CStr(Path))

End Sub

Public Sub ImportBankCsv()

'
'------------------------------------------------------------------------------
'                           IMPORT A BANK CSV EXPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Button "Import bank CSV": ask for a CSV export
'   (RowColumn_Index,Suffix,Amount) and import it
'
' WHY THIS EXISTS
'   The button needs a file dialog and a message box; the import itself is
'   ImportBankCsvFrom, which the regression harness calls without them
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Shows the file dialog; does nothing when it is cancelled; otherwise
'   imports the chosen file and shows the outcome
'
' ERROR POLICY
'   No handler: ImportBankCsvFrom returns its errors as text, which
'   ReportImport shows
'
' DEPENDENCIES
'   Application.GetOpenFilename, ImportBankCsvFrom, ReportImport
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.ImportBankCsv"

    Dim Path                    As Variant      'Chosen file, or False

'------------------------------------------------------------------------------
' CHOOSE THE FILE
'------------------------------------------------------------------------------
    'Ask for the CSV file; stop when the dialog is cancelled
        Path = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Bank NSFR CSV export")
        If VarType(Path) = vbBoolean Then Exit Sub

'------------------------------------------------------------------------------
' IMPORT
'------------------------------------------------------------------------------
    'Import the file and show the outcome
        ReportImport ImportBankCsvFrom(CStr(Path))

End Sub

Public Function ImportBankWorkbookFrom( _
    ByVal Path As String) As String
'
'------------------------------------------------------------------------------
'                   IMPORT THE TEMPLATE SHEETS OF A WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Open a workbook read-only, read every visible sheet titled C 80.00,
'   C 81.00 or C 84.00 and replace BankCells with its figures
'
' WHY THIS EXISTS
'   A bank's figures usually come from its own reporting workbook; this
'   variant shows no dialog or message box, so the regression harness can
'   call it
'
' INPUTS
'   Path
'     Full path of the bank's reporting workbook; not this workbook
'
' RETURNS
'   "OK: <count> bank figures imported from <path>" followed by one line per
'   sheet read or skipped, or "FAIL: <description>"
'
' BEHAVIOR
'   Rejects this workbook, opens the bank workbook read-only, and scans each
'   visible sheet whose template is C 80.00, C 81.00 or C 84.00; a second
'   sheet of the same template is skipped, and C 82.00 / C 83.00 sheets
'   (simplified NSFR) are reported as not imported. Closes the workbook
'   without saving, then replaces the BankCells rows with the figures read
'
' ERROR POLICY
'   Contained: raises ERR_HOST internally for this workbook or when no
'   figure is found; any error closes the bank workbook, restores the
'   application settings and is returned as "FAIL: <description>"
'
' DEPENDENCIES
'   BeginOperation, EndOperation, SheetValues, TemplateOfSheet,
'   ScanTemplateSheet, ReplaceBankCells
'
' NOTES
'   Called by tools/lo_regression.py as
'   NSFR_Host_Main.ImportBankWorkbookFrom
'
' STATE OWNERSHIP
'   Opens and closes the bank workbook without saving; restores the
'   application settings
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.ImportBankWorkbookFrom"

    Dim Book                    As Workbook     'Bank workbook
    Dim Sheet                   As Worksheet    'One of its sheets
    Dim Values()                As Variant      'Sheet values from A1
    Dim Cells()                 As Variant      'Imported rows
    Dim CellCount               As Long         'Number of imported rows
    Dim TemplateCode            As String       'Template of a sheet
    Dim Problem                 As String       'Why a sheet could not be read
    Dim Report                  As String       'Per-sheet counts
    Dim CellsRead               As Long         'Cells read from a sheet
    Dim Seen                    As String       'Templates already imported
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain every error in the returned text
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject this workbook
        If StrComp(Path, ThisWorkbook.FullName, vbTextCompare) = 0 Then
            Err.Raise ERR_HOST, "ImportBankWorkbookFrom", _
                "Choose the bank's workbook, not this one."
        End If

'------------------------------------------------------------------------------
' OPEN THE WORKBOOK
'------------------------------------------------------------------------------
    'Switch off screen updating, events and automatic calculation
        BeginOperation
    'Open the bank workbook read-only, without updating its links
        Set Book = Workbooks.Open(Filename:=Path, ReadOnly:=True, UpdateLinks:=0)

'------------------------------------------------------------------------------
' READ THE TEMPLATE SHEETS
'------------------------------------------------------------------------------
    'Scan each visible sheet of a supported template, once per template
        For Each Sheet In Book.Worksheets
            If Sheet.Visible = xlSheetVisible Then
                SheetValues Sheet, Values
                TemplateCode = TemplateOfSheet(Values)
                Select Case TemplateCode
                    Case TEMPLATE_RSF, TEMPLATE_ASF, TEMPLATE_SUMMARY
                        If InStr(1, Seen, "|" & TemplateCode, vbBinaryCompare) > 0 Then
                            Report = Report & vbLf & Sheet.Name & ": " & TemplateCode _
                                     & " already read, skipped"
                        Else
                            CellsRead = ScanTemplateSheet(Values, TemplateCode, Book.Name & "!" & Sheet.Name, _
                                                          Cells, CellCount, Problem)
                            If Len(Problem) > 0 Then
                                Report = Report & vbLf & Sheet.Name & ": " & Problem
                            Else
                                Report = Report & vbLf & Sheet.Name & ": " & TemplateCode _
                                         & ", " & CellsRead & " cells"
                                Seen = Seen & "|" & TemplateCode
                            End If
                        End If
                    Case "C 82.00", "C 83.00"
                        Report = Report & vbLf & Sheet.Name & ": " & TemplateCode _
                                 & " (simplified NSFR) not imported"
                    Case Else
                        'No action for other sheets
                End Select
            End If
        Next Sheet

'------------------------------------------------------------------------------
' CLOSE THE WORKBOOK
'------------------------------------------------------------------------------
    'Close the bank workbook without saving
        Book.Close SaveChanges:=False
        Set Book = Nothing

'------------------------------------------------------------------------------
' STORE THE FIGURES
'------------------------------------------------------------------------------
    'Fail when no figure was found
        If CellCount = 0 Then
            Err.Raise ERR_HOST, "ImportBankWorkbookFrom", _
                "No C 80.00, C 81.00 or C 84.00 figures found." & Report
        End If
    'Replace the BankCells rows
        ReplaceBankCells Cells, CellCount

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Report the count and the per-sheet lines
        ImportBankWorkbookFrom = "OK: " & CellCount & " bank figures imported from " & Path & Report

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Close a bank workbook still open after an error; the primary error is kept
        If Not Book Is Nothing Then
            On Error Resume Next
            Book.Close SaveChanges:=False
            On Error GoTo 0
        End If
    'Restore the application settings
        EndOperation
    'Return a captured error instead of the summary
        If ErrorNumber <> 0 Then ImportBankWorkbookFrom = "FAIL: " & ErrorDescription
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Clean up and return the error on the common exit path
        Resume CleanUp

End Function

Public Function ImportBankCsvFrom( _
    ByVal Path As String) As String
'
'------------------------------------------------------------------------------
'                             IMPORT A CSV EXPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Read a CSV export (RowColumn_Index,Suffix,Amount) and replace BankCells
'   with its figures
'
' WHY THIS EXISTS
'   A bank may provide its figures as a CSV export instead of a workbook;
'   this variant shows no dialog or message box, so the regression harness
'   can call it
'
' INPUTS
'   Path
'     Full path of the CSV file
'
' RETURNS
'   "OK: ..." or "FAIL: ..."; malformed lines are counted and the first
'   five named
'
' BEHAVIOR
'   Reads the file line by line; each line ParseCsvLine accepts becomes a
'   BankCells row whose source is the file name and line number; a line
'   rejected with a problem is counted as skipped. Then replaces the
'   BankCells rows with the rows read
'
' ERROR POLICY
'   Contained: raises ERR_HOST internally when no figure is found; any
'   error closes the file, restores the application settings and is
'   returned as "FAIL: <description>"
'
' DEPENDENCIES
'   BeginOperation, EndOperation, ParseCsvLine, NewRow, ListAdd,
'   ReplaceBankCells; VBA file statements
'
' NOTES
'   Called by tools/lo_regression.py as NSFR_Host_Main.ImportBankCsvFrom
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.ImportBankCsvFrom"

    Dim FileNumber              As Integer      'Open file
    Dim LineText                As String       'One line
    Dim LineNumber              As Long         'Number of the line read
    Dim Cells()                 As Variant      'Imported rows
    Dim CellCount               As Long         'Number of imported rows
    Dim RowItem                 As Variant      'BankCells row
    Dim TemplateCode            As String       'Template of the line
    Dim RowCode                 As String       'Row code of the line
    Dim ColumnCode              As String       'Column code of the line
    Dim Value                   As Double       'Amount of the line
    Dim Problem                 As String       'Why the line was rejected
    Dim Bad                     As Long         'Malformed lines
    Dim BadText                 As String       'First malformed lines
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain every error in the returned text
        On Error GoTo ErrorHandler
    'Switch off screen updating, events and automatic calculation
        BeginOperation

'------------------------------------------------------------------------------
' READ THE FILE
'------------------------------------------------------------------------------
    'Open the file
        FileNumber = FreeFile
        Open Path For Input As #FileNumber
    'Keep each accepted line as a BankCells row; count and name the malformed ones
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
    'Close the file
        Close #FileNumber
        FileNumber = 0

'------------------------------------------------------------------------------
' STORE THE FIGURES
'------------------------------------------------------------------------------
    'Fail when no figure was found
        If CellCount = 0 Then Err.Raise ERR_HOST, "ImportBankCsvFrom", "No figures found in " & Path & BadText
    'Replace the BankCells rows
        ReplaceBankCells Cells, CellCount

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Report the count and the skipped lines
        ImportBankCsvFrom = "OK: " & CellCount & " bank figures imported from " & Path _
                            & IIf(Bad > 0, vbLf & Bad & " line(s) skipped:" & BadText, "")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Close a file still open after an error
        If FileNumber <> 0 Then Close #FileNumber
    'Restore the application settings
        EndOperation
    'Return a captured error instead of the summary
        If ErrorNumber <> 0 Then ImportBankCsvFrom = "FAIL: " & ErrorDescription
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Clean up and return the error on the common exit path
        Resume CleanUp

End Function

Private Sub ReportImport( _
    ByVal Outcome As String)
'
'------------------------------------------------------------------------------
'                        SHOW THE OUTCOME OF AN IMPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Show the outcome of an import
'
' INPUTS
'   Outcome
'     The text an import returned: "OK: ..." or "FAIL: ..."
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Show the summary with the next step, or the failure
        If Left$(Outcome, 3) = "OK:" Then
            MsgBox Mid$(Outcome, 5) & vbLf & vbLf & "Press Run benchmark to reconcile.", _
                   vbInformation, APP_TITLE
        Else
            MsgBox Outcome, vbExclamation, APP_TITLE
        End If

End Sub

Private Sub ReplaceBankCells( _
    ByRef Cells() As Variant, _
    ByVal CellCount As Long)
'
'------------------------------------------------------------------------------
'                           REPLACE THE BANK FIGURES
'------------------------------------------------------------------------------
' PURPOSE
'   Replace the rows of the BankCells table
'
' INPUTS
'   Cells
'     The imported rows, as a list
'
'   CellCount
'     Number of rows in Cells
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.ReplaceBankCells"

    Dim Header                  As Range        'Header cell of the table

'------------------------------------------------------------------------------
' REPLACE THE ROWS
'------------------------------------------------------------------------------
    'Clear the table, then write the rows below its header
        Set Header = NamedCell("hdrBankCells")
        ClearBelow Header, BC_COLUMNS
        Header.Offset(1, 0).Resize(CellCount, BC_COLUMNS).Value = ListToMatrix(Cells, CellCount, BC_COLUMNS)

End Sub

Private Sub SheetValues( _
    ByVal Sheet As Worksheet, _
    ByRef Values() As Variant)
'
'------------------------------------------------------------------------------
'                          READ THE VALUES OF A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Values of a sheet from A1 to the end of its used range, as a 1-based
'   2-D array (at least 1 x 1)
'
' INPUTS
'   Sheet
'     The sheet to read
'
'   Values
'     Output: its values
'
' NOTES
'   The range read is at least A1:B2, so Range.Value always returns an
'   array
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.SheetValues"

    Dim LastRow                 As Long         'Last row read
    Dim LastColumn              As Long         'Last column read
    Dim Block                   As Variant      'Range values

'------------------------------------------------------------------------------
' LOCATE THE RANGE
'------------------------------------------------------------------------------
    'Find the end of the used range, at least row 2 and column 2
        LastRow = Sheet.UsedRange.Row + Sheet.UsedRange.Rows.Count - 1
        LastColumn = Sheet.UsedRange.Column + Sheet.UsedRange.Columns.Count - 1
        If LastRow < 2 Then LastRow = 2
        If LastColumn < 2 Then LastColumn = 2

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Read the values from A1
        Block = Sheet.Range(Sheet.Cells(1, 1), Sheet.Cells(LastRow, LastColumn)).Value
        Values = Block

End Sub

'
'------------------------------------------------------------------------------
'
'                                   READING
'
'------------------------------------------------------------------------------
'

Private Function NamedCell( _
    ByVal NameText As String) As Range
'
'------------------------------------------------------------------------------
'                       FIND THE CELL OF A WORKBOOK NAME
'------------------------------------------------------------------------------
' PURPOSE
'   The cell a workbook name refers to
'
' INPUTS
'   NameText
'     The workbook name
'
' ERROR POLICY
'   Raises ERR_HOST naming the missing name
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.NamedCell"

    Dim Found                   As Range        'Referenced cell

'------------------------------------------------------------------------------
' FIND THE CELL
'------------------------------------------------------------------------------
    'Look the name up; a missing name is reported below
        On Error Resume Next
        Set Found = ThisWorkbook.Names(NameText).RefersToRange
        On Error GoTo 0
    'Fail on a missing name
        If Found Is Nothing Then
            Err.Raise ERR_HOST, "NamedCell", _
                "Workbook name '" & NameText & "' is missing: rebuild the template."
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the cell
        Set NamedCell = Found

End Function

Private Function ReadTable( _
    ByVal HeaderName As String, _
    ByVal Columns As Long, _
    ByRef Data() As Variant, _
    ByVal StopAtBlank As Boolean) As Long
'
'------------------------------------------------------------------------------
'                           READ THE ROWS OF A TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Rows of the table below a header cell
'
' WHY THIS EXISTS
'   The input tables have no fixed length; Params stacks several tables on
'   one sheet, separated by blank rows
'
' INPUTS
'   HeaderName
'     Workbook name of the header cell (first column)
'
'   Columns
'     Number of columns of the table
'
'   Data
'     Output: the rows kept
'
'   StopAtBlank
'     True: the table ends at its first entirely blank row (Params, where
'     tables are stacked). False: blank rows are dropped and reading goes on
'     to the end of the used range
'
' RETURNS
'   Row count; Data (1 To Max(count,1), 1 To Columns)
'
' BEHAVIOR
'   Reads the block from the row below the header to the last used row of
'   the sheet, finds the rows that belong to the table, then copies the
'   rows that are not entirely blank
'
' ERROR POLICY
'   No handler: NamedCell raises ERR_HOST for a missing name; other errors
'   propagate
'
' DEPENDENCIES
'   NamedCell, IsBlankValue
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.ReadTable"

    Dim Header                  As Range        'Header cell (first column)
    Dim Sheet                   As Worksheet    'Table sheet
    Dim LastRow                 As Long         'Last used row of the sheet
    Dim Block                   As Variant      'Raw values
    Dim Count                   As Long         'Rows kept
    Dim LastUsed                As Long         'Last row of Block that belongs to the table
    Dim Blank                   As Boolean      'Row entirely blank
    Dim RowIndex                As Long         'Row of Block
    Dim ColumnIndex             As Long         'Column of Block

'------------------------------------------------------------------------------
' LOCATE THE TABLE
'------------------------------------------------------------------------------
    'Find the header and the last used row of its sheet
        Set Header = NamedCell(HeaderName)
        Set Sheet = Header.Worksheet
        LastRow = Sheet.UsedRange.Row + Sheet.UsedRange.Rows.Count - 1
    'Return an empty table when nothing is below the header
        ReDim Data(1 To 1, 1 To Columns)
        If LastRow <= Header.Row Then Exit Function

'------------------------------------------------------------------------------
' READ THE BLOCK
'------------------------------------------------------------------------------
    'Read the values below the header
        Block = Sheet.Range(Header.Offset(1, 0), Sheet.Cells(LastRow, Header.Column + Columns - 1)).Value
        If Not IsArray(Block) Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' COUNT THE ROWS
'------------------------------------------------------------------------------
    'Count the non-blank rows up to the end of the table
        For RowIndex = 1 To UBound(Block, 1)
            Blank = True
            For ColumnIndex = 1 To Columns
                If Not IsBlankValue(Block(RowIndex, ColumnIndex)) Then
                    Blank = False
                    Exit For
                End If
            Next ColumnIndex
            If Blank And StopAtBlank Then
                LastUsed = RowIndex - 1
                Exit For
            End If
            If Not Blank Then Count = Count + 1
            LastUsed = RowIndex
        Next RowIndex
        If Count = 0 Then Exit Function

'------------------------------------------------------------------------------
' COPY THE ROWS
'------------------------------------------------------------------------------
    'Copy the non-blank rows of the table
        ReDim Data(1 To Count, 1 To Columns)
        Count = 0
        For RowIndex = 1 To LastUsed
            Blank = True
            For ColumnIndex = 1 To Columns
                If Not IsBlankValue(Block(RowIndex, ColumnIndex)) Then
                    Blank = False
                    Exit For
                End If
            Next ColumnIndex
            If Not Blank Then
                Count = Count + 1
                For ColumnIndex = 1 To Columns
                    Data(Count, ColumnIndex) = Block(RowIndex, ColumnIndex)
                Next ColumnIndex
            End If
        Next RowIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the row count
        ReadTable = Count

End Function

'
'------------------------------------------------------------------------------
'
'                                   WRITING
'
'------------------------------------------------------------------------------
'

Private Sub ClearOutputSheets()

'
'------------------------------------------------------------------------------
'                            EMPTY THE OUTPUT AREAS
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the output areas of every output sheet
'
' BEHAVIOR
'   Clears the PositionCalc, Reconciliation and Checks tables and the value
'   areas of Results, C80, C81 and Reconciliation, then writes the "no
'   results yet" texts
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR THE TABLES
'------------------------------------------------------------------------------
    'Clear the rows below the output headers
        ClearBelow NamedCell("hdrPositionCalc"), PO_COLUMNS
        ClearBelow NamedCell("hdrRecon"), RL_COLUMNS
        ClearBelow NamedCell("hdrChecks"), DG_COLUMNS

'------------------------------------------------------------------------------
' CLEAR THE VALUE AREAS
'------------------------------------------------------------------------------
    'Clear the values of Results, the replicas and the attribution
        NamedCell("resValues").ClearContents
        NamedCell("resSummaryValues").ClearContents
        NamedCell("c80Values").ClearContents
        NamedCell("c81Values").ClearContents
        NamedCell("recHeadline").ClearContents
    'Say that there are no results yet
        NamedCell("recStatus").Value = "No reconciliation yet: enter or import bank figures on BankCells, then run."
        NamedCell("resRunInfo").Value = "No results yet: press Run benchmark."

End Sub

Private Function WithSheet( _
    ByVal SheetName As String) As Worksheet
'
'------------------------------------------------------------------------------
'                             GET AN OUTPUT SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   An output sheet of this workbook
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the sheet of that name
        Set WithSheet = ThisWorkbook.Worksheets(SheetName)

End Function

Private Sub ClearBelow( _
    ByVal Header As Range, _
    ByVal Columns As Long)
'
'------------------------------------------------------------------------------
'                        CLEAR A TABLE BELOW ITS HEADER
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the contents of a table below its header
'
' NOTES
'   Clears MAX_ROWS rows
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.ClearBelow"

    Dim Sheet                   As Worksheet    'Sheet of the table

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'Clear MAX_ROWS rows below the header
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
'------------------------------------------------------------------------------
'                         WRITE A BLOCK BELOW A HEADER
'------------------------------------------------------------------------------
' PURPOSE
'   Write a 2-D array below a table header
'
' INPUTS
'   Rows, Columns
'     Size of Block; nothing is written when Rows is below 1
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.WriteBlock"

    Dim Header                  As Range        'Header cell of the table

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Skip an empty block
        If Rows < 1 Then Exit Sub
    'Write the block below the header
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
'------------------------------------------------------------------------------
'                           WRITE THE RESULTS SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Headline cells and the C 84.00 values of Results
'
' WHY THIS EXISTS
'   Results is the first sheet the user reads after a run
'
' INPUTS
'   Status, RunClass
'     Run status and class from NSFR_API_Run
'
'   Totals
'     ASF, RSF and NSFR
'
'   Counts
'     Accepted and rejected positions
'
'   SummaryOut, SummaryRows
'     C 84.00 values and their row count
'
'   DiagCount, Diag
'     The diagnostics, counted by severity in the run information
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the nine headline values; the totals only when the status is
'   VALID, WITHHELD otherwise. Writes the C 84.00 values only when VALID,
'   then the run information line with the time and the diagnostic counts
'
' ERROR POLICY
'   No handler: errors propagate to the caller
'
' DEPENDENCIES
'   WithSheet, NamedCell, CountSeverity
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.WriteResults"

    Dim Sheet                   As Worksheet    'Results sheet
    Dim Values(1 To 9, 1 To 1)  As Variant      'Headline column
    Dim SummaryBlock()          As Variant      'C 84.00 values
    Dim RowIndex                As Long         'Row of SummaryOut
    Dim ColumnIndex             As Long         'Column of SummaryOut

'------------------------------------------------------------------------------
' WRITE THE HEADLINE
'------------------------------------------------------------------------------
    'Status, class, date and perimeter
        Set Sheet = WithSheet(SH_RESULTS)
        Values(1, 1) = Status
        Values(2, 1) = RunClass
        Values(3, 1) = NamedCell("ReportingDate").Value
        Values(4, 1) = NamedCell("Perimeter").Value
    'Totals only for a valid run
        If Status = STATUS_VALID Then
            Values(5, 1) = Totals(1)
            Values(6, 1) = Totals(2)
            Values(7, 1) = Totals(3)
        Else
            Values(5, 1) = "WITHHELD"
            Values(6, 1) = "WITHHELD"
            Values(7, 1) = "WITHHELD"
        End If
    'Position counts
        If UBound(Counts) >= 2 Then
            Values(8, 1) = Counts(1)
            Values(9, 1) = Counts(2)
        End If
        NamedCell("resValues").Value = Values

'------------------------------------------------------------------------------
' WRITE THE C 84.00 VALUES
'------------------------------------------------------------------------------
    'Copy the four value columns of a valid run
        If Status = STATUS_VALID And SummaryRows > 0 Then
            ReDim SummaryBlock(1 To SummaryRows, 1 To 4)
            For RowIndex = 1 To SummaryRows
                For ColumnIndex = 1 To 4
                    SummaryBlock(RowIndex, ColumnIndex) = SummaryOut(RowIndex, ColumnIndex)
                Next ColumnIndex
            Next RowIndex
            NamedCell("resSummaryValues").Resize(SummaryRows, 4).Value = SummaryBlock
        End If

'------------------------------------------------------------------------------
' WRITE THE RUN INFORMATION
'------------------------------------------------------------------------------
    'Time, status, class and diagnostic counts
        NamedCell("resRunInfo").Value = "Run " & Format$(Now, "yyyy-mm-dd hh:nn") & ": " & Status _
                                        & ", " & RunClass & "; " _
                                        & CountSeverity(Diag, DiagCount, SEV_ERROR) & " error(s), " _
                                        & CountSeverity(Diag, DiagCount, SEV_WARNING) & " warning(s) on Checks."

End Sub

Private Sub WriteTemplates( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef RowAmounts() As Double, _
    ByRef RowStable() As Double, _
    ByVal WithFigures As Boolean)
'
'------------------------------------------------------------------------------
'                        FILL THE C80 AND C81 REPLICAS
'------------------------------------------------------------------------------
' PURPOSE
'   Fill the C80 and C81 replicas: amounts, regulatory and applied factors
'   (effective where the cell holds an amount) and stable funding, matched
'   by row code
'
' WHY THIS EXISTS
'   The replicas show the benchmark in the layout of the reporting
'   templates
'
' INPUTS
'   Factors, FactorRows
'     The factor table and its row count
'
'   RowAmounts, RowStable
'     Benchmark amount and stable funding per factor row and bucket
'
'   WithFigures
'     False writes the factors only (totals withheld)
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   For C80 (4 buckets) then C81 (3 buckets), and for each row code of the
'   replica found in the factor table: writes the regulatory and applied
'   factors; with figures, also the amounts of the amount cells, the
'   effective factor where the amount is not zero (not for aggregate rows)
'   and the row's stable funding
'
' ERROR POLICY
'   No handler: errors propagate to the caller
'
' DEPENDENCIES
'   WithSheet, NamedCell, FindRow2, FactorValue, HasBucket, TextOf
'
' NOTES
'   The row codes sit three columns left of the values block
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.WriteTemplates"

    Dim Sheet                   As Worksheet    'Replica sheet
    Dim Area                    As Range        'Values block of the replica
    Dim Codes                   As Variant      'Row codes of the replica
    Dim Block()                 As Variant      'Values written
    Dim TemplateCode            As String       'Template of the replica
    Dim Buckets                 As Long         'Amount buckets of the template
    Dim FactorRow               As Long         'Row of the factor table
    Dim Stable                  As Double       'Row stable funding
    Dim Present                 As Boolean      'Factor present for the bucket
    Dim Factor                  As Double       'One factor
    Dim TemplateIndex           As Long         '1 for C80, 2 for C81
    Dim RowIndex                As Long         'Row of the replica
    Dim BucketIndex             As Long         'Amount bucket

'------------------------------------------------------------------------------
' FILL THE REPLICAS
'------------------------------------------------------------------------------
    'Fill each replica row by row
        For TemplateIndex = 1 To 2
            If TemplateIndex = 1 Then
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
            For RowIndex = 1 To Area.Rows.Count
                FactorRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ROW, _
                                     TextOf(Codes(RowIndex, 1)))
                If FactorRow > 0 Then
                    Stable = 0
                    For BucketIndex = 1 To Buckets
                        Factor = FactorValue(Factors, FactorRow, FC_REG, BucketIndex, Present)
                        If Present Then Block(RowIndex, Buckets + BucketIndex) = Factor
                        Factor = FactorValue(Factors, FactorRow, FC_APPLIED, BucketIndex, Present)
                        If Present Then Block(RowIndex, 2 * Buckets + BucketIndex) = Factor
                        If WithFigures Then
                            If HasBucket(TextOf(Factors(FactorRow, FC_AMOUNT_CELLS)), BucketIndex) Then
                                Block(RowIndex, BucketIndex) = RowAmounts(FactorRow, BucketIndex)
                            End If
                            If RowAmounts(FactorRow, BucketIndex) <> 0 _
                               And TextOf(Factors(FactorRow, FC_KIND)) <> KIND_AGG Then
                                Block(RowIndex, 2 * Buckets + BucketIndex) = _
                                    RowStable(FactorRow, BucketIndex) / RowAmounts(FactorRow, BucketIndex)
                            End If
                            Stable = Stable + RowStable(FactorRow, BucketIndex)
                        End If
                    Next BucketIndex
                    If WithFigures Then Block(RowIndex, 3 * Buckets + 1) = Stable
                End If
            Next RowIndex
            Area.Value = Block
        Next TemplateIndex

End Sub

Private Sub WriteReconciliation( _
    ByVal Lines As Variant, _
    ByVal LineCount As Long, _
    ByVal Headline As Variant)
'
'------------------------------------------------------------------------------
'                        WRITE THE RECONCILIATION SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Attribution block and lines of Reconciliation
'
' INPUTS
'   Lines, LineCount
'     The reconciliation lines and their count
'
'   Headline
'     The attribution block
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Main.WriteReconciliation"

    Dim Sheet                   As Worksheet    'Reconciliation sheet

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Write the attribution and the status line
        Set Sheet = WithSheet(SH_RECON)
        NamedCell("recHeadline").Value = Headline
        NamedCell("recStatus").Value = "Reconciliation " & Format$(Now, "yyyy-mm-dd hh:nn") & ": " _
                                       & LineCount & " line(s); filter Result for DIFF and MISSING."
    'Write the lines
        If LineCount > 0 Then WriteBlock SH_RECON, "hdrRecon", Lines, LineCount, RL_COLUMNS

End Sub

'
'------------------------------------------------------------------------------
'
'                              APPLICATION STATE
'
'------------------------------------------------------------------------------
'

Private Sub BeginOperation()

'
'------------------------------------------------------------------------------
'                      SWITCH OFF SCREEN AND CALCULATION
'------------------------------------------------------------------------------
' PURPOSE
'   Capture and switch off screen updating, events and automatic
'   calculation
'
' NOTES
'   EndOperation restores the captured settings
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CAPTURE THE SETTINGS
'------------------------------------------------------------------------------
    'Keep the current settings for EndOperation
        mSavedScreen = Application.ScreenUpdating
        mSavedEvents = Application.EnableEvents
        mSavedCalc = Application.Calculation
        mStateCaptured = True

'------------------------------------------------------------------------------
' SWITCH OFF
'------------------------------------------------------------------------------
    'Switch off screen updating, events and automatic calculation
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual

End Sub

Private Sub EndOperation()

'
'------------------------------------------------------------------------------
'                        RESTORE SCREEN AND CALCULATION
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the captured settings, each independently, and recalculate
'
' ERROR POLICY
'   Best-effort: each restore is attempted even if another fails; does
'   nothing when no settings are captured
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Do nothing unless BeginOperation captured the settings
        If Not mStateCaptured Then Exit Sub
    'Attempt each restore even if another fails, then recalculate
        On Error Resume Next
        Application.Calculation = mSavedCalc
        Application.EnableEvents = mSavedEvents
        Application.ScreenUpdating = mSavedScreen
        Application.Calculate
        On Error GoTo 0
    'Mark the settings as restored
        mStateCaptured = False

End Sub
