Attribute VB_Name = "NSFR_API_Functions"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: NSFR_API_Functions
'------------------------------------------------------------------------------
' PURPOSE
'   The prototype's calculation entry point, NSFR_API_Run: validates the
'   request (profile, reporting date, Params tables, bank settings), runs
'   the benchmark on the positions and reconciles the bank's figures
'
' WHY THIS EXISTS
'   The host calls one function for a whole run; the core modules work on
'   arrays only
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - NSFR_API_Run: validation, benchmark and reconciliation of one
'       request
'     - ValidateCategories: the checks of the category table
'
' ERROR POLICY
'   No On Error. An invalid request returns STOPPED with diagnostics; an
'   invalid position returns INVALID and withholds the benchmark totals.
'   Unexpected runtime errors propagate to the host
'
' DEPENDENCIES
'   NSFR_Core_Config, NSFR_Core_Util, NSFR_Core_Engine, NSFR_Core_Reconcile
'
' NOTES
'   Excel VBA7, 64-bit; no references beyond the defaults
'
' PUBLIC SURFACE
'   NSFR_API_Run. Prototype contract: tables are passed as 1-based 2-D
'   arrays in the column order of NSFR_Core_Config; results come back in
'   ByRef arguments. The typed request and result of IMPLEMENTATION.md are
'   not implemented yet (PROTOTYPE.md, Departures)
'
' STATE OWNERSHIP
'   None: no module state, no workbook access. Inputs are read only
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------


'
'------------------------------------------------------------------------------
'
'                                     RUN
'
'------------------------------------------------------------------------------
'

Public Function NSFR_API_Run( _
    ByVal ProfileId As String, _
    ByVal ReportingDateValue As Variant, _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef Categories() As Variant, _
    ByVal CategoryRows As Long, _
    ByRef Summary() As Variant, _
    ByVal SummaryRows As Long, _
    ByRef Positions() As Variant, _
    ByVal PositionRows As Long, _
    ByRef BankCells() As Variant, _
    ByVal BankRows As Long, _
    ByRef BankSettings() As Variant, _
    ByRef RowAmounts() As Double, _
    ByRef RowStable() As Double, _
    ByRef PositionOut() As Variant, _
    ByRef SummaryOut() As Variant, _
    ByRef Totals() As Variant, _
    ByRef RunClass As String, _
    ByRef Counts() As Long, _
    ByRef Lines As Variant, _
    ByRef LineCount As Long, _
    ByRef Headline As Variant, _
    ByRef HasBankFigures As Boolean, _
    ByRef Diag() As Variant, _
    ByRef DiagCount As Long) As String
'
'------------------------------------------------------------------------------
'                VALIDATE, CALCULATE AND RECONCILE ONE REQUEST
'------------------------------------------------------------------------------
' PURPOSE
'   Validates, calculates and reconciles one request
'
' WHY THIS EXISTS
'   The host and the regression run call one entry point for a whole run
'
' INPUTS
'   ProfileId
'     Must be PROFILE_ID (eu-crr-standard)
'
'   ReportingDateValue
'     An Excel date or yyyy-mm-dd text
'
'   Factors, FactorRows
'     The factor table of Params and its number of rows
'
'   Categories, CategoryRows
'     The category table of Params and its number of rows
'
'   Summary, SummaryRows
'     The summary map of Params and its number of rows
'
'   Positions, PositionRows
'     The positions (may be empty) and their number
'
'   BankCells, BankRows
'     The bank's figures (may be empty) and their number
'
'   BankSettings
'     (1 To 4): amount scale (> 0), absolute, relative and factor
'     tolerances (>= 0)
'
'   RowAmounts, RowStable, PositionOut, SummaryOut
'     Output: the outputs of CalculateBenchmark and ComputeSummary
'
'   Totals
'     Output (1 To 3): ASF, RSF, NSFR of the benchmark (Empty unless VALID)
'
'   RunClass
'     Output: REFERENCE or WHAT-IF
'
'   Counts
'     Output (1 To 3): positions accepted, rejected, with a factor override
'
'   Lines, LineCount, Headline
'     Output: the reconciliation of ReconcileBank; left Empty when no bank
'     figure is usable
'
'   HasBankFigures
'     Output: True when a bank figure is usable
'
'   Diag, DiagCount
'     The diagnostic list, appended to here
'
' RETURNS
'   Status: VALID, NO POSITIONS, INVALID: ... or STOPPED: ...
'
' BEHAVIOR
'   Steps:
'     - validates the profile, the reporting date, the Params tables and
'       the bank settings; any error stops the run (STOPPED)
'     - without positions, returns NO POSITIONS with empty outputs
'     - otherwise calculates the benchmark; a rejected position or a zero
'       RSF gives INVALID and empty benchmark outputs
'     - sets the counts, and WHAT-IF when an applied factor differs from
'       the regulatory one or a position overrides its factor
'     - reconciles the bank's figures when there are any
'
' ERROR POLICY
'   No On Error. Request and position problems are reported as
'   diagnostics and in the status; unexpected runtime errors propagate
'
' DEPENDENCIES
'   NSFR_Core_Util: AddDiagnostic, DateOf, NumberOf.
'   NSFR_Core_Engine: ValidateFactorTable, CalculateBenchmark,
'   ComputeSummary. NSFR_Core_Reconcile: ReconcileBank.
'   ValidateCategories
'
' NOTES
'   The outputs of the calculation are those of CalculateBenchmark and
'   ReconcileBank
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_API_Functions.NSFR_API_Run"

    Dim ReportingDate           As Date         'As-of date
    Dim Valid                   As Boolean      'Conversion outcome
    Dim Errors                  As Long         'Request errors
    Dim WhatIfCount             As Long         'Applied factors differing from the regulatory ones
    Dim Accepted                As Long         'Positions accepted
    Dim Rejected                As Long         'Positions rejected
    Dim Overrides               As Long         'Positions with a factor override
    Dim Settings(1 To 4)        As Double       'Bank settings, converted
    Dim TotalAmt()              As Double       'Benchmark amount per row
    Dim TotalVal()              As Double       'Benchmark stable funding per row
    Dim TotalAsf                As Double       'Benchmark ASF
    Dim TotalRsf                As Double       'Benchmark RSF
    Dim Status                  As String       'Run status
    Dim HasBenchmark            As Boolean      'Benchmark totals usable
    Dim LoopIndex               As Long         'Bank setting, then row of the factor table
    Dim BucketIndex             As Long         'Maturity bucket

'------------------------------------------------------------------------------
' RESET OUTPUTS
'------------------------------------------------------------------------------
    'Start from empty totals and counts, a reference run and no bank figures
        ReDim Totals(1 To 3)
        ReDim Counts(1 To 3)
        RunClass = CLASS_REFERENCE
        LineCount = 0
        HasBankFigures = False

'------------------------------------------------------------------------------
' VALIDATE THE REQUEST
'------------------------------------------------------------------------------
    'Require the supported profile
        If ProfileId <> PROFILE_ID Then
            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", 0, "ProfileID", _
                "Profile '" & ProfileId & "' is not supported: the prototype runs " & PROFILE_ID & " only."
            Errors = Errors + 1
        End If
    'Require a reporting date
        ReportingDate = DateOf(ReportingDateValue, Valid)
        If Not Valid Then
            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", 0, "ReportingDate", _
                "Reporting date is blank or not a date."
            Errors = Errors + 1
        End If
    'Require the Params tables, then validate them
        If FactorRows = 0 Or CategoryRows = 0 Or SummaryRows = 0 Then
            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", 0, "Tables", _
                "The factor table, the category table and the summary map must not be empty."
            Errors = Errors + 1
        Else
            Errors = Errors + ValidateFactorTable(Factors, FactorRows, Diag, DiagCount, WhatIfCount)
            Errors = Errors + ValidateCategories(Factors, FactorRows, Categories, CategoryRows, _
                Diag, DiagCount)
        End If
    'Convert the bank settings: a scale above 0, tolerances of 0 or more
        For LoopIndex = 1 To 4
            Settings(LoopIndex) = NumberOf(BankSettings(LoopIndex), Valid)
            If Not Valid Or Settings(LoopIndex) < 0 Or _
                (LoopIndex = 1 And Settings(LoopIndex) <= 0) Then
                AddDiagnostic Diag, DiagCount, SEV_ERROR, "BankCells", 0, _
                    Choose(LoopIndex, "AmountScale", "AbsTolerance", "RelTolerance", _
                    "FactorTolerance"), _
                    "Setting must be a number" & IIf(LoopIndex = 1, " above 0.", " of 0 or more.")
                Errors = Errors + 1
            End If
        Next LoopIndex
    'Stop the run on any request error
        If Errors > 0 Then
            NSFR_API_Run = STATUS_STOPPED & ": " & Errors & " input error(s) on Params or BankCells"
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE THE BENCHMARK
'------------------------------------------------------------------------------
    'Without positions, return empty outputs; otherwise calculate and total
        If PositionRows = 0 Then
            Status = "NO POSITIONS"
            ReDim RowAmounts(1 To FactorRows, 1 To BUCKET_COUNT)
            ReDim RowStable(1 To FactorRows, 1 To BUCKET_COUNT)
            ReDim PositionOut(1 To 1, 1 To PO_COLUMNS)
            ReDim SummaryOut(1 To SummaryRows, 1 To 4)
        Else
            If CalculateBenchmark(Factors, FactorRows, Categories, CategoryRows, Positions, _
                PositionRows, ReportingDate, RowAmounts, RowStable, PositionOut, Diag, DiagCount, _
                Accepted, Rejected, Overrides) Then
                'Total each row over its buckets
                ReDim TotalAmt(1 To FactorRows)
                ReDim TotalVal(1 To FactorRows)
                For LoopIndex = 1 To FactorRows
                    For BucketIndex = 1 To BUCKET_COUNT
                        TotalAmt(LoopIndex) = TotalAmt(LoopIndex) + _
                            RowAmounts(LoopIndex, BucketIndex)
                        TotalVal(LoopIndex) = TotalVal(LoopIndex) + _
                            RowStable(LoopIndex, BucketIndex)
                    Next BucketIndex
                Next LoopIndex
                'Compute the summary and the totals, unless the RSF is zero
                If ComputeSummary(Factors, FactorRows, Summary, SummaryRows, TotalAmt, TotalVal, _
                    SummaryOut, TotalAsf, TotalRsf) Then
                    Status = STATUS_VALID
                    HasBenchmark = True
                    Totals(1) = TotalAsf
                    Totals(2) = TotalRsf
                    Totals(3) = TotalAsf / TotalRsf
                Else
                    Status = STATUS_INVALID & ": required stable funding is zero"
                    AddDiagnostic Diag, DiagCount, SEV_ERROR, "Positions", 0, "RSF", _
                        "Required stable funding is zero: the NSFR is not defined."
                End If
            Else
                Status = STATUS_INVALID & ": " & Rejected & " position(s) rejected, totals withheld"
                ReDim SummaryOut(1 To SummaryRows, 1 To 4)
            End If
            'Withhold the benchmark outputs without usable totals
            If Not HasBenchmark Then
                ReDim RowAmounts(1 To FactorRows, 1 To BUCKET_COUNT)
                ReDim RowStable(1 To FactorRows, 1 To BUCKET_COUNT)
                ReDim SummaryOut(1 To SummaryRows, 1 To 4)
            End If
        End If

'------------------------------------------------------------------------------
' CLASSIFY THE RUN
'------------------------------------------------------------------------------
    'Report the position counts
        Counts(1) = Accepted
        Counts(2) = Rejected
        Counts(3) = Overrides
    'Mark a run with changed factors or overrides as what-if
        If WhatIfCount > 0 Or Overrides > 0 Then RunClass = CLASS_WHATIF

'------------------------------------------------------------------------------
' RECONCILE THE BANK FIGURES
'------------------------------------------------------------------------------
    'Reconcile only when the bank's figures are given
        If BankRows > 0 Then
            HasBankFigures = ReconcileBank(Factors, FactorRows, Summary, SummaryRows, BankCells, _
                BankRows, Settings(1), Settings(2), Settings(3), Settings(4), HasBenchmark, _
                RowAmounts, RowStable, TotalAsf, TotalRsf, Lines, LineCount, _
                Headline, Diag, DiagCount)
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the run status
        NSFR_API_Run = Status

End Function

Private Function ValidateCategories( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef Categories() As Variant, _
    ByVal CategoryRows As Long, _
    ByRef Diag() As Variant, _
    ByRef DiagCount As Long) As Long
'
'------------------------------------------------------------------------------
'                         VALIDATE THE CATEGORY TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that every category is unique, has a known bucket rule and maps
'   to rows of the factor table
'
' WHY THIS EXISTS
'   A position is classified through its category; a faulty category would
'   misplace every position that uses it
'
' INPUTS
'   Factors, FactorRows
'     The factor table and its number of rows, read only
'
'   Categories, CategoryRows
'     The category table and its number of rows, read only
'
'   Diag, DiagCount
'     The diagnostic list, appended to here
'
' RETURNS
'   Number of errors, each reported as a diagnostic
'
' BEHAVIOR
'   For each category:
'     - a blank code, or a code already listed above, is an error
'     - RULE_DERIVATIVE needs no row mapping
'     - RULE_MATURITY, RULE_HQLA, RULE_FIXED1 and RULE_FIXED3 need a row ID
'       in each mapping column from CC_ROW_UNENCUMBERED to before CC_MEMO;
'       each row ID given, CC_MEMO included, must be a row of the
'       category's template in the factor table
'     - any other rule is an error
'
' ERROR POLICY
'   No On Error; problems are reported as diagnostics
'
' DEPENDENCIES
'   NSFR_Core_Util: TextOf, FindRow1, FindRow2, AddDiagnostic.
'   NSFR_Core_Config: the CC_ and FC_ columns and the RULE_ codes
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
    Const PROC_NAME             As String = "NSFR_API_Functions.ValidateCategories"

    Dim Code                    As String       'Category code
    Dim Rule                    As String       'Bucket rule
    Dim TemplateCode            As String       'Category template
    Dim RowId                   As String       'Mapped row ID
    Dim Column                  As Long         'Mapping column
    Dim CategoryIndex           As Long         'Row of the category table

'------------------------------------------------------------------------------
' CHECK EACH CATEGORY
'------------------------------------------------------------------------------
    'Check the code, then the rule and its row mapping
        For CategoryIndex = 1 To CategoryRows
            Code = TextOf(Categories(CategoryIndex, CC_CODE))
            Rule = TextOf(Categories(CategoryIndex, CC_RULE))
            TemplateCode = TextOf(Categories(CategoryIndex, CC_TEMPLATE))
            'Reject a blank or repeated code
            If VBA.Len(Code) = 0 Then
                AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", CategoryIndex, "Categories", _
                    "Category code is blank."
                ValidateCategories = ValidateCategories + 1
            ElseIf FindRow1(Categories, CategoryIndex - 1, CC_CODE, Code) > 0 Then
                AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", CategoryIndex, Code, _
                    "Category listed twice."
                ValidateCategories = ValidateCategories + 1
            End If
            'Check the row mapping that the rule needs
            Select Case Rule
                Case RULE_DERIVATIVE
                    'No row mapping to check for the derivative rule
                Case RULE_MATURITY, RULE_HQLA, RULE_FIXED1, RULE_FIXED3
                    For Column = CC_ROW_UNENCUMBERED To CC_MEMO
                        RowId = TextOf(Categories(CategoryIndex, Column))
                        If VBA.Len(RowId) = 0 And Column < CC_MEMO Then
                            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", CategoryIndex, Code, _
                                "Row ID missing in the category mapping."
                            ValidateCategories = ValidateCategories + 1
                        ElseIf VBA.Len(RowId) > 0 Then
                            If FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ID, _
                                RowId) = 0 Then
                                AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", CategoryIndex, Code, _
                                    "Row " & TemplateCode & " " & RowId & " is not in the factor table."
                                ValidateCategories = ValidateCategories + 1
                            End If
                        End If
                    Next Column
                Case Else
                    AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", CategoryIndex, Code, _
                        "Unknown bucket rule '" & Rule & "'."
                    ValidateCategories = ValidateCategories + 1
            End Select
        Next CategoryIndex

End Function
