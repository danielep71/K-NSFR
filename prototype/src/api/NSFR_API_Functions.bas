Attribute VB_Name = "NSFR_API_Functions"
'==============================================================================
' MODULE: NSFR_API_Functions
'------------------------------------------------------------------------------
' PURPOSE
'   The prototype's calculation entry point, NSFR_API_Run: validate the
'   request (profile, reporting date, Params tables, bank settings), run
'   the benchmark on the positions and reconcile the bank's figures.
'
' PUBLIC SURFACE
'   NSFR_API_Run. Prototype contract: tables are passed as 1-based 2-D
'   arrays in the column order of NSFR_Core_Config; results come back in
'   ByRef arguments. The typed request and result of IMPLEMENTATION.md are
'   not implemented yet (PROTOTYPE.md, Departures).
'
' DEPENDENCIES
'   NSFR_Core_Config, NSFR_Core_Util, NSFR_Core_Engine, NSFR_Core_Reconcile.
'
' STATE OWNERSHIP
'   None: no module state, no workbook access. Inputs are read only.
'
' ERROR POLICY
'   No On Error. An invalid request returns STOPPED with diagnostics; an
'   invalid position returns INVALID and withholds the benchmark totals.
'   Unexpected runtime errors propagate to the host.
'
' COMPATIBILITY
'   Excel VBA7, 64-bit; no references beyond the defaults.
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
'                                     RUN
'------------------------------------------------------------------------------
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
    ByRef DiagCount As Long) _
    As String
'
'   PURPOSE    Validate, calculate and reconcile one request.
'   INPUTS     ProfileId: must be eu-crr-standard. ReportingDateValue: an
'              Excel date or yyyy-mm-dd text. Factors, Categories, Summary:
'              Params tables. Positions: the positions (may be empty).
'              BankCells: the bank's figures (may be empty). BankSettings
'              (1 To 4): amount scale (> 0), absolute, relative and factor
'              tolerances (>= 0).
'   RETURNS    Status: VALID, NO POSITIONS, INVALID: ... or STOPPED: ....
'              Totals (1 To 3): ASF, RSF, NSFR of the benchmark (Empty
'              unless VALID). RunClass: REFERENCE or WHAT-IF. Counts (1 To
'              3): positions accepted, rejected, with a factor override.
'              The other outputs are those of CalculateBenchmark and
'              ReconcileBank; Lines and Headline are left Empty when no
'              bank figure is usable (HasBankFigures = False).
'   UPDATED    2026-11-05
'
    Dim ReportingDate   As Date        'As-of date
    Dim Valid           As Boolean     'Conversion outcome
    Dim Errors          As Long        'Request errors
    Dim WhatIfCount     As Long        'Applied factors differing from the regulatory ones
    Dim Accepted        As Long        'Positions accepted
    Dim Rejected        As Long        'Positions rejected
    Dim Overrides       As Long        'Positions with a factor override
    Dim Settings(1 To 4) As Double     'Bank settings, converted
    Dim TotalAmt()      As Double      'Benchmark amount per row
    Dim TotalVal()      As Double      'Benchmark stable funding per row
    Dim TotalAsf        As Double      'Benchmark ASF
    Dim TotalRsf        As Double      'Benchmark RSF
    Dim Status          As String      'Run status
    Dim HasBenchmark    As Boolean     'Benchmark totals usable
    Dim i               As Long
    Dim b               As Long

        ReDim Totals(1 To 3)
        ReDim Counts(1 To 3)
        RunClass = CLASS_REFERENCE
        LineCount = 0
        HasBankFigures = False

        '--- request
        If ProfileId <> PROFILE_ID Then
            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", 0, "ProfileID", _
                          "Profile '" & ProfileId & "' is not supported: the prototype runs " & PROFILE_ID & " only."
            Errors = Errors + 1
        End If
        ReportingDate = DateOf(ReportingDateValue, Valid)
        If Not Valid Then
            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", 0, "ReportingDate", _
                          "Reporting date is blank or not a date."
            Errors = Errors + 1
        End If
        If FactorRows = 0 Or CategoryRows = 0 Or SummaryRows = 0 Then
            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", 0, "Tables", _
                          "The factor table, the category table and the summary map must not be empty."
            Errors = Errors + 1
        Else
            Errors = Errors + ValidateFactorTable(Factors, FactorRows, Diag, DiagCount, WhatIfCount)
            Errors = Errors + ValidateCategories(Factors, FactorRows, Categories, CategoryRows, Diag, DiagCount)
        End If
        For i = 1 To 4
            Settings(i) = NumberOf(BankSettings(i), Valid)
            If Not Valid Or Settings(i) < 0 Or (i = 1 And Settings(i) <= 0) Then
                AddDiagnostic Diag, DiagCount, SEV_ERROR, "BankCells", 0, _
                              Choose(i, "AmountScale", "AbsTolerance", "RelTolerance", "FactorTolerance"), _
                              "Setting must be a number" & IIf(i = 1, " above 0.", " of 0 or more.")
                Errors = Errors + 1
            End If
        Next i
        If Errors > 0 Then
            NSFR_API_Run = STATUS_STOPPED & ": " & Errors & " input error(s) on Params or BankCells"
            Exit Function
        End If

        '--- benchmark
        If PositionRows = 0 Then
            Status = "NO POSITIONS"
            ReDim RowAmounts(1 To FactorRows, 1 To BUCKET_COUNT)
            ReDim RowStable(1 To FactorRows, 1 To BUCKET_COUNT)
            ReDim PositionOut(1 To 1, 1 To PO_COLUMNS)
            ReDim SummaryOut(1 To SummaryRows, 1 To 4)
        Else
            If CalculateBenchmark(Factors, FactorRows, Categories, CategoryRows, Positions, PositionRows, _
                                  ReportingDate, RowAmounts, RowStable, PositionOut, Diag, DiagCount, _
                                  Accepted, Rejected, Overrides) Then
                ReDim TotalAmt(1 To FactorRows)
                ReDim TotalVal(1 To FactorRows)
                For i = 1 To FactorRows
                    For b = 1 To BUCKET_COUNT
                        TotalAmt(i) = TotalAmt(i) + RowAmounts(i, b)
                        TotalVal(i) = TotalVal(i) + RowStable(i, b)
                    Next b
                Next i
                If ComputeSummary(Factors, FactorRows, Summary, SummaryRows, TotalAmt, TotalVal, SummaryOut, _
                                  TotalAsf, TotalRsf) Then
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
            If Not HasBenchmark Then
                ReDim RowAmounts(1 To FactorRows, 1 To BUCKET_COUNT)
                ReDim RowStable(1 To FactorRows, 1 To BUCKET_COUNT)
                ReDim SummaryOut(1 To SummaryRows, 1 To 4)
            End If
        End If
        Counts(1) = Accepted
        Counts(2) = Rejected
        Counts(3) = Overrides
        If WhatIfCount > 0 Or Overrides > 0 Then RunClass = CLASS_WHATIF

        '--- reconciliation
        If BankRows > 0 Then
            HasBankFigures = ReconcileBank(Factors, FactorRows, Summary, SummaryRows, BankCells, BankRows, _
                                           Settings(1), Settings(2), Settings(3), Settings(4), HasBenchmark, _
                                           RowAmounts, RowStable, TotalAsf, TotalRsf, Lines, LineCount, _
                                           Headline, Diag, DiagCount)
        End If
        NSFR_API_Run = Status

End Function


Private Function ValidateCategories( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef Categories() As Variant, _
    ByVal CategoryRows As Long, _
    ByRef Diag() As Variant, _
    ByRef DiagCount As Long) _
    As Long
'
'   PURPOSE    Check that every category is unique, has a known bucket
'              rule and maps to rows of the factor table.
'   RETURNS    Number of errors, each reported as a diagnostic.
'   UPDATED    2026-11-05
'
    Dim Code            As String      'Category code
    Dim Rule            As String      'Bucket rule
    Dim TemplateCode    As String      'Category template
    Dim RowId           As String      'Mapped row ID
    Dim Column          As Long        'Mapping column
    Dim i               As Long

        For i = 1 To CategoryRows
            Code = TextOf(Categories(i, CC_CODE))
            Rule = TextOf(Categories(i, CC_RULE))
            TemplateCode = TextOf(Categories(i, CC_TEMPLATE))
            If Len(Code) = 0 Then
                AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", i, "Categories", "Category code is blank."
                ValidateCategories = ValidateCategories + 1
            ElseIf FindRow1(Categories, i - 1, CC_CODE, Code) > 0 Then
                AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", i, Code, "Category listed twice."
                ValidateCategories = ValidateCategories + 1
            End If
            Select Case Rule
                Case RULE_DERIVATIVE
                Case RULE_MATURITY, RULE_HQLA, RULE_FIXED1, RULE_FIXED3
                    For Column = CC_ROW_UNENCUMBERED To CC_MEMO
                        RowId = TextOf(Categories(i, Column))
                        If Len(RowId) = 0 And Column < CC_MEMO Then
                            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", i, Code, _
                                          "Row ID missing in the category mapping."
                            ValidateCategories = ValidateCategories + 1
                        ElseIf Len(RowId) > 0 Then
                            If FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ID, RowId) = 0 Then
                                AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", i, Code, _
                                              "Row " & TemplateCode & " " & RowId & " is not in the factor table."
                                ValidateCategories = ValidateCategories + 1
                            End If
                        End If
                    Next Column
                Case Else
                    AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", i, Code, _
                                  "Unknown bucket rule '" & Rule & "'."
                    ValidateCategories = ValidateCategories + 1
            End Select
        Next i

End Function
