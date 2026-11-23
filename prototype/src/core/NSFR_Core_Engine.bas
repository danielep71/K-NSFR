Attribute VB_Name = "NSFR_Core_Engine"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: NSFR_Core_Engine
'------------------------------------------------------------------------------
' PURPOSE
'   Benchmark calculation of the prototype: classifies each position into a
'   row and bucket of C 80.00 or C 81.00, weights it with the applied
'   factor, nets the derivatives, aggregates the template hierarchy and
'   derives the C 84.00 summary and the NSFR
'
' WHY THIS EXISTS
'   The benchmark is computed from the positions and the Params tables
'   alone, with no workbook access, so that the API and the reconciliation
'   can compare a bank's figures with it
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Computes:
'
'       stable funding of a cell = amount x applied factor
'       NSFR                     = ASF (C 81.00 row 2) / RSF (C 80.00 row 1)
'
'   Provides:
'     - factor table: FactorValue, ValidateFactorTable
'     - calculation: CalculateBenchmark, AggregateRows
'     - summary: ComputeSummary
'
' ERROR POLICY
'   No On Error. An unusable position is an ERROR diagnostic and makes the
'   run INVALID: its totals are withheld by the API. Unexpected runtime
'   errors propagate to the host
'
' DEPENDENCIES
'   NSFR_Core_Config (layout), NSFR_Core_Util (conversion, look-ups, lists)
'
' NOTES
'   Excel VBA7, 64-bit; no references beyond the defaults
'
' PUBLIC SURFACE
'   None outside this VBA project. ValidateFactorTable, CalculateBenchmark,
'   ComputeSummary and FactorValue are Public for NSFR_API_Functions and
'   NSFR_Core_Reconcile; NSFR_Host_Main also calls FactorValue.
'   AggregateRows is Public but called only from this module
'
' STATE OWNERSHIP
'   None: no module state, no workbook access. Inputs are read only
'
' REFERENCE
'   docs/methodology/PROTOTYPE.md; CRR Art. 428a-428ah
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const SHEET_PARAMS          As String = "Params"        'Sheet named in factor-table diagnostics
    Private Const SHEET_POSITIONS       As String = "Positions"     'Sheet named in position diagnostics


'
'------------------------------------------------------------------------------
'
'                                 FACTOR TABLE
'
'------------------------------------------------------------------------------
'

Public Function FactorValue( _
    ByRef Factors() As Variant, _
    ByVal FactorRow As Long, _
    ByVal FirstColumn As Long, _
    ByVal Bucket As Long, _
    ByRef Present As Boolean) As Double
'
'------------------------------------------------------------------------------
'                               READ ONE FACTOR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the regulatory (FirstColumn = FC_REG) or applied (FC_APPLIED)
'   factor of one bucket of a factor-table row
'
' WHY THIS EXISTS
'   The validation, the calculation, the reconciliation and the host read
'   factor cells the same way, telling a blank cell from a number
'
' INPUTS
'   Factors
'     The factor table, FC_COLUMNS wide, read only
'
'   FactorRow
'     Row of the factor table
'
'   FirstColumn
'     Column of bucket 1: FC_REG or FC_APPLIED
'
'   Bucket
'     Bucket, 1 to BUCKET_COUNT
'
'   Present
'     Output: False when the cell is blank or not a number
'
' RETURNS
'   The factor; 0 when Present is False
'
' BEHAVIOR
'   Reads column FirstColumn + Bucket - 1. A blank cell sets Present to
'   False; any other value is converted by NumberOf, which sets Present
'
' ERROR POLICY
'   No On Error; runtime errors propagate to the caller
'
' DEPENDENCIES
'   NSFR_Core_Util: IsBlankValue, NumberOf
'
' NOTES
'   ValidateFactorTable reports a factor that is not a number
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Core_Engine.FactorValue"

    Dim Value                   As Variant      'Cell value

'------------------------------------------------------------------------------
' READ THE FACTOR
'------------------------------------------------------------------------------
    'Read the cell of the bucket
        Value = Factors(FactorRow, FirstColumn + Bucket - 1)
    'Report a blank cell as absent, else convert the value
        If IsBlankValue(Value) Then
            Present = False
        Else
            FactorValue = NumberOf(Value, Present)
        End If

End Function

Public Function ValidateFactorTable( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef Diag() As Variant, _
    ByRef DiagCount As Long, _
    ByRef WhatIfCount As Long) As Long
'
'------------------------------------------------------------------------------
'                          VALIDATE THE FACTOR TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Checks the factor table of Params and counts the applied factors that
'   differ from the regulatory ones
'
' WHY THIS EXISTS
'   The calculation reads its factors from this table; the check reports
'   what makes the table unusable and lists every WHAT-IF factor
'
' INPUTS
'   Factors
'     The table, FC_COLUMNS wide, read only
'
'   FactorRows
'     Number of rows of the table
'
'   Diag
'     The diagnostics list, extended by AddDiagnostic
'
'   DiagCount
'     Number of diagnostics in Diag, updated
'
'   WhatIfCount
'     Output: applied factors that differ from a regulatory factor, each
'     listed as INFO
'
' RETURNS
'   Number of errors (kind unknown, factor not a number or outside 0-1,
'   duplicate row)
'
' BEHAVIOR
'   For each row, in order:
'     - an unknown kind is an ERROR
'     - a row code already listed for the same template is an ERROR
'     - each factor cell that is not blank must be a number from 0 to 1
'     - per bucket: an applied factor that differs from the regulatory one
'       by more than 1E-12 is a WHAT-IF INFO; a blank applied factor where
'       a regulatory one exists is a WARNING; an applied factor of an APPL
'       row is an INFO
'
' ERROR POLICY
'   No On Error; table problems are diagnostics, runtime errors propagate
'
' DEPENDENCIES
'   FactorValue; NSFR_Core_Util: TextOf, FindRow2, IsBlankValue, NumberOf,
'   AddDiagnostic, BucketCode
'
' NOTES
'   Diagnostics name the 1-based row of the factor table
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Core_Engine.ValidateFactorTable"

    Dim Kind                    As String       'Row kind
    Dim Reference               As String       'Template and ID, for messages
    Dim RegValue                As Double       'Regulatory factor
    Dim AppValue                As Double       'Applied factor
    Dim RegPresent              As Boolean      'Regulatory factor entered
    Dim AppPresent              As Boolean      'Applied factor entered
    Dim Column                  As Long         'Factor column under test
    Dim Valid                   As Boolean      'Conversion outcome
    Dim RowIndex                As Long         'Row of the factor table
    Dim BucketIndex             As Long         'Bucket under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the WHAT-IF count
        WhatIfCount = 0

'------------------------------------------------------------------------------
' CHECK EACH ROW
'------------------------------------------------------------------------------
    'Check the kind, the row code and the factors of every row
        For RowIndex = 1 To FactorRows
            'Name the row for messages and read its kind
            Reference = TextOf(Factors(RowIndex, FC_TEMPLATE)) & " " & TextOf(Factors(RowIndex, FC_ID))
            Kind = TextOf(Factors(RowIndex, FC_KIND))
            'Reject an unknown row kind
            If Kind <> KIND_AGG And Kind <> KIND_LEAF And Kind <> KIND_APPL And Kind <> KIND_MEMO Then
                AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_PARAMS, RowIndex, Reference, _
                    "Unknown row kind '" & Kind & "'."
                ValidateFactorTable = ValidateFactorTable + 1
            End If
            'Reject a row code already listed for the same template
            If FindRow2(Factors, RowIndex - 1, FC_TEMPLATE, TextOf(Factors(RowIndex, FC_TEMPLATE)), _
                FC_ROW, TextOf(Factors(RowIndex, FC_ROW))) > 0 Then
                AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_PARAMS, RowIndex, Reference, _
                    "Row code listed twice in the factor table."
                ValidateFactorTable = ValidateFactorTable + 1
            End If
            'Require every factor entered to be a number from 0 to 1
            For Column = FC_REG To FC_APPLIED + BUCKET_COUNT - 1
                If Not IsBlankValue(Factors(RowIndex, Column)) Then
                    AppValue = NumberOf(Factors(RowIndex, Column), Valid)
                    If Not Valid Then
                        AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_PARAMS, RowIndex, Reference, _
                            "Factor is not a number."
                        ValidateFactorTable = ValidateFactorTable + 1
                    ElseIf AppValue < 0 Or AppValue > 1 Then
                        AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_PARAMS, RowIndex, Reference, _
                            "Factor outside 0-1 (enter decimals: 0.05, not 5)."
                        ValidateFactorTable = ValidateFactorTable + 1
                    End If
                End If
            Next Column
            'Compare the applied factor with the regulatory one in each bucket
            For BucketIndex = 1 To BUCKET_COUNT
                RegValue = FactorValue(Factors, RowIndex, FC_REG, BucketIndex, RegPresent)
                AppValue = FactorValue(Factors, RowIndex, FC_APPLIED, BucketIndex, AppPresent)
                If RegPresent And AppPresent Then
                    If Abs(RegValue - AppValue) > 0.000000000001 Then
                        WhatIfCount = WhatIfCount + 1
                        AddDiagnostic Diag, DiagCount, SEV_INFO, SHEET_PARAMS, RowIndex, Reference, _
                            "WHAT-IF: applied factor " & AppValue & " differs from the regulatory " & _
                            RegValue & " (" & BucketCode(BucketIndex) & ")."
                    End If
                ElseIf RegPresent And Not AppPresent Then
                    AddDiagnostic Diag, DiagCount, SEV_WARNING, SHEET_PARAMS, RowIndex, Reference, _
                        "Applied factor blank where a regulatory factor exists (" & _
                        BucketCode(BucketIndex) & "): positions in this cell are rejected."
                ElseIf AppPresent And Kind = KIND_APPL Then
                    AddDiagnostic Diag, DiagCount, SEV_INFO, SHEET_PARAMS, RowIndex, Reference, _
                        "Factor set by the competent authority: " & AppValue & " (" & _
                        BucketCode(BucketIndex) & ")."
                End If
            Next BucketIndex
        Next RowIndex

End Function

'
'------------------------------------------------------------------------------
'
'                                 CALCULATION
'
'------------------------------------------------------------------------------
'

Public Function CalculateBenchmark( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef Categories() As Variant, _
    ByVal CategoryRows As Long, _
    ByRef Positions() As Variant, _
    ByVal PositionRows As Long, _
    ByVal ReportingDate As Date, _
    ByRef RowAmounts() As Double, _
    ByRef RowStable() As Double, _
    ByRef PositionOut() As Variant, _
    ByRef Diag() As Variant, _
    ByRef DiagCount As Long, _
    ByRef Accepted As Long, _
    ByRef Rejected As Long, _
    ByRef OverrideCount As Long) As Boolean
'
'------------------------------------------------------------------------------
'                           CALCULATE THE BENCHMARK
'------------------------------------------------------------------------------
' PURPOSE
'   Classifies and weights every position, nets the derivatives and
'   aggregates the template rows
'
' WHY THIS EXISTS
'   The benchmark gives the amount and stable funding of every template
'   cell from the positions, for the summary and the reconciliation
'
' INPUTS
'   Factors, Categories, Positions
'     1-based tables in the column order of NSFR_Core_Config, read only.
'     Amounts in units of the reporting currency, >= 0
'
'   FactorRows, CategoryRows, PositionRows
'     Number of rows of each table
'
'   ReportingDate
'     The as-of date
'
'   RowAmounts, RowStable
'     Output (1 To FactorRows, 1 To 4): amount and stable funding per row
'     and bucket, parents included
'
'   PositionOut
'     Output (1 To PositionRows, 1 To PO_COLUMNS): PositionCalc rows
'
'   Diag, DiagCount
'     The diagnostics list and its count, extended by AddDiagnostic
'
'   Accepted, Rejected
'     Output: position counts
'
'   OverrideCount
'     Output: positions with a factor override
'
' RETURNS
'   True when every position was accepted
'
' BEHAVIOR
'   For each position, in order: identity, category and amount; factor
'   override; derivative inputs set aside for netting; bucket from the
'   category's bucket rule; encumbrance; target row, cell and factor; then
'   the position is booked (with its memo row) or rejected. The derivatives
'   are then netted and booked, and the AGG rows aggregated
'
' ERROR POLICY
'   A rejected position is an ERROR diagnostic; the caller withholds the
'   totals. No On Error; runtime errors propagate
'
' DEPENDENCIES
'   AddDerivativeRow, FactorValue, AggregateRows; NSFR_Core_Util: NewRow,
'   TextOf, FindRow1, FindRow2, NumberOf, DateOf, IsBlankValue,
'   BucketFromCode, MaturityBucket, HasBucket, BucketCode, AddDiagnostic,
'   FormatAmount
'
' NOTES
'   A derivative row that cannot be booked also counts in Rejected
'
' REFERENCE
'   PROTOTYPE.md, Classification; CRR Art. 428d, 428s(2), 428ah(2),
'   428k(3)(a) for derivatives
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Core_Engine.CalculateBenchmark"

    Dim PositionId              As String       'Position ID
    Dim CategoryCode            As String       'Category code
    Dim CategoryRow             As Long         'Row of the category table
    Dim Rule                    As String       'Bucket rule of the category
    Dim TemplateCode            As String       'Template of the category
    Dim Amount                  As Double       'Carrying amount
    Dim Factor                  As Double       'Applied factor
    Dim FactorPresent           As Boolean      'Applied factor found
    Dim Bucket                  As Long         'Bucket of the amount
    Dim EncBucket               As Long         'Encumbrance bucket, 1-3
    Dim RowId                   As String       'Target row ID
    Dim TargetRow               As Long         'Target row of the factor table
    Dim MemoRow                 As Long         'Memo row of the factor table
    Dim MaturityDate            As Date         'Position maturity
    Dim EncumberedDate          As Date         'End of encumbrance
    Dim Valid                   As Boolean      'Conversion outcome
    Dim Problem                 As String       'Rejection reason; "" when accepted
    Dim Note                    As String       'Remark kept with an accepted position
    Dim Override                As Long         'Bucket override
    Dim FvAssets                As Double       'Derivatives with positive fair value
    Dim FvLiabilities           As Double       'Derivatives with negative fair value (absolute)
    Dim VmReceived              As Double       'Eligible variation margin received
    Dim VmPosted                As Double       'Variation margin posted
    Dim NetAssets               As Double       'NSFR derivative assets
    Dim NetLiabilities          As Double       'NSFR derivative liabilities
    Dim OutRow                  As Variant      'One PositionCalc row
    Dim Netted                  As Boolean      'Derivative input, booked in the netting below
    Dim PositionIndex           As Long         'Row of the position table
    Dim ColumnIndex             As Long         'Column of PositionCalc

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Size the outputs, keeping at least one PositionCalc row
        ReDim RowAmounts(1 To FactorRows, 1 To BUCKET_COUNT)
        ReDim RowStable(1 To FactorRows, 1 To BUCKET_COUNT)
        If PositionRows > 0 Then
            ReDim PositionOut(1 To PositionRows, 1 To PO_COLUMNS)
        Else
            ReDim PositionOut(1 To 1, 1 To PO_COLUMNS)
        End If
    'Reset the counts
        Accepted = 0
        Rejected = 0
        OverrideCount = 0

'------------------------------------------------------------------------------
' CLASSIFY THE POSITIONS
'------------------------------------------------------------------------------
    'Classify, weight and book each position, or reject it
        For PositionIndex = 1 To PositionRows
            'Start the position's PositionCalc row
            Problem = vbNullString
            Note = vbNullString
            OutRow = NewRow(PO_COLUMNS)
            PositionId = TextOf(Positions(PositionIndex, PC_ID))
            CategoryCode = TextOf(Positions(PositionIndex, PC_CATEGORY))
            OutRow(PO_ID) = PositionId
            OutRow(PO_CATEGORY) = CategoryCode
            Bucket = 0
            TargetRow = 0
            MemoRow = 0
            Netted = False

            'Check the identity, the category and the amount
            If Len(PositionId) = 0 Then
                Problem = "PositionID is blank."
            ElseIf FindRow1(Positions, PositionIndex - 1, PC_ID, PositionId) > 0 Then
                Problem = "PositionID listed twice."
            End If
            If Len(Problem) = 0 Then
                CategoryRow = FindRow1(Categories, CategoryRows, CC_CODE, CategoryCode)
                If CategoryRow = 0 Then Problem = "Unknown category '" & CategoryCode & "'."
            End If
            If Len(Problem) = 0 Then
                Amount = NumberOf(Positions(PositionIndex, PC_AMOUNT), Valid)
                If Not Valid Then
                    Problem = "Amount is blank or not a number."
                ElseIf Amount < 0 Then
                    Problem = "Amount is negative: enter the carrying amount, derivatives as absolute fair values."
                End If
            End If
            If Len(Problem) = 0 Then
                Rule = TextOf(Categories(CategoryRow, CC_RULE))
                TemplateCode = TextOf(Categories(CategoryRow, CC_TEMPLATE))
                OutRow(PO_TEMPLATE) = TemplateCode
                OutRow(PO_AMOUNT) = Amount
            End If

            'Check the factor override
            If Len(Problem) = 0 And Not IsBlankValue(Positions(PositionIndex, PC_FACTOR_OVERRIDE)) Then
                Factor = NumberOf(Positions(PositionIndex, PC_FACTOR_OVERRIDE), Valid)
                If Not Valid Then
                    Problem = "FactorOverride is not a number."
                ElseIf Factor < 0 Or Factor > 1 Then
                    Problem = "FactorOverride outside 0-1 (enter decimals)."
                ElseIf Rule = RULE_DERIVATIVE Then
                    Problem = "FactorOverride does not apply to a derivative input."
                End If
            End If

            'Set a derivative input aside: derivatives are netted below
            If Len(Problem) = 0 And Rule = RULE_DERIVATIVE Then
                Select Case CategoryCode
                    Case DRV_FV_ASSET
                        FvAssets = FvAssets + Amount
                    Case DRV_FV_LIAB
                        FvLiabilities = FvLiabilities + Amount
                    Case DRV_VM_RECEIVED
                        VmReceived = VmReceived + Amount
                    Case DRV_VM_POSTED
                        VmPosted = VmPosted + Amount
                    Case Else
                        Problem = "Derivative category without a netting role."
                End Select
                If Len(Problem) = 0 Then
                    If Not IsBlankValue(Positions(PositionIndex, PC_ENCUMBERED)) Or _
                        Not IsBlankValue(Positions(PositionIndex, PC_BUCKET_OVERRIDE)) Then
                        Note = "Maturity, encumbrance and bucket are not used for derivative inputs."
                    End If
                    Netted = True
                End If
            End If

            'Check the bucket override
            If Len(Problem) = 0 And Not Netted Then
                Override = 0
                If Not IsBlankValue(Positions(PositionIndex, PC_BUCKET_OVERRIDE)) Then
                    Override = BucketFromCode(TextOf(Positions(PositionIndex, PC_BUCKET_OVERRIDE)))
                    If Override = 0 Then
                        Problem = "BucketOverride must be <6M, 6M-1Y or >=1Y."
                    ElseIf Rule <> RULE_MATURITY Then
                        Problem = "BucketOverride does not apply to a category with bucket rule " & Rule & "."
                    End If
                End If
            End If
            'Set the bucket from the category's bucket rule
            If Len(Problem) = 0 And Not Netted Then
                Select Case Rule
                    Case RULE_HQLA
                        Bucket = BUCKET_HQLA
                    Case RULE_FIXED1
                        Bucket = BUCKET_LT6M
                    Case RULE_FIXED3
                        Bucket = BUCKET_GE1Y
                    Case RULE_MATURITY
                        If Override > 0 Then
                            Bucket = Override
                            Note = "Bucket set by override."
                        ElseIf Not IsBlankValue(Positions(PositionIndex, PC_MATURITY)) Then
                            MaturityDate = DateOf(Positions(PositionIndex, PC_MATURITY), Valid)
                            If Not Valid Then
                                Problem = "MaturityDate is not a date."
                            Else
                                Bucket = MaturityBucket(ReportingDate, MaturityDate)
                                If MaturityDate < ReportingDate Then
                                    Note = "Maturity before the reporting date: treated as < 6 months."
                                End If
                            End If
                        ElseIf IsBlankValue(Categories(CategoryRow, CC_UNDATED)) Then
                            Problem = "MaturityDate required for category " & CategoryCode & "."
                        Else
                            Bucket = CLng(Val(TextOf(Categories(CategoryRow, CC_UNDATED))))
                            If Bucket < BUCKET_LT6M Or Bucket > BUCKET_GE1Y Then
                                Problem = "Undated bucket of category " & CategoryCode & " is not 1, 2 or 3."
                            End If
                        End If
                    Case Else
                        Problem = "Category " & CategoryCode & " has an unknown bucket rule '" & Rule & "'."
                End Select
            End If

            'Set the encumbrance bucket, which selects the row
            If Len(Problem) = 0 And Not Netted Then
                EncBucket = 1
                If Not IsBlankValue(Positions(PositionIndex, PC_ENCUMBERED)) Then
                    EncumberedDate = DateOf(Positions(PositionIndex, PC_ENCUMBERED), Valid)
                    If Not Valid Then
                        Problem = "EncumberedUntil is not a date."
                    ElseIf TemplateCode = TEMPLATE_ASF Then
                        Problem = "EncumberedUntil applies to assets only."
                    ElseIf EncumberedDate > ReportingDate Then
                        EncBucket = MaturityBucket(ReportingDate, EncumberedDate)
                    End If
                End If
            End If

            'Find the target row and its cell
            If Len(Problem) = 0 And Not Netted Then
                RowId = TextOf(Categories(CategoryRow, CC_ROW_UNENCUMBERED + EncBucket - 1))
                TargetRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ID, RowId)
                If TargetRow = 0 Then
                    Problem = "Row " & TemplateCode & " " & RowId & " of category " & CategoryCode & _
                        " is not in the factor table."
                ElseIf TextOf(Factors(TargetRow, FC_KIND)) = KIND_AGG Or _
                    TextOf(Factors(TargetRow, FC_KIND)) = KIND_MEMO Then
                    Problem = "Row " & RowId & " is an aggregate or memo row: map the category to a detail row."
                ElseIf Not HasBucket(TextOf(Factors(TargetRow, FC_AMOUNT_CELLS)), Bucket) Then
                    Problem = "Row " & RowId & " has no " & BucketCode(Bucket) & " cell in " & TemplateCode & "."
                End If
            End If
            'Take the factor override, else the applied factor of the cell
            If Len(Problem) = 0 And Not Netted Then
                If Not IsBlankValue(Positions(PositionIndex, PC_FACTOR_OVERRIDE)) Then
                    Factor = NumberOf(Positions(PositionIndex, PC_FACTOR_OVERRIDE), Valid)
                    OverrideCount = OverrideCount + 1
                    Note = Trim$(Note & " WHAT-IF: factor override.")
                Else
                    Factor = FactorValue(Factors, TargetRow, FC_APPLIED, Bucket, FactorPresent)
                    If Not FactorPresent Then
                        Problem = "No applied factor for " & TemplateCode & " " & RowId & " " & _
                            BucketCode(Bucket) & " on Params."
                    End If
                End If
            End If

            'Accept a netted derivative input, book an accepted position, or reject it
            If Len(Problem) = 0 And Netted Then
                OutRow(PO_STATUS) = "ACCEPTED: netted with the derivatives"
                If Len(Note) > 0 Then OutRow(PO_STATUS) = OutRow(PO_STATUS) & "; " & Note
                Accepted = Accepted + 1
            ElseIf Len(Problem) = 0 Then
                RowAmounts(TargetRow, Bucket) = RowAmounts(TargetRow, Bucket) + Amount
                RowStable(TargetRow, Bucket) = RowStable(TargetRow, Bucket) + Amount * Factor
                OutRow(PO_ROW) = TextOf(Factors(TargetRow, FC_ROW))
                OutRow(PO_ROW_ID) = RowId
                OutRow(PO_BUCKET) = BucketCode(Bucket)
                OutRow(PO_ENCUMBRANCE) = Choose(EncBucket, "Unencumbered or < 6M", "6M-1Y", ">=1Y")
                OutRow(PO_FACTOR) = Factor
                OutRow(PO_STABLE_FUNDING) = Amount * Factor
                'Report the amount in the category's memo row, with the main row's factor
                If Not IsBlankValue(Categories(CategoryRow, CC_MEMO)) Then
                    MemoRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ID, _
                        TextOf(Categories(CategoryRow, CC_MEMO)))
                    If MemoRow = 0 Then
                        Note = Trim$(Note & " Memo row not in the factor table.")
                    ElseIf HasBucket(TextOf(Factors(MemoRow, FC_AMOUNT_CELLS)), Bucket) Then
                        RowAmounts(MemoRow, Bucket) = RowAmounts(MemoRow, Bucket) + Amount
                        If Len(TextOf(Factors(MemoRow, FC_FACTOR_CELLS))) > 0 Then
                            RowStable(MemoRow, Bucket) = RowStable(MemoRow, Bucket) + Amount * Factor
                        End If
                        OutRow(PO_MEMO) = TextOf(Factors(MemoRow, FC_ID))
                    Else
                        Note = Trim$(Note & " Not reported in memo row " & TextOf(Factors(MemoRow, FC_ID)) & _
                            ": no " & BucketCode(Bucket) & " cell.")
                    End If
                End If
                OutRow(PO_STATUS) = "ACCEPTED"
                If Len(Note) > 0 Then OutRow(PO_STATUS) = "ACCEPTED: " & Note
                Accepted = Accepted + 1
            Else
                OutRow(PO_STATUS) = "REJECTED: " & Problem
                AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_POSITIONS, PositionIndex, PositionId, _
                    Problem
                Rejected = Rejected + 1
            End If

            'Copy the PositionCalc row to the output
            For ColumnIndex = 1 To PO_COLUMNS
                PositionOut(PositionIndex, ColumnIndex) = OutRow(ColumnIndex)
            Next ColumnIndex
        Next PositionIndex

'------------------------------------------------------------------------------
' NET THE DERIVATIVES
'------------------------------------------------------------------------------
    'Net the fair values with the variation margin [Art. 428d, 428s(2), 428ah(2), 428k(3)(a)]
        NetAssets = FvAssets - VmReceived
        If NetAssets < 0 Then NetAssets = 0
        NetLiabilities = FvLiabilities - VmPosted
        If NetLiabilities < 0 Then NetLiabilities = 0
    'Book the gross negative fair value in its row
        If Not AddDerivativeRow(Factors, FactorRows, TEMPLATE_RSF, DRV_ROW_LIABILITIES, _
            FvLiabilities, RowAmounts, RowStable, Diag, DiagCount) Then Rejected = Rejected + 1
    'Book the net derivative assets or the net derivative liabilities
        If NetAssets > NetLiabilities Then
            If Not AddDerivativeRow(Factors, FactorRows, TEMPLATE_RSF, DRV_ROW_NET_ASSETS, _
                NetAssets - NetLiabilities, RowAmounts, RowStable, Diag, DiagCount) Then _
                Rejected = Rejected + 1
        ElseIf NetLiabilities > NetAssets Then
            If Not AddDerivativeRow(Factors, FactorRows, TEMPLATE_ASF, DRV_ROW_NET_LIABS, _
                NetLiabilities - NetAssets, RowAmounts, RowStable, Diag, DiagCount) Then _
                Rejected = Rejected + 1
        End If
    'Report the netting as INFO when there is any derivative input
        If FvAssets + FvLiabilities + VmReceived + VmPosted > 0 Then
            AddDiagnostic Diag, DiagCount, SEV_INFO, SHEET_POSITIONS, 0, "Derivatives", _
                "NSFR derivative assets " & FormatAmount(NetAssets) & ", liabilities " & _
                FormatAmount(NetLiabilities) & "; gross negative fair value " & _
                FormatAmount(FvLiabilities) & " at 5%."
        End If

'------------------------------------------------------------------------------
' AGGREGATE THE TEMPLATE ROWS
'------------------------------------------------------------------------------
    'Fill every AGG row with the sum of its children
        AggregateRows Factors, FactorRows, RowAmounts, RowStable

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Succeed only when nothing was rejected
        CalculateBenchmark = (Rejected = 0)

End Function

Private Function AddDerivativeRow( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByVal TemplateCode As String, _
    ByVal RowId As String, _
    ByVal Amount As Double, _
    ByRef RowAmounts() As Double, _
    ByRef RowStable() As Double, _
    ByRef Diag() As Variant, _
    ByRef DiagCount As Long) As Boolean
'
'------------------------------------------------------------------------------
'                           BOOK A DERIVATIVE AMOUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Books a derivative amount in bucket < 6M of its row with the applied
'   factor
'
' WHY THIS EXISTS
'   The netting books up to three derivative rows the same way
'
' INPUTS
'   Factors, FactorRows
'     The factor table and its number of rows, read only
'
'   TemplateCode, RowId
'     Template and row ID of the derivative row
'
'   Amount
'     The amount to book
'
'   RowAmounts, RowStable
'     Amount and stable funding per row and bucket, updated
'
'   Diag, DiagCount
'     The diagnostics list and its count, extended by AddDiagnostic
'
' RETURNS
'   False (with an ERROR diagnostic) when the row or its factor is missing;
'   True otherwise, also when Amount is 0
'
' BEHAVIOR
'   Books nothing for a zero amount; otherwise finds the row, reads its
'   applied factor in bucket < 6M and adds the amount and the amount x
'   factor
'
' ERROR POLICY
'   No On Error; a missing row or factor is an ERROR diagnostic
'
' DEPENDENCIES
'   FactorValue; NSFR_Core_Util: FindRow2, AddDiagnostic
'
' NOTES
'   The diagnostics name the sheet with the literal "Params"
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Core_Engine.AddDerivativeRow"

    Dim TargetRow               As Long         'Row of the factor table
    Dim Factor                  As Double       'Applied factor
    Dim Present                 As Boolean      'Factor found

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Succeed by default and book nothing for a zero amount
        AddDerivativeRow = True
        If Amount = 0 Then Exit Function

'------------------------------------------------------------------------------
' FIND THE ROW AND ITS FACTOR
'------------------------------------------------------------------------------
    'Reject a row missing from the factor table
        TargetRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ID, RowId)
        If TargetRow = 0 Then
            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", 0, TemplateCode & " " & RowId, _
                "Derivative row missing from the factor table."
            AddDerivativeRow = False
            Exit Function
        End If
    'Reject a row without an applied factor in bucket < 6M
        Factor = FactorValue(Factors, TargetRow, FC_APPLIED, BUCKET_LT6M, Present)
        If Not Present Then
            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", TargetRow, TemplateCode & " " & RowId, _
                "No applied factor for the derivative row."
            AddDerivativeRow = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' BOOK THE AMOUNT
'------------------------------------------------------------------------------
    'Add the amount and its stable funding to bucket < 6M
        RowAmounts(TargetRow, BUCKET_LT6M) = RowAmounts(TargetRow, BUCKET_LT6M) + Amount
        RowStable(TargetRow, BUCKET_LT6M) = RowStable(TargetRow, BUCKET_LT6M) + Amount * Factor

End Function

Public Sub AggregateRows( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef RowAmounts() As Double, _
    ByRef RowStable() As Double)
'
'------------------------------------------------------------------------------
'                         AGGREGATE THE TEMPLATE ROWS
'------------------------------------------------------------------------------
' PURPOSE
'   Fills every AGG row with the sum of its children, deepest rows first,
'   so that a parent sums completed children
'
' WHY THIS EXISTS
'   Positions are booked in detail rows only; the parents of the template
'   hierarchy are derived from them
'
' INPUTS
'   Factors, FactorRows
'     The factor table and its number of rows, read only
'
'   RowAmounts, RowStable
'     Amount and stable funding per row and bucket; AGG rows are
'     overwritten
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Finds the deepest row ID, then, from that depth up to 0, resets each
'   AGG row of the depth and adds the rows whose codes its Children column
'   lists. A child code not found in the same template adds nothing
'
' ERROR POLICY
'   No On Error; runtime errors propagate
'
' DEPENDENCIES
'   NSFR_Core_Util: TextOf, IdDepth, FindRow2
'
' NOTES
'   Children are matched by row code (FC_ROW) within the parent's template
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Core_Engine.AggregateRows"

    Dim MaxDepth                As Long         'Deepest row ID
    Dim Depth                   As Long         'Depth being aggregated
    Dim Children                As Variant      'Row codes of the children
    Dim ChildRow                As Long         'Row of a child
    Dim RowIndex                As Long         'Row of the factor table
    Dim ChildIndex              As Long         'Index into Children
    Dim BucketIndex             As Long         'Bucket being summed

'------------------------------------------------------------------------------
' FIND THE DEEPEST ROW
'------------------------------------------------------------------------------
    'Take the largest depth of the row IDs
        For RowIndex = 1 To FactorRows
            If IdDepth(TextOf(Factors(RowIndex, FC_ID))) > MaxDepth Then _
                MaxDepth = IdDepth(TextOf(Factors(RowIndex, FC_ID)))
        Next RowIndex

'------------------------------------------------------------------------------
' AGGREGATE BY DEPTH
'------------------------------------------------------------------------------
    'Sum the children of each AGG row, deepest rows first
        For Depth = MaxDepth To 0 Step -1
            For RowIndex = 1 To FactorRows
                If TextOf(Factors(RowIndex, FC_KIND)) = KIND_AGG And _
                    IdDepth(TextOf(Factors(RowIndex, FC_ID))) = Depth Then
                    'Reset the row, then add each child found in the same template
                    For BucketIndex = 1 To BUCKET_COUNT
                        RowAmounts(RowIndex, BucketIndex) = 0
                        RowStable(RowIndex, BucketIndex) = 0
                    Next BucketIndex
                    Children = Split(TextOf(Factors(RowIndex, FC_CHILDREN)), ",")
                    For ChildIndex = LBound(Children) To UBound(Children)
                        ChildRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, _
                            TextOf(Factors(RowIndex, FC_TEMPLATE)), FC_ROW, _
                            Trim$(Children(ChildIndex)))
                        If ChildRow > 0 Then
                            For BucketIndex = 1 To BUCKET_COUNT
                                RowAmounts(RowIndex, BucketIndex) = _
                                    RowAmounts(RowIndex, BucketIndex) + RowAmounts(ChildRow, BucketIndex)
                                RowStable(RowIndex, BucketIndex) = _
                                    RowStable(RowIndex, BucketIndex) + RowStable(ChildRow, BucketIndex)
                            Next BucketIndex
                        End If
                    Next ChildIndex
                End If
            Next RowIndex
        Next Depth

End Sub

'
'------------------------------------------------------------------------------
'
'                                   SUMMARY
'
'------------------------------------------------------------------------------
'

Public Function ComputeSummary( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef Summary() As Variant, _
    ByVal SummaryRows As Long, _
    ByRef RowTotalAmount() As Double, _
    ByRef RowTotalValue() As Double, _
    ByRef SummaryOut() As Variant, _
    ByRef TotalAsf As Double, _
    ByRef TotalRsf As Double) As Boolean
'
'------------------------------------------------------------------------------
'                         COMPUTE THE C 84.00 SUMMARY
'------------------------------------------------------------------------------
' PURPOSE
'   C 84.00 lines from per-row totals: amount (column 0010), RSF (0020),
'   ASF (0030) and the ratio (0040)
'
' WHY THIS EXISTS
'   The summary map of Params defines each C 84.00 line as sums and
'   differences of factor-table rows; the benchmark and the reconciliation
'   both derive C 84.00 from it
'
' INPUTS
'   Factors, FactorRows
'     The factor table and its number of rows, read only
'
'   Summary, SummaryRows
'     The summary map of Params (plus and minus row lists) and its number
'     of rows
'
'   RowTotalAmount, RowTotalValue
'     (1 To FactorRows): amount over all buckets and stable funding of
'     each factor-table row
'
'   SummaryOut
'     Output (1 To SummaryRows, 1 To 4)
'
'   TotalAsf, TotalRsf
'     Output: the AVAILABLE and REQUIRED STABLE FUNDING lines
'
' RETURNS
'   True when RSF > 0 and the ratio is defined
'
' BEHAVIOR
'   For each line other than NSFR: amount and stable funding are the plus
'   list minus the minus list; an RSF line fills column 2, any other line
'   column 3. The RSF line with ID 1 gives TotalRsf, the ASF line with ID 2
'   gives TotalAsf. The NSFR line receives TotalAsf / TotalRsf when TotalRsf
'   is above 0
'
' ERROR POLICY
'   No On Error; runtime errors propagate
'
' DEPENDENCIES
'   SumRowList; NSFR_Core_Util: TextOf
'
' NOTES
'   At least one SummaryOut row is allocated when SummaryRows is 0
'
' REFERENCE
'   C 84.00 of Annex XII, Implementing Regulation (EU) 2021/451
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Core_Engine.ComputeSummary"

    Dim Measure                 As String       'RSF, ASF or NSFR
    Dim Amount                  As Double       'Line amount
    Dim Value                   As Double       'Line stable funding
    Dim RatioLine               As Long         'Line of the NSFR
    Dim LineIndex               As Long         'Line of the summary map

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Size the output and reset the totals
        ReDim SummaryOut(1 To IIf(SummaryRows > 0, SummaryRows, 1), 1 To 4)
        TotalAsf = 0
        TotalRsf = 0

'------------------------------------------------------------------------------
' SUM THE LINES
'------------------------------------------------------------------------------
    'Sum each RSF and ASF line and remember the NSFR line
        For LineIndex = 1 To SummaryRows
            Measure = TextOf(Summary(LineIndex, SC_MEASURE))
            If Measure = "NSFR" Then
                RatioLine = LineIndex
            Else
                Amount = SumRowList(Factors, FactorRows, TextOf(Summary(LineIndex, SC_PLUS)), _
                    RowTotalAmount) - SumRowList(Factors, FactorRows, _
                    TextOf(Summary(LineIndex, SC_MINUS)), RowTotalAmount)
                Value = SumRowList(Factors, FactorRows, TextOf(Summary(LineIndex, SC_PLUS)), _
                    RowTotalValue) - SumRowList(Factors, FactorRows, _
                    TextOf(Summary(LineIndex, SC_MINUS)), RowTotalValue)
                SummaryOut(LineIndex, 1) = Amount
                If Measure = "RSF" Then
                    SummaryOut(LineIndex, 2) = Value
                    If TextOf(Summary(LineIndex, SC_ID)) = "1" Then TotalRsf = Value
                Else
                    SummaryOut(LineIndex, 3) = Value
                    If TextOf(Summary(LineIndex, SC_ID)) = "2" Then TotalAsf = Value
                End If
            End If
        Next LineIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Set the ratio and succeed only when the RSF is above zero
        If TotalRsf > 0 Then
            If RatioLine > 0 Then SummaryOut(RatioLine, 4) = TotalAsf / TotalRsf
            ComputeSummary = True
        End If

End Function

Private Function SumRowList( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByVal ListText As String, _
    ByRef Totals() As Double) As Double
'
'------------------------------------------------------------------------------
'                              SUM A LIST OF ROWS
'------------------------------------------------------------------------------
' PURPOSE
'   Sum of Totals over a list "C 81.00|2.3;C 81.00|2.7" of
'   template-qualified row IDs; an unknown entry adds nothing
'
' RETURNS
'   The sum; 0 for an empty list
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Core_Engine.SumRowList"

    Dim Entries                 As Variant      'List entries
    Dim Parts                   As Variant      'Template and ID of one entry
    Dim FoundRow                As Long         'Row of the factor table
    Dim EntryIndex              As Long         'Index into Entries

'------------------------------------------------------------------------------
' SUM THE ENTRIES
'------------------------------------------------------------------------------
    'Return 0 for an empty list
        If Len(ListText) = 0 Then Exit Function
    'Add the total of each "template|ID" entry found in the factor table
        Entries = Split(ListText, ";")
        For EntryIndex = LBound(Entries) To UBound(Entries)
            Parts = Split(Entries(EntryIndex), "|")
            If UBound(Parts) = 1 Then
                FoundRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, Trim$(Parts(0)), FC_ID, _
                    Trim$(Parts(1)))
                If FoundRow > 0 Then SumRowList = SumRowList + Totals(FoundRow)
            End If
        Next EntryIndex

End Function
