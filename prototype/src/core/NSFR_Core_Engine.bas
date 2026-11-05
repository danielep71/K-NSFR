Attribute VB_Name = "NSFR_Core_Engine"
'==============================================================================
' MODULE: NSFR_Core_Engine
'------------------------------------------------------------------------------
' PURPOSE
'   Benchmark calculation of the prototype: classify each position into a
'   row and bucket of C 80.00 or C 81.00, weight it with the applied
'   factor, net the derivatives, aggregate the template hierarchy and
'   derive the C 84.00 summary and the NSFR.
'
'     stable funding of a cell = amount x applied factor
'     NSFR                     = ASF (C 81.00 row 2) / RSF (C 80.00 row 1)
'
' PUBLIC SURFACE
'   None outside this VBA project. ValidateFactorTable, CalculateBenchmark,
'   ComputeSummary and FactorValue are Public for NSFR_API_Functions and
'   NSFR_Core_Reconcile.
'
' DEPENDENCIES
'   NSFR_Core_Config (layout), NSFR_Core_Util (conversion, look-ups, lists).
'
' STATE OWNERSHIP
'   None: no module state, no workbook access. Inputs are read only.
'
' ERROR POLICY
'   No On Error. An unusable position is an ERROR diagnostic and makes the
'   run INVALID: its totals are withheld by the API. Unexpected runtime
'   errors propagate to the host.
'
' REFERENCE
'   docs/methodology/PROTOTYPE.md; CRR Art. 428a-428ah.
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
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const SHEET_PARAMS      As String = "Params"
        Private Const SHEET_POSITIONS   As String = "Positions"

'------------------------------------------------------------------------------
'                                FACTOR TABLE
'------------------------------------------------------------------------------
Public Function FactorValue( _
    ByRef Factors() As Variant, _
    ByVal FactorRow As Long, _
    ByVal FirstColumn As Long, _
    ByVal Bucket As Long, _
    ByRef Present As Boolean) _
    As Double
'
'   PURPOSE    Regulatory (FirstColumn = FC_REG) or applied (FC_APPLIED)
'              factor of one bucket of a factor-table row.
'   RETURNS    The factor; Present = False when the cell is blank or not a
'              number (ValidateFactorTable reports the latter).
'   UPDATED    2026-11-05
'
    Dim Value           As Variant     'Cell value

        Value = Factors(FactorRow, FirstColumn + Bucket - 1)
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
    ByRef WhatIfCount As Long) _
    As Long
'
'   PURPOSE    Check the factor table of Params and count the applied
'              factors that differ from the regulatory ones.
'   INPUTS     Factors: the table, FC_COLUMNS wide, read only.
'   RETURNS    Number of errors (kind unknown, factor not a number or
'              outside 0-1, duplicate row). WhatIfCount: applied factors
'              that differ from a regulatory factor, each listed as INFO.
'   UPDATED    2026-11-05
'
    Dim Kind            As String      'Row kind
    Dim Reference       As String      'Template and ID, for messages
    Dim RegValue        As Double      'Regulatory factor
    Dim AppValue        As Double      'Applied factor
    Dim RegPresent      As Boolean     'Regulatory factor entered
    Dim AppPresent      As Boolean     'Applied factor entered
    Dim Column          As Long        'Factor column under test
    Dim Valid           As Boolean     'Conversion outcome
    Dim i               As Long
    Dim b               As Long

        WhatIfCount = 0
        For i = 1 To FactorRows
            Reference = TextOf(Factors(i, FC_TEMPLATE)) & " " & TextOf(Factors(i, FC_ID))
            Kind = TextOf(Factors(i, FC_KIND))
            If Kind <> KIND_AGG And Kind <> KIND_LEAF And Kind <> KIND_APPL And Kind <> KIND_MEMO Then
                AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_PARAMS, i, Reference, _
                              "Unknown row kind '" & Kind & "'."
                ValidateFactorTable = ValidateFactorTable + 1
            End If
            If FindRow2(Factors, i - 1, FC_TEMPLATE, TextOf(Factors(i, FC_TEMPLATE)), _
                        FC_ROW, TextOf(Factors(i, FC_ROW))) > 0 Then
                AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_PARAMS, i, Reference, _
                              "Row code listed twice in the factor table."
                ValidateFactorTable = ValidateFactorTable + 1
            End If
            For Column = FC_REG To FC_APPLIED + BUCKET_COUNT - 1
                If Not IsBlankValue(Factors(i, Column)) Then
                    AppValue = NumberOf(Factors(i, Column), Valid)
                    If Not Valid Then
                        AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_PARAMS, i, Reference, _
                                      "Factor is not a number."
                        ValidateFactorTable = ValidateFactorTable + 1
                    ElseIf AppValue < 0 Or AppValue > 1 Then
                        AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_PARAMS, i, Reference, _
                                      "Factor outside 0-1 (enter decimals: 0.05, not 5)."
                        ValidateFactorTable = ValidateFactorTable + 1
                    End If
                End If
            Next Column
            For b = 1 To BUCKET_COUNT
                RegValue = FactorValue(Factors, i, FC_REG, b, RegPresent)
                AppValue = FactorValue(Factors, i, FC_APPLIED, b, AppPresent)
                If RegPresent And AppPresent Then
                    If Abs(RegValue - AppValue) > 0.000000000001 Then
                        WhatIfCount = WhatIfCount + 1
                        AddDiagnostic Diag, DiagCount, SEV_INFO, SHEET_PARAMS, i, Reference, _
                                      "WHAT-IF: applied factor " & AppValue & " differs from the regulatory " _
                                      & RegValue & " (" & BucketCode(b) & ")."
                    End If
                ElseIf RegPresent And Not AppPresent Then
                    AddDiagnostic Diag, DiagCount, SEV_WARNING, SHEET_PARAMS, i, Reference, _
                                  "Applied factor blank where a regulatory factor exists (" & BucketCode(b) _
                                  & "): positions in this cell are rejected."
                ElseIf AppPresent And Kind = KIND_APPL Then
                    AddDiagnostic Diag, DiagCount, SEV_INFO, SHEET_PARAMS, i, Reference, _
                                  "Factor set by the competent authority: " & AppValue & " (" & BucketCode(b) & ")."
                End If
            Next b
        Next i

End Function

'------------------------------------------------------------------------------
'                                 CALCULATION
'------------------------------------------------------------------------------
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
    ByRef OverrideCount As Long) _
    As Boolean
'
'   PURPOSE    Classify and weight every position, net the derivatives and
'              aggregate the template rows.
'   INPUTS     Factors, Categories, Positions: 1-based tables in the column
'              order of NSFR_Core_Config, read only. Amounts in units of the
'              reporting currency, >= 0. ReportingDate: the as-of date.
'   RETURNS    True when every position was accepted. RowAmounts and
'              RowStable (1 To FactorRows, 1 To 4): amount and stable
'              funding per row and bucket, parents included. PositionOut
'              (1 To PositionRows, 1 To PO_COLUMNS): PositionCalc rows.
'              Accepted, Rejected: position counts. OverrideCount: positions
'              with a factor override.
'   ERROR POLICY
'              A rejected position is an ERROR diagnostic; the caller
'              withholds the totals.
'   REFERENCE  PROTOTYPE.md, Classification; CRR Art. 428d, 428s(2),
'              428ah(2), 428k(3)(a) for derivatives.
'   UPDATED    2026-11-05
'
    Dim PositionId      As String      'Position ID
    Dim CategoryCode    As String      'Category code
    Dim CategoryRow     As Long        'Row of the category table
    Dim Rule            As String      'Bucket rule of the category
    Dim TemplateCode    As String      'Template of the category
    Dim Amount          As Double      'Carrying amount
    Dim Factor          As Double      'Applied factor
    Dim FactorPresent   As Boolean     'Applied factor found
    Dim Bucket          As Long        'Bucket of the amount
    Dim EncBucket       As Long        'Encumbrance bucket, 1-3
    Dim RowId           As String      'Target row ID
    Dim TargetRow       As Long        'Target row of the factor table
    Dim MemoRow         As Long        'Memo row of the factor table
    Dim MaturityDate    As Date        'Position maturity
    Dim EncumberedDate  As Date        'End of encumbrance
    Dim Valid           As Boolean     'Conversion outcome
    Dim Problem         As String      'Rejection reason; "" when accepted
    Dim Note            As String      'Remark kept with an accepted position
    Dim Override        As Long        'Bucket override
    Dim FvAssets        As Double      'Derivatives with positive fair value
    Dim FvLiabilities   As Double      'Derivatives with negative fair value (absolute)
    Dim VmReceived      As Double      'Eligible variation margin received
    Dim VmPosted        As Double      'Variation margin posted
    Dim NetAssets       As Double      'NSFR derivative assets
    Dim NetLiabilities  As Double      'NSFR derivative liabilities
    Dim OutRow          As Variant     'One PositionCalc row
    Dim Netted          As Boolean     'Derivative input, booked in the netting below
    Dim i               As Long
    Dim c               As Long

        ReDim RowAmounts(1 To FactorRows, 1 To BUCKET_COUNT)
        ReDim RowStable(1 To FactorRows, 1 To BUCKET_COUNT)
        If PositionRows > 0 Then
            ReDim PositionOut(1 To PositionRows, 1 To PO_COLUMNS)
        Else
            ReDim PositionOut(1 To 1, 1 To PO_COLUMNS)
        End If
        Accepted = 0
        Rejected = 0
        OverrideCount = 0

        For i = 1 To PositionRows
            Problem = vbNullString
            Note = vbNullString
            OutRow = NewRow(PO_COLUMNS)
            PositionId = TextOf(Positions(i, PC_ID))
            CategoryCode = TextOf(Positions(i, PC_CATEGORY))
            OutRow(PO_ID) = PositionId
            OutRow(PO_CATEGORY) = CategoryCode
            Bucket = 0
            TargetRow = 0
            MemoRow = 0
            Netted = False

            '--- identity, category and amount
            If Len(PositionId) = 0 Then
                Problem = "PositionID is blank."
            ElseIf FindRow1(Positions, i - 1, PC_ID, PositionId) > 0 Then
                Problem = "PositionID listed twice."
            End If
            If Len(Problem) = 0 Then
                CategoryRow = FindRow1(Categories, CategoryRows, CC_CODE, CategoryCode)
                If CategoryRow = 0 Then Problem = "Unknown category '" & CategoryCode & "'."
            End If
            If Len(Problem) = 0 Then
                Amount = NumberOf(Positions(i, PC_AMOUNT), Valid)
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

            '--- factor override
            If Len(Problem) = 0 And Not IsBlankValue(Positions(i, PC_FACTOR_OVERRIDE)) Then
                Factor = NumberOf(Positions(i, PC_FACTOR_OVERRIDE), Valid)
                If Not Valid Then
                    Problem = "FactorOverride is not a number."
                ElseIf Factor < 0 Or Factor > 1 Then
                    Problem = "FactorOverride outside 0-1 (enter decimals)."
                ElseIf Rule = RULE_DERIVATIVE Then
                    Problem = "FactorOverride does not apply to a derivative input."
                End If
            End If

            '--- derivatives are netted below
            If Len(Problem) = 0 And Rule = RULE_DERIVATIVE Then
                Select Case CategoryCode
                    Case DRV_FV_ASSET: FvAssets = FvAssets + Amount
                    Case DRV_FV_LIAB: FvLiabilities = FvLiabilities + Amount
                    Case DRV_VM_RECEIVED: VmReceived = VmReceived + Amount
                    Case DRV_VM_POSTED: VmPosted = VmPosted + Amount
                    Case Else: Problem = "Derivative category without a netting role."
                End Select
                If Len(Problem) = 0 Then
                    If Not IsBlankValue(Positions(i, PC_ENCUMBERED)) Or _
                       Not IsBlankValue(Positions(i, PC_BUCKET_OVERRIDE)) Then
                        Note = "Maturity, encumbrance and bucket are not used for derivative inputs."
                    End If
                    Netted = True
                End If
            End If

            '--- bucket
            If Len(Problem) = 0 And Not Netted Then
                Override = 0
                If Not IsBlankValue(Positions(i, PC_BUCKET_OVERRIDE)) Then
                    Override = BucketFromCode(TextOf(Positions(i, PC_BUCKET_OVERRIDE)))
                    If Override = 0 Then
                        Problem = "BucketOverride must be <6M, 6M-1Y or >=1Y."
                    ElseIf Rule <> RULE_MATURITY Then
                        Problem = "BucketOverride does not apply to a category with bucket rule " & Rule & "."
                    End If
                End If
            End If
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
                        ElseIf Not IsBlankValue(Positions(i, PC_MATURITY)) Then
                            MaturityDate = DateOf(Positions(i, PC_MATURITY), Valid)
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

            '--- encumbrance
            If Len(Problem) = 0 And Not Netted Then
                EncBucket = 1
                If Not IsBlankValue(Positions(i, PC_ENCUMBERED)) Then
                    EncumberedDate = DateOf(Positions(i, PC_ENCUMBERED), Valid)
                    If Not Valid Then
                        Problem = "EncumberedUntil is not a date."
                    ElseIf TemplateCode = TEMPLATE_ASF Then
                        Problem = "EncumberedUntil applies to assets only."
                    ElseIf EncumberedDate > ReportingDate Then
                        EncBucket = MaturityBucket(ReportingDate, EncumberedDate)
                    End If
                End If
            End If

            '--- target row, cell and factor
            If Len(Problem) = 0 And Not Netted Then
                RowId = TextOf(Categories(CategoryRow, CC_ROW_UNENCUMBERED + EncBucket - 1))
                TargetRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ID, RowId)
                If TargetRow = 0 Then
                    Problem = "Row " & TemplateCode & " " & RowId & " of category " & CategoryCode _
                              & " is not in the factor table."
                ElseIf TextOf(Factors(TargetRow, FC_KIND)) = KIND_AGG Or _
                       TextOf(Factors(TargetRow, FC_KIND)) = KIND_MEMO Then
                    Problem = "Row " & RowId & " is an aggregate or memo row: map the category to a detail row."
                ElseIf Not HasBucket(TextOf(Factors(TargetRow, FC_AMOUNT_CELLS)), Bucket) Then
                    Problem = "Row " & RowId & " has no " & BucketCode(Bucket) & " cell in " & TemplateCode & "."
                End If
            End If
            If Len(Problem) = 0 And Not Netted Then
                If Not IsBlankValue(Positions(i, PC_FACTOR_OVERRIDE)) Then
                    Factor = NumberOf(Positions(i, PC_FACTOR_OVERRIDE), Valid)
                    OverrideCount = OverrideCount + 1
                    Note = Trim$(Note & " WHAT-IF: factor override.")
                Else
                    Factor = FactorValue(Factors, TargetRow, FC_APPLIED, Bucket, FactorPresent)
                    If Not FactorPresent Then
                        Problem = "No applied factor for " & TemplateCode & " " & RowId & " " & BucketCode(Bucket) _
                                  & " on Params."
                    End If
                End If
            End If

            '--- accept
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
                        Note = Trim$(Note & " Not reported in memo row " & TextOf(Factors(MemoRow, FC_ID)) _
                                     & ": no " & BucketCode(Bucket) & " cell.")
                    End If
                End If
                OutRow(PO_STATUS) = "ACCEPTED"
                If Len(Note) > 0 Then OutRow(PO_STATUS) = "ACCEPTED: " & Note
                Accepted = Accepted + 1
            Else
                OutRow(PO_STATUS) = "REJECTED: " & Problem
                AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_POSITIONS, i, PositionId, Problem
                Rejected = Rejected + 1
            End If
            For c = 1 To PO_COLUMNS
                PositionOut(i, c) = OutRow(c)
            Next c
        Next i

        '--- derivatives [Art. 428d, 428s(2), 428ah(2), 428k(3)(a)]
        NetAssets = FvAssets - VmReceived
        If NetAssets < 0 Then NetAssets = 0
        NetLiabilities = FvLiabilities - VmPosted
        If NetLiabilities < 0 Then NetLiabilities = 0
        If Not AddDerivativeRow(Factors, FactorRows, TEMPLATE_RSF, DRV_ROW_LIABILITIES, FvLiabilities, _
                                RowAmounts, RowStable, Diag, DiagCount) Then Rejected = Rejected + 1
        If NetAssets > NetLiabilities Then
            If Not AddDerivativeRow(Factors, FactorRows, TEMPLATE_RSF, DRV_ROW_NET_ASSETS, _
                                    NetAssets - NetLiabilities, RowAmounts, RowStable, Diag, DiagCount) Then _
                Rejected = Rejected + 1
        ElseIf NetLiabilities > NetAssets Then
            If Not AddDerivativeRow(Factors, FactorRows, TEMPLATE_ASF, DRV_ROW_NET_LIABS, _
                                    NetLiabilities - NetAssets, RowAmounts, RowStable, Diag, DiagCount) Then _
                Rejected = Rejected + 1
        End If
        If FvAssets + FvLiabilities + VmReceived + VmPosted > 0 Then
            AddDiagnostic Diag, DiagCount, SEV_INFO, SHEET_POSITIONS, 0, "Derivatives", _
                          "NSFR derivative assets " & FormatAmount(NetAssets) & ", liabilities " _
                          & FormatAmount(NetLiabilities) & "; gross negative fair value " _
                          & FormatAmount(FvLiabilities) & " at 5%."
        End If

        AggregateRows Factors, FactorRows, RowAmounts, RowStable
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
    ByRef DiagCount As Long) _
    As Boolean
'
'   PURPOSE    Book a derivative amount in bucket < 6M of its row with the
'              applied factor.
'   RETURNS    False (with an ERROR diagnostic) when the row or its factor
'              is missing; True otherwise, also when Amount is 0.
'   UPDATED    2026-11-05
'
    Dim TargetRow       As Long        'Row of the factor table
    Dim Factor          As Double      'Applied factor
    Dim Present         As Boolean     'Factor found

        AddDerivativeRow = True
        If Amount = 0 Then Exit Function
        TargetRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ID, RowId)
        If TargetRow = 0 Then
            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", 0, TemplateCode & " " & RowId, _
                          "Derivative row missing from the factor table."
            AddDerivativeRow = False
            Exit Function
        End If
        Factor = FactorValue(Factors, TargetRow, FC_APPLIED, BUCKET_LT6M, Present)
        If Not Present Then
            AddDiagnostic Diag, DiagCount, SEV_ERROR, "Params", TargetRow, TemplateCode & " " & RowId, _
                          "No applied factor for the derivative row."
            AddDerivativeRow = False
            Exit Function
        End If
        RowAmounts(TargetRow, BUCKET_LT6M) = RowAmounts(TargetRow, BUCKET_LT6M) + Amount
        RowStable(TargetRow, BUCKET_LT6M) = RowStable(TargetRow, BUCKET_LT6M) + Amount * Factor

End Function


Public Sub AggregateRows( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef RowAmounts() As Double, _
    ByRef RowStable() As Double)
'
'   PURPOSE    Fill every AGG row with the sum of its children, deepest
'              rows first, so that a parent sums completed children.
'   INPUTS     RowAmounts, RowStable: per row and bucket; AGG rows are
'              overwritten.
'   UPDATED    2026-11-05
'
    Dim MaxDepth        As Long        'Deepest row ID
    Dim Depth           As Long        'Depth being aggregated
    Dim Children        As Variant     'Row codes of the children
    Dim ChildRow        As Long        'Row of a child
    Dim i               As Long
    Dim j               As Long
    Dim b               As Long

        For i = 1 To FactorRows
            If IdDepth(TextOf(Factors(i, FC_ID))) > MaxDepth Then MaxDepth = IdDepth(TextOf(Factors(i, FC_ID)))
        Next i
        For Depth = MaxDepth To 0 Step -1
            For i = 1 To FactorRows
                If TextOf(Factors(i, FC_KIND)) = KIND_AGG And IdDepth(TextOf(Factors(i, FC_ID))) = Depth Then
                    For b = 1 To BUCKET_COUNT
                        RowAmounts(i, b) = 0
                        RowStable(i, b) = 0
                    Next b
                    Children = Split(TextOf(Factors(i, FC_CHILDREN)), ",")
                    For j = LBound(Children) To UBound(Children)
                        ChildRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TextOf(Factors(i, FC_TEMPLATE)), _
                                            FC_ROW, Trim$(Children(j)))
                        If ChildRow > 0 Then
                            For b = 1 To BUCKET_COUNT
                                RowAmounts(i, b) = RowAmounts(i, b) + RowAmounts(ChildRow, b)
                                RowStable(i, b) = RowStable(i, b) + RowStable(ChildRow, b)
                            Next b
                        End If
                    Next j
                End If
            Next i
        Next Depth

End Sub

'------------------------------------------------------------------------------
'                                   SUMMARY
'------------------------------------------------------------------------------
Public Function ComputeSummary( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef Summary() As Variant, _
    ByVal SummaryRows As Long, _
    ByRef RowTotalAmount() As Double, _
    ByRef RowTotalValue() As Double, _
    ByRef SummaryOut() As Variant, _
    ByRef TotalAsf As Double, _
    ByRef TotalRsf As Double) _
    As Boolean
'
'   PURPOSE    C 84.00 lines from per-row totals: amount (column 0010),
'              RSF (0020), ASF (0030) and the ratio (0040).
'   INPUTS     RowTotalAmount, RowTotalValue (1 To FactorRows): amount over
'              all buckets and stable funding of each factor-table row.
'              Summary: the map of Params (plus and minus row lists).
'   RETURNS    True when RSF > 0 and the ratio is defined. SummaryOut
'              (1 To SummaryRows, 1 To 4); TotalAsf, TotalRsf: the
'              AVAILABLE and REQUIRED STABLE FUNDING lines.
'   REFERENCE  C 84.00 of Annex XII, Implementing Regulation (EU) 2021/451.
'   UPDATED    2026-11-05
'
    Dim Measure         As String      'RSF, ASF or NSFR
    Dim Amount          As Double      'Line amount
    Dim Value           As Double      'Line stable funding
    Dim RatioLine       As Long        'Line of the NSFR
    Dim s               As Long

        ReDim SummaryOut(1 To IIf(SummaryRows > 0, SummaryRows, 1), 1 To 4)
        TotalAsf = 0
        TotalRsf = 0
        For s = 1 To SummaryRows
            Measure = TextOf(Summary(s, SC_MEASURE))
            If Measure = "NSFR" Then
                RatioLine = s
            Else
                Amount = SumRowList(Factors, FactorRows, TextOf(Summary(s, SC_PLUS)), RowTotalAmount) _
                         - SumRowList(Factors, FactorRows, TextOf(Summary(s, SC_MINUS)), RowTotalAmount)
                Value = SumRowList(Factors, FactorRows, TextOf(Summary(s, SC_PLUS)), RowTotalValue) _
                        - SumRowList(Factors, FactorRows, TextOf(Summary(s, SC_MINUS)), RowTotalValue)
                SummaryOut(s, 1) = Amount
                If Measure = "RSF" Then
                    SummaryOut(s, 2) = Value
                    If TextOf(Summary(s, SC_ID)) = "1" Then TotalRsf = Value
                Else
                    SummaryOut(s, 3) = Value
                    If TextOf(Summary(s, SC_ID)) = "2" Then TotalAsf = Value
                End If
            End If
        Next s
        If TotalRsf > 0 Then
            If RatioLine > 0 Then SummaryOut(RatioLine, 4) = TotalAsf / TotalRsf
            ComputeSummary = True
        End If

End Function


Private Function SumRowList( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByVal ListText As String, _
    ByRef Totals() As Double) _
    As Double
'
'   PURPOSE    Sum of Totals over a list "C 81.00|2.3;C 81.00|2.7" of
'              template-qualified row IDs; an unknown entry adds nothing.
'   UPDATED    2026-11-05
'
    Dim Entries         As Variant     'List entries
    Dim Parts           As Variant     'Template and ID of one entry
    Dim FoundRow        As Long        'Row of the factor table
    Dim j               As Long

        If Len(ListText) = 0 Then Exit Function
        Entries = Split(ListText, ";")
        For j = LBound(Entries) To UBound(Entries)
            Parts = Split(Entries(j), "|")
            If UBound(Parts) = 1 Then
                FoundRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, Trim$(Parts(0)), FC_ID, Trim$(Parts(1)))
                If FoundRow > 0 Then SumRowList = SumRowList + Totals(FoundRow)
            End If
        Next j

End Function
