Attribute VB_Name = "NSFR_Core_Reconcile"
'==============================================================================
' MODULE: NSFR_Core_Reconcile
'------------------------------------------------------------------------------
' PURPOSE
'   Benchmark the bank's own C 80.00, C 81.00 and C 84.00 figures, cell by
'   cell, in four checks:
'
'     ARITHMETIC      the bank's stable funding against its own amounts x
'                     its own applicable factors, parents against their
'                     children, C 84.00 against C 80.00 and C 81.00;
'     FACTOR          each applicable factor carrying an amount against the
'                     regulatory factor, with the stable-funding impact
'                     amount x (regulatory - bank);
'     AMOUNT          each amount against the benchmark's (positions run);
'     STABLE FUNDING  each row's stable funding against the benchmark's.
'
'   and attribute the NSFR difference in four levels:
'
'     reported     the bank's C 84.00 (or its C 80.00 / C 81.00 totals)
'     re-performed bank amounts x bank factors, summed bottom-up
'     restated     bank amounts x regulatory factors
'     benchmark    the positions run
'
'   so that arithmetic = re-performed - reported, factor effect =
'   restated - re-performed, amount and classification effect =
'   benchmark - restated.
'
'     difference = other - bank; relative = difference / |bank|
'     MATCH      : |difference| <= absolute tolerance, or relative
'                  <= relative tolerance (factors: |difference| <= factor
'                  tolerance)
'
'   A template cell without an applicable-factor cell takes 0% in the
'   template's own formula; the bank's figures are re-performed the same
'   way, and the FACTOR check names such a cell when the regulation gives
'   another factor (Tier 2 at 6M-1Y, EBA Q&A 2021_6017).
'
' PUBLIC SURFACE
'   None outside this VBA project. ReconcileBank is Public for the API.
'
' DEPENDENCIES
'   NSFR_Core_Config, NSFR_Core_Util, NSFR_Core_Engine (FactorValue,
'   ComputeSummary).
'
' STATE OWNERSHIP
'   None: no module state, no workbook access.
'
' ERROR POLICY
'   No On Error. An unusable bank value is a diagnostic, not a VBA error.
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
        Private Const SHEET_BANK        As String = "BankCells"
        Private Const ROLE_AMOUNT       As Long = 1      'Amount cell
        Private Const ROLE_FACTOR       As Long = 2      'Applicable-factor cell
        Private Const ROLE_STABLE       As Long = 3      'Stable-funding cell
        Private Const ROLE_IGNORED      As Long = 4      'Standard-factor cell, not compared
        Private Const ROLE_SUMMARY      As Long = 5      'C 84.00 cell

'------------------------------------------------------------------------------
'                               RECONCILIATION
'------------------------------------------------------------------------------
Public Function ReconcileBank( _
    ByRef Factors() As Variant, _
    ByVal FactorRows As Long, _
    ByRef Summary() As Variant, _
    ByVal SummaryRows As Long, _
    ByRef BankCells() As Variant, _
    ByVal BankRows As Long, _
    ByVal AmountScale As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double, _
    ByVal FactorTolerance As Double, _
    ByVal HasBenchmark As Boolean, _
    ByRef RowAmounts() As Double, _
    ByRef RowStable() As Double, _
    ByVal BenchmarkAsf As Double, _
    ByVal BenchmarkRsf As Double, _
    ByRef Lines As Variant, _
    ByRef LineCount As Long, _
    ByRef Headline As Variant, _
    ByRef Diag() As Variant, _
    ByRef DiagCount As Long) _
    As Boolean
'
'   PURPOSE    Run the four checks and the attribution described in the
'              module header.
'   INPUTS     Factors, Summary: tables of Params, read only. BankCells:
'              (template, row, column, value, source) rows, read only.
'              AmountScale: units of the reporting currency per bank amount
'              (> 0). Tolerances >= 0. HasBenchmark: True when the
'              positions run is VALID; RowAmounts, RowStable and the
'              benchmark totals are then compared.
'   RETURNS    True when at least one bank figure was usable. Lines: a 2-D
'              array (1 To LineCount, 1 To RL_COLUMNS), sorted by check,
'              template, row and column. Headline (1 To 8, 1 To 3): ASF,
'              RSF and NSFR of the four levels (rows 1-4) and of the
'              effects arithmetic, factor, amounts, total (rows 5-8); Empty
'              where a level is not available.
'   UPDATED    2026-11-05
'
    Dim BankAmount()    As Double      'Bank amount per row and bucket, scaled
    Dim AmountSeen()    As Boolean     'Amount cell present in the bank file
    Dim BankFactor()    As Double      'Bank applicable factor per row and bucket
    Dim FactorSeen()    As Boolean     'Factor cell present
    Dim BankStable()    As Double      'Bank stable funding per row, scaled
    Dim StableSeen()    As Boolean     'Stable-funding cell present
    Dim Bank84()        As Double      'Bank C 84.00 per line and column
    Dim Seen84()        As Boolean     'C 84.00 cell present
    Dim Reperformed()   As Double      'Bank amounts x bank factors, per row
    Dim Restated()      As Double      'Bank amounts x regulatory factors, per row
    Dim DerivedAmt()    As Double      'Bank amount, or the children's where a parent has no cell
    Dim RowTotalAmt()   As Double      'Bank amount over all buckets, per row
    Dim RowTotalVal()   As Double      'Bank stable funding (reported, else re-performed)
    Dim Derived()       As Variant     'C 84.00 derived from the bank's rows
    Dim List()          As Variant     'Reconciliation lines
    Dim Keys()          As String      'Sort key per line
    Dim Usable          As Long        'Usable bank figures
    Dim Ignored         As Long        'Figures of rows or templates not compared
    Dim TemplateCode    As String      'Template of a bank cell
    Dim RowCode         As String      'Row code of a bank cell
    Dim ColumnCode      As String      'Column code of a bank cell
    Dim Role            As Long        'ROLE_* of the column
    Dim Bucket          As Long        'Bucket of the column
    Dim Value           As Double      'Bank value
    Dim Valid           As Boolean     'Conversion outcome
    Dim TargetRow       As Long        'Row of the factor table or summary map
    Dim Kind            As String      'Kind of a factor-table row
    Dim MaxDepth        As Long        'Deepest row ID
    Dim Depth           As Long        'Depth being processed
    Dim Children        As Variant     'Row codes of the children
    Dim ChildRow        As Long        'Row of a child
    Dim Expected        As Double      'Expected figure of an arithmetic check
    Dim Amount          As Double      'Bank amount of a bucket
    Dim Effective       As Double      'Bank factor as the template applies it
    Dim HasFactor       As Boolean     'Effective factor known
    Dim RegFactor       As Double      'Regulatory factor
    Dim RegPresent      As Boolean     'Regulatory factor defined
    Dim Total           As Double      'Re-performed stable funding of a row
    Dim RestTotal       As Double      'Restated stable funding of a row
    Dim DerivedAsf      As Double      'C 84.00 ASF from the bank's rows
    Dim DerivedRsf      As Double      'C 84.00 RSF from the bank's rows
    Dim DerivedOk       As Boolean     'Derived ratio defined
    Dim Rsf80           As Long        'Factor-table row of C 80.00 ID 1
    Dim Asf81           As Long        'Factor-table row of C 81.00 ID 2
    Dim Levels(1 To 8, 1 To 3) As Variant   'Headline
    Dim BucketsOf       As Long        'Amount buckets of a template
    Dim EngineStable    As Double      'Benchmark stable funding of a row
    Dim BankRowStable   As Double      'Bank stable funding of a row, 0 if absent
    Dim Measure         As String      'Summary measure
    Dim ValueColumn     As Long        'C 84.00 column of a summary value
    Dim i               As Long
    Dim j               As Long
    Dim b               As Long
    Dim s               As Long

        ReDim BankAmount(1 To FactorRows, 1 To BUCKET_COUNT)
        ReDim AmountSeen(1 To FactorRows, 1 To BUCKET_COUNT)
        ReDim BankFactor(1 To FactorRows, 1 To BUCKET_COUNT)
        ReDim FactorSeen(1 To FactorRows, 1 To BUCKET_COUNT)
        ReDim BankStable(1 To FactorRows)
        ReDim StableSeen(1 To FactorRows)
        ReDim Reperformed(1 To FactorRows)
        ReDim Restated(1 To FactorRows)
        ReDim DerivedAmt(1 To FactorRows, 1 To BUCKET_COUNT)
        ReDim RowTotalAmt(1 To FactorRows)
        ReDim RowTotalVal(1 To FactorRows)
        ReDim Bank84(1 To IIf(SummaryRows > 0, SummaryRows, 1), 1 To 4)
        ReDim Seen84(1 To IIf(SummaryRows > 0, SummaryRows, 1), 1 To 4)
        LineCount = 0

        '--- 1. read the bank cells into the grid
        For i = 1 To BankRows
            TemplateCode = TextOf(BankCells(i, BC_TEMPLATE))
            RowCode = TextOf(BankCells(i, BC_ROW))
            ColumnCode = TextOf(BankCells(i, BC_COLUMN))
            If Len(TemplateCode) = 0 And Len(RowCode) = 0 And Len(ColumnCode) = 0 Then
                'blank line: skipped
            Else
                Role = ColumnRole(TemplateCode, ColumnCode, Bucket)
                If Role = 0 Then
                    Ignored = Ignored + 1
                ElseIf Role = ROLE_IGNORED Then
                    'standard factors of the template are not compared
                Else
                    Value = NumberOf(BankCells(i, BC_VALUE), Valid)
                    If Role = ROLE_SUMMARY Then
                        TargetRow = FindRow1(Summary, SummaryRows, SC_ROW, RowCode)
                    Else
                        TargetRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ROW, RowCode)
                    End If
                    If TargetRow = 0 Then
                        Ignored = Ignored + 1
                    ElseIf Not Valid Then
                        If Not IsBlankValue(BankCells(i, BC_VALUE)) Then
                            AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_BANK, i, _
                                          TemplateCode & " r" & RowCode & " c" & ColumnCode, _
                                          "Bank value is not a number: the cell is not compared."
                        End If
                    ElseIf Role = ROLE_SUMMARY Then
                        If Seen84(TargetRow, Bucket) Then
                            AddDiagnostic Diag, DiagCount, SEV_WARNING, SHEET_BANK, i, _
                                          TemplateCode & " r" & RowCode & " c" & ColumnCode, _
                                          "Cell listed twice: the first value is used."
                        Else
                            If Bucket = 4 Then
                                Bank84(TargetRow, Bucket) = Value               'ratio: not scaled
                            Else
                                Bank84(TargetRow, Bucket) = Value * AmountScale
                            End If
                            Seen84(TargetRow, Bucket) = True
                            Usable = Usable + 1
                        End If
                    ElseIf Role = ROLE_AMOUNT Then
                        If AmountSeen(TargetRow, Bucket) Then
                            AddDiagnostic Diag, DiagCount, SEV_WARNING, SHEET_BANK, i, _
                                          TemplateCode & " r" & RowCode & " c" & ColumnCode, _
                                          "Cell listed twice: the first value is used."
                        Else
                            BankAmount(TargetRow, Bucket) = Value * AmountScale
                            AmountSeen(TargetRow, Bucket) = True
                            Usable = Usable + 1
                        End If
                    ElseIf Role = ROLE_FACTOR Then
                        If FactorSeen(TargetRow, Bucket) Then
                            AddDiagnostic Diag, DiagCount, SEV_WARNING, SHEET_BANK, i, _
                                          TemplateCode & " r" & RowCode & " c" & ColumnCode, _
                                          "Cell listed twice: the first value is used."
                        ElseIf Value < 0 Or Value > 1 Then
                            AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_BANK, i, _
                                          TemplateCode & " r" & RowCode & " c" & ColumnCode, _
                                          "Bank factor outside 0-1: treated as missing."
                        Else
                            BankFactor(TargetRow, Bucket) = Value
                            FactorSeen(TargetRow, Bucket) = True
                            Usable = Usable + 1
                        End If
                    ElseIf Role = ROLE_STABLE Then
                        If StableSeen(TargetRow) Then
                            AddDiagnostic Diag, DiagCount, SEV_WARNING, SHEET_BANK, i, _
                                          TemplateCode & " r" & RowCode & " c" & ColumnCode, _
                                          "Cell listed twice: the first value is used."
                        Else
                            BankStable(TargetRow) = Value * AmountScale
                            StableSeen(TargetRow) = True
                            Usable = Usable + 1
                        End If
                    End If
                End If
            End If
        Next i
        If Ignored > 0 Then
            AddDiagnostic Diag, DiagCount, SEV_INFO, SHEET_BANK, 0, "Bank cells", _
                          Ignored & " bank figure(s) of other templates, rows or columns are not compared " _
                          & "(simplified NSFR C 82.00 / C 83.00 is out of scope)."
        End If
        ReconcileBank = (Usable > 0)
        If Usable = 0 Then
            Lines = ListToMatrix(List, 0, RL_COLUMNS)
            Headline = Levels
            Exit Function
        End If

        '--- 2. rows, deepest first: arithmetic, factors, re-performance, restatement
        For i = 1 To FactorRows
            If IdDepth(TextOf(Factors(i, FC_ID))) > MaxDepth Then MaxDepth = IdDepth(TextOf(Factors(i, FC_ID)))
        Next i
        For Depth = MaxDepth To 0 Step -1
            For i = 1 To FactorRows
                If IdDepth(TextOf(Factors(i, FC_ID))) = Depth Then
                    Kind = TextOf(Factors(i, FC_KIND))
                    TemplateCode = TextOf(Factors(i, FC_TEMPLATE))
                    If Kind = KIND_AGG Then
                        Children = Split(TextOf(Factors(i, FC_CHILDREN)), ",")
                        Total = 0
                        RestTotal = 0
                        Expected = 0
                        For j = LBound(Children) To UBound(Children)
                            ChildRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ROW, _
                                                Trim$(Children(j)))
                            If ChildRow > 0 Then
                                Total = Total + Reperformed(ChildRow)
                                RestTotal = RestTotal + Restated(ChildRow)
                                If StableSeen(ChildRow) Then
                                    Expected = Expected + BankStable(ChildRow)
                                Else
                                    Expected = Expected + Reperformed(ChildRow)
                                End If
                            End If
                        Next j
                        Reperformed(i) = Total
                        Restated(i) = RestTotal
                        For b = 1 To BUCKET_COUNT
                            'A parent without a cell in a bucket carries its children's
                            'amounts upwards (C 80.00 row 1.7 has no HQLA cell, 1.7.3 has).
                            Amount = 0
                            For j = LBound(Children) To UBound(Children)
                                ChildRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ROW, _
                                                    Trim$(Children(j)))
                                If ChildRow > 0 Then Amount = Amount + DerivedAmt(ChildRow, b)
                            Next j
                            If Not AmountSeen(i, b) Then
                                DerivedAmt(i, b) = Amount
                            Else
                                DerivedAmt(i, b) = BankAmount(i, b)
                                AddLine List, LineCount, Keys, CHECK_ARITHMETIC, Factors, i, _
                                        AmountColumn(TemplateCode, b), BankAmount(i, b), Amount, _
                                        ToleranceResult(Amount - BankAmount(i, b), BankAmount(i, b), _
                                                        AbsTolerance, RelTolerance), Empty
                            End If
                        Next b
                    Else
                        Total = 0
                        RestTotal = 0
                        For b = 1 To BUCKET_COUNT
                            DerivedAmt(i, b) = BankAmount(i, b)
                            Amount = BankAmount(i, b)
                            If Amount <> 0 Then
                                Effective = EffectiveFactor(Factors, i, b, BankFactor, FactorSeen, HasFactor)
                                RegFactor = FactorValue(Factors, i, FC_REG, b, RegPresent)
                                If Kind = KIND_MEMO Then
                                    If Not HasFactor Then
                                        'A memo row's factor feeds C 84.00 (2.3 and 2.4).
                                        Effective = 0
                                        AddLine List, LineCount, Keys, CHECK_FACTOR, Factors, i, _
                                                FactorColumn(TemplateCode, b), Empty, Empty, RESULT_MISSING, Empty
                                    End If
                                    RestTotal = RestTotal + Amount * Effective
                                ElseIf Not HasFactor Then
                                    Effective = 0
                                    If RegPresent Then
                                        AddLine List, LineCount, Keys, CHECK_FACTOR, Factors, i, _
                                                FactorColumn(TemplateCode, b), Empty, RegFactor, RESULT_MISSING, _
                                                Amount * RegFactor
                                        RestTotal = RestTotal + Amount * RegFactor
                                    Else
                                        AddLine List, LineCount, Keys, CHECK_FACTOR, Factors, i, _
                                                FactorColumn(TemplateCode, b), Empty, Empty, RESULT_MISSING, Empty
                                    End If
                                ElseIf Not RegPresent Then
                                    AddLine List, LineCount, Keys, CHECK_FACTOR, Factors, i, _
                                            FactorColumn(TemplateCode, b), Effective, Empty, RESULT_INFO, Empty
                                    RestTotal = RestTotal + Amount * Effective
                                Else
                                    If HasBucket(TextOf(Factors(i, FC_FACTOR_CELLS)), b) Then
                                        ColumnCode = FactorColumn(TemplateCode, b)
                                    Else
                                        ColumnCode = AmountColumn(TemplateCode, b)
                                    End If
                                    AddLine List, LineCount, Keys, CHECK_FACTOR, Factors, i, ColumnCode, _
                                            Effective, RegFactor, _
                                            IIf(Abs(Effective - RegFactor) <= FactorTolerance, RESULT_MATCH, RESULT_DIFF), _
                                            Amount * (RegFactor - Effective)
                                    RestTotal = RestTotal + Amount * RegFactor
                                End If
                                Total = Total + Amount * Effective
                            End If
                        Next b
                        Reperformed(i) = Total
                        Restated(i) = RestTotal
                        Expected = Total
                    End If
                    If StableSeen(i) Then
                        AddLine List, LineCount, Keys, CHECK_ARITHMETIC, Factors, i, StableColumn(TemplateCode), _
                                BankStable(i), Expected, _
                                ToleranceResult(Expected - BankStable(i), BankStable(i), AbsTolerance, RelTolerance), _
                                Empty
                    End If
                End If
            Next i
        Next Depth

        '--- 3. C 84.00 against the bank's own rows
        For i = 1 To FactorRows
            RowTotalAmt(i) = 0
            For b = 1 To BUCKET_COUNT
                RowTotalAmt(i) = RowTotalAmt(i) + DerivedAmt(i, b)
            Next b
            If StableSeen(i) Then RowTotalVal(i) = BankStable(i) Else RowTotalVal(i) = Reperformed(i)
        Next i
        DerivedOk = ComputeSummary(Factors, FactorRows, Summary, SummaryRows, RowTotalAmt, RowTotalVal, _
                                   Derived, DerivedAsf, DerivedRsf)
        For s = 1 To SummaryRows
            Measure = TextOf(Summary(s, SC_MEASURE))
            If Measure = "NSFR" Then
                If Seen84(s, 4) And DerivedOk Then
                    AddSummaryLine List, LineCount, Keys, CHECK_ARITHMETIC, Summary, s, "0040", Bank84(s, 4), _
                                   Derived(s, 4), _
                                   IIf(Abs(Derived(s, 4) - Bank84(s, 4)) <= FactorTolerance, RESULT_MATCH, RESULT_DIFF)
                End If
            Else
                If Seen84(s, 1) Then
                    AddSummaryLine List, LineCount, Keys, CHECK_ARITHMETIC, Summary, s, "0010", Bank84(s, 1), _
                                   Derived(s, 1), ToleranceResult(Derived(s, 1) - Bank84(s, 1), Bank84(s, 1), _
                                                                  AbsTolerance, RelTolerance)
                End If
                If Measure = "RSF" Then ValueColumn = 2 Else ValueColumn = 3
                If Seen84(s, ValueColumn) Then
                    AddSummaryLine List, LineCount, Keys, CHECK_ARITHMETIC, Summary, s, "00" & ValueColumn & "0", _
                                   Bank84(s, ValueColumn), Derived(s, ValueColumn), _
                                   ToleranceResult(Derived(s, ValueColumn) - Bank84(s, ValueColumn), _
                                                   Bank84(s, ValueColumn), AbsTolerance, RelTolerance)
                End If
            End If
        Next s

        '--- 4. benchmark comparison
        If HasBenchmark Then
            For i = 1 To FactorRows
                TemplateCode = TextOf(Factors(i, FC_TEMPLATE))
                If TemplateCode = TEMPLATE_RSF Then BucketsOf = 4 Else BucketsOf = 3
                For b = 1 To BucketsOf
                    If HasBucket(TextOf(Factors(i, FC_AMOUNT_CELLS)), b) Then
                        If BankAmount(i, b) <> 0 Or RowAmounts(i, b) <> 0 Then
                            AddLine List, LineCount, Keys, CHECK_AMOUNT, Factors, i, AmountColumn(TemplateCode, b), _
                                    BankAmount(i, b), RowAmounts(i, b), _
                                    ToleranceResult(RowAmounts(i, b) - BankAmount(i, b), BankAmount(i, b), _
                                                    AbsTolerance, RelTolerance), Empty
                        End If
                    End If
                Next b
                EngineStable = 0
                For b = 1 To BUCKET_COUNT
                    EngineStable = EngineStable + RowStable(i, b)
                Next b
                If StableSeen(i) Then BankRowStable = BankStable(i) Else BankRowStable = 0
                If BankRowStable <> 0 Or EngineStable <> 0 Then
                    AddLine List, LineCount, Keys, CHECK_STABLE, Factors, i, StableColumn(TemplateCode), _
                            BankRowStable, EngineStable, _
                            ToleranceResult(EngineStable - BankRowStable, BankRowStable, AbsTolerance, RelTolerance), _
                            Empty
                End If
            Next i
        End If

        '--- 5. headline and attribution
        Rsf80 = FindRow2(Factors, FactorRows, FC_TEMPLATE, TEMPLATE_RSF, FC_ID, "1")
        Asf81 = FindRow2(Factors, FactorRows, FC_TEMPLATE, TEMPLATE_ASF, FC_ID, "2")
        For s = 1 To SummaryRows
            If TextOf(Summary(s, SC_ID)) = "1" And Seen84(s, 2) Then Levels(HL_REPORTED, HL_RSF) = Bank84(s, 2)
            If TextOf(Summary(s, SC_ID)) = "2" And Seen84(s, 3) Then Levels(HL_REPORTED, HL_ASF) = Bank84(s, 3)
            If TextOf(Summary(s, SC_MEASURE)) = "NSFR" And Seen84(s, 4) Then Levels(HL_REPORTED, HL_NSFR) = Bank84(s, 4)
        Next s
        If IsEmpty(Levels(HL_REPORTED, HL_RSF)) And Rsf80 > 0 Then
            If StableSeen(Rsf80) Then Levels(HL_REPORTED, HL_RSF) = BankStable(Rsf80)
        End If
        If IsEmpty(Levels(HL_REPORTED, HL_ASF)) And Asf81 > 0 Then
            If StableSeen(Asf81) Then Levels(HL_REPORTED, HL_ASF) = BankStable(Asf81)
        End If
        If IsEmpty(Levels(HL_REPORTED, HL_NSFR)) Then
            Levels(HL_REPORTED, HL_NSFR) = RatioOf(Levels(HL_REPORTED, HL_ASF), Levels(HL_REPORTED, HL_RSF))
            If SummaryHasCells(Seen84, SummaryRows) Then
                AddDiagnostic Diag, DiagCount, SEV_WARNING, SHEET_BANK, 0, "C 84.00 r0220 c0040", _
                              "The bank's NSFR is missing or not a number: the reported level uses its " _
                              & "C 80.00 / C 81.00 totals."
            End If
        End If
        If Rsf80 > 0 And Asf81 > 0 Then
            Levels(HL_REPERFORMED, HL_ASF) = Reperformed(Asf81)
            Levels(HL_REPERFORMED, HL_RSF) = Reperformed(Rsf80)
            Levels(HL_REPERFORMED, HL_NSFR) = RatioOf(Reperformed(Asf81), Reperformed(Rsf80))
            Levels(HL_RESTATED, HL_ASF) = Restated(Asf81)
            Levels(HL_RESTATED, HL_RSF) = Restated(Rsf80)
            Levels(HL_RESTATED, HL_NSFR) = RatioOf(Restated(Asf81), Restated(Rsf80))
        End If
        If HasBenchmark Then
            Levels(HL_BENCHMARK, HL_ASF) = BenchmarkAsf
            Levels(HL_BENCHMARK, HL_RSF) = BenchmarkRsf
            Levels(HL_BENCHMARK, HL_NSFR) = RatioOf(BenchmarkAsf, BenchmarkRsf)
        End If
        For j = 1 To 3
            Levels(5, j) = DifferenceOf(Levels(HL_REPERFORMED, j), Levels(HL_REPORTED, j))
            Levels(6, j) = DifferenceOf(Levels(HL_RESTATED, j), Levels(HL_REPERFORMED, j))
            Levels(7, j) = DifferenceOf(Levels(HL_BENCHMARK, j), Levels(HL_RESTATED, j))
            Levels(8, j) = DifferenceOf(Levels(HL_BENCHMARK, j), Levels(HL_REPORTED, j))
        Next j
        Headline = Levels

        SortLines List, Keys, LineCount
        Lines = ListToMatrix(List, LineCount, RL_COLUMNS)

End Function

'------------------------------------------------------------------------------
'                                   COLUMNS
'------------------------------------------------------------------------------
Private Function ColumnRole( _
    ByVal TemplateCode As String, _
    ByVal ColumnCode As String, _
    ByRef Bucket As Long) _
    As Long
'
'   PURPOSE    Role and bucket of a column code of the reporting templates.
'   RETURNS    ROLE_* constant, or 0 for a template or column not compared.
'              C 80.00: 0010-0040 amounts, 0050-0080 standard factors,
'              0090-0120 applicable factors, 0130 RSF. C 81.00: 0010-0030
'              amounts, 0040-0060 standard, 0070-0090 applicable, 0100 ASF.
'              C 84.00: 0010-0040, Bucket = column number.
'   UPDATED    2026-11-05
'
    Dim Number          As Long        'Column code / 10

        Bucket = 0
        If Not IsAllDigits(ColumnCode) Or Len(ColumnCode) <> 4 Then Exit Function
        Number = CLng(ColumnCode) \ 10
        If CLng(ColumnCode) Mod 10 <> 0 Then Exit Function
        Select Case TemplateCode
            Case TEMPLATE_RSF
                Select Case Number
                    Case 1 To 4: ColumnRole = ROLE_AMOUNT: Bucket = Number
                    Case 5 To 8: ColumnRole = ROLE_IGNORED
                    Case 9 To 12: ColumnRole = ROLE_FACTOR: Bucket = Number - 8
                    Case 13: ColumnRole = ROLE_STABLE
                End Select
            Case TEMPLATE_ASF
                Select Case Number
                    Case 1 To 3: ColumnRole = ROLE_AMOUNT: Bucket = Number
                    Case 4 To 6: ColumnRole = ROLE_IGNORED
                    Case 7 To 9: ColumnRole = ROLE_FACTOR: Bucket = Number - 6
                    Case 10: ColumnRole = ROLE_STABLE
                End Select
            Case TEMPLATE_SUMMARY
                If Number >= 1 And Number <= 4 Then
                    ColumnRole = ROLE_SUMMARY
                    Bucket = Number
                End If
        End Select

End Function


Private Function AmountColumn( _
    ByVal TemplateCode As String, _
    ByVal Bucket As Long) _
    As String
'
'   PURPOSE    Column code of the amount of a bucket (0010-0040).
'   UPDATED    2026-11-05
'
        AmountColumn = Format$(10 * Bucket, "0000")

End Function


Private Function FactorColumn( _
    ByVal TemplateCode As String, _
    ByVal Bucket As Long) _
    As String
'
'   PURPOSE    Column code of the applicable factor of a bucket: C 80.00
'              0090-0120, C 81.00 0070-0090.
'   UPDATED    2026-11-05
'
        If TemplateCode = TEMPLATE_RSF Then
            FactorColumn = Format$(10 * (Bucket + 8), "0000")
        Else
            FactorColumn = Format$(10 * (Bucket + 6), "0000")
        End If

End Function


Private Function StableColumn( _
    ByVal TemplateCode As String) _
    As String
'
'   PURPOSE    Column code of the stable funding: 0130 or 0100.
'   UPDATED    2026-11-05
'
        If TemplateCode = TEMPLATE_RSF Then StableColumn = "0130" Else StableColumn = "0100"

End Function

'------------------------------------------------------------------------------
'                                   HELPERS
'------------------------------------------------------------------------------
Private Function SummaryHasCells( _
    ByRef Seen84() As Boolean, _
    ByVal SummaryRows As Long) _
    As Boolean
'
'   PURPOSE    True when at least one C 84.00 cell was usable.
'   UPDATED    2026-11-05
'
    Dim s               As Long
    Dim c               As Long

        For s = 1 To SummaryRows
            For c = 1 To 4
                If Seen84(s, c) Then
                    SummaryHasCells = True
                    Exit Function
                End If
            Next c
        Next s

End Function


Private Function EffectiveFactor( _
    ByRef Factors() As Variant, _
    ByVal FactorRow As Long, _
    ByVal Bucket As Long, _
    ByRef BankFactor() As Double, _
    ByRef FactorSeen() As Boolean, _
    ByRef Known As Boolean) _
    As Double
'
'   PURPOSE    The bank's factor as the template applies it: the entered
'              applicable factor where the template has a factor cell, 0
'              where it has none.
'   RETURNS    The factor; Known = False when the template has a factor
'              cell for the bucket and the bank left it blank.
'   UPDATED    2026-11-05
'
        If HasBucket(TextOf(Factors(FactorRow, FC_FACTOR_CELLS)), Bucket) Then
            Known = FactorSeen(FactorRow, Bucket)
            If Known Then EffectiveFactor = BankFactor(FactorRow, Bucket)
        Else
            Known = True
            EffectiveFactor = 0
        End If

End Function


Private Function RatioOf( _
    ByVal Numerator As Variant, _
    ByVal Denominator As Variant) _
    As Variant
'
'   PURPOSE    Numerator / Denominator, Empty when either is Empty or the
'              denominator is not positive.
'   UPDATED    2026-11-05
'
        If IsEmpty(Numerator) Or IsEmpty(Denominator) Then Exit Function
        If CDbl(Denominator) <= 0 Then Exit Function
        RatioOf = CDbl(Numerator) / CDbl(Denominator)

End Function


Private Function DifferenceOf( _
    ByVal Minuend As Variant, _
    ByVal Subtrahend As Variant) _
    As Variant
'
'   PURPOSE    Minuend - Subtrahend, Empty when either is Empty.
'   UPDATED    2026-11-05
'
        If IsEmpty(Minuend) Or IsEmpty(Subtrahend) Then Exit Function
        DifferenceOf = CDbl(Minuend) - CDbl(Subtrahend)

End Function


Private Sub AddLine( _
    ByRef List() As Variant, _
    ByRef LineCount As Long, _
    ByRef Keys() As String, _
    ByVal CheckName As String, _
    ByRef Factors() As Variant, _
    ByVal FactorRow As Long, _
    ByVal ColumnCode As String, _
    ByVal BankValue As Variant, _
    ByVal OtherValue As Variant, _
    ByVal Outcome As String, _
    ByVal Impact As Variant)
'
'   PURPOSE    Append one reconciliation line of a C 80.00 / C 81.00 row.
'   UPDATED    2026-11-05
'
        AppendLine List, LineCount, Keys, CheckName, TextOf(Factors(FactorRow, FC_TEMPLATE)), _
                   TextOf(Factors(FactorRow, FC_ROW)), TextOf(Factors(FactorRow, FC_ID)), _
                   TextOf(Factors(FactorRow, FC_ITEM)), ColumnCode, BankValue, OtherValue, Outcome, Impact

End Sub


Private Sub AddSummaryLine( _
    ByRef List() As Variant, _
    ByRef LineCount As Long, _
    ByRef Keys() As String, _
    ByVal CheckName As String, _
    ByRef Summary() As Variant, _
    ByVal SummaryRow As Long, _
    ByVal ColumnCode As String, _
    ByVal BankValue As Variant, _
    ByVal OtherValue As Variant, _
    ByVal Outcome As String)
'
'   PURPOSE    Append one reconciliation line of a C 84.00 line.
'   UPDATED    2026-11-05
'
        AppendLine List, LineCount, Keys, CheckName, TEMPLATE_SUMMARY, TextOf(Summary(SummaryRow, SC_ROW)), _
                   TextOf(Summary(SummaryRow, SC_ID)), TextOf(Summary(SummaryRow, SC_ITEM)), ColumnCode, _
                   BankValue, OtherValue, Outcome, Empty

End Sub


Private Sub AppendLine( _
    ByRef List() As Variant, _
    ByRef LineCount As Long, _
    ByRef Keys() As String, _
    ByVal CheckName As String, _
    ByVal TemplateCode As String, _
    ByVal RowCode As String, _
    ByVal RowId As String, _
    ByVal Item As String, _
    ByVal ColumnCode As String, _
    ByVal BankValue As Variant, _
    ByVal OtherValue As Variant, _
    ByVal Outcome As String, _
    ByVal Impact As Variant)
'
'   PURPOSE    Build a line with its difference and relative difference,
'              append it and its sort key.
'   UPDATED    2026-11-05
'
    Dim RowItem         As Variant     'Line
    Dim Rank            As String      'Check order

        RowItem = NewRow(RL_COLUMNS)
        RowItem(RL_CHECK) = CheckName
        RowItem(RL_TEMPLATE) = TemplateCode
        RowItem(RL_ROW) = RowCode
        RowItem(RL_ID) = RowId
        RowItem(RL_ITEM) = Item
        RowItem(RL_COLUMN) = ColumnCode
        RowItem(RL_BANK) = BankValue
        RowItem(RL_OTHER) = OtherValue
        If Not IsEmpty(BankValue) And Not IsEmpty(OtherValue) Then
            RowItem(RL_DIFFERENCE) = CDbl(OtherValue) - CDbl(BankValue)
            If CDbl(BankValue) <> 0 Then RowItem(RL_RELATIVE) = RowItem(RL_DIFFERENCE) / Abs(CDbl(BankValue))
        End If
        RowItem(RL_RESULT) = Outcome
        RowItem(RL_IMPACT) = Impact
        ListAdd List, LineCount, RowItem
        Select Case CheckName
            Case CHECK_ARITHMETIC: Rank = "1"
            Case CHECK_FACTOR: Rank = "2"
            Case CHECK_AMOUNT: Rank = "3"
            Case Else: Rank = "4"
        End Select
        If LineCount = 1 Then
            ReDim Keys(1 To UBound(List))
        ElseIf UBound(Keys) < UBound(List) Then
            ReDim Preserve Keys(1 To UBound(List))
        End If
        Keys(LineCount) = Rank & "|" & TemplateCode & "|" & RowCode & "|" & ColumnCode

End Sub


Private Sub SortLines( _
    ByRef List() As Variant, _
    ByRef Keys() As String, _
    ByVal LineCount As Long)
'
'   PURPOSE    Stable insertion sort of the lines by their keys (check,
'              template, row, column).
'   UPDATED    2026-11-05
'
    Dim KeyValue        As String      'Key being inserted
    Dim Item            As Variant     'Line being inserted
    Dim i               As Long
    Dim j               As Long

        For i = 2 To LineCount
            KeyValue = Keys(i)
            Item = List(i)
            j = i - 1
            Do While j >= 1
                If StrComp(Keys(j), KeyValue, vbBinaryCompare) <= 0 Then Exit Do
                Keys(j + 1) = Keys(j)
                List(j + 1) = List(j)
                j = j - 1
            Loop
            Keys(j + 1) = KeyValue
            List(j + 1) = Item
        Next i

End Sub
