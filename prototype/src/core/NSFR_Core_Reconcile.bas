Attribute VB_Name = "NSFR_Core_Reconcile"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: NSFR_Core_Reconcile
'------------------------------------------------------------------------------
' PURPOSE
'   Benchmarks the bank's own C 80.00, C 81.00 and C 84.00 figures, cell by
'   cell, in four checks, and attributes the NSFR difference in four levels
'
' WHY THIS EXISTS
'   The attribution splits the difference between the bank's NSFR and the
'   benchmark into arithmetic = re-performed - reported, factor effect =
'   restated - re-performed, and amount and classification effect =
'   benchmark - restated
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   The four checks:
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
'   The four levels of the attribution:
'
'     reported     the bank's C 84.00 (or its C 80.00 / C 81.00 totals)
'     re-performed bank amounts x bank factors, summed bottom-up
'     restated     bank amounts x regulatory factors
'     benchmark    the positions run
'
'   The result of a comparison:
'
'     difference = other - bank; relative = difference / |bank|
'     MATCH      : |difference| <= absolute tolerance, or relative
'                  <= relative tolerance (factors: |difference| <= factor
'                  tolerance)
'
' ERROR POLICY
'   No On Error. An unusable bank value is a diagnostic, not a VBA error
'
' DEPENDENCIES
'   NSFR_Core_Config, NSFR_Core_Util, NSFR_Core_Engine (FactorValue,
'   ComputeSummary)
'
' NOTES
'   A template cell without an applicable-factor cell takes 0% in the
'   template's own formula; the bank's figures are re-performed the same
'   way, and the FACTOR check names such a cell when the regulation gives
'   another factor (Tier 2 at 6M-1Y, EBA Q&A 2021_6017)
'
' PUBLIC SURFACE
'   None outside this VBA project. ReconcileBank is Public for the API
'
' STATE OWNERSHIP
'   None: no module state, no workbook access
'
' COMPATIBILITY
'   Excel VBA7, 64-bit; no references beyond the defaults
'
' AUTHOR
'   Daniele Penza
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const SHEET_BANK            As String = "BankCells"         'Sheet named in bank-cell diagnostics

    'Roles of a template column
    Private Const ROLE_AMOUNT           As Long = 1                     'Amount cell
    Private Const ROLE_FACTOR           As Long = 2                     'Applicable-factor cell
    Private Const ROLE_STABLE           As Long = 3                     'Stable-funding cell
    Private Const ROLE_IGNORED          As Long = 4                     'Standard-factor cell, not compared
    Private Const ROLE_SUMMARY          As Long = 5                     'C 84.00 cell


'
'------------------------------------------------------------------------------
'
'                                RECONCILIATION
'
'------------------------------------------------------------------------------
'

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
'------------------------------------------------------------------------------
'                         RECONCILE THE BANK'S FIGURES
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the four checks and the attribution described in the module header
'
' WHY THIS EXISTS
'   One entry point for the API turns the bank's cells into reconciliation
'   lines and the headline of the four levels and their effects
'
' INPUTS
'   Factors, FactorRows
'     The factor table of Params and its row count, read only
'
'   Summary, SummaryRows
'     The C 84.00 summary map of Params and its row count, read only
'
'   BankCells, BankRows
'     The bank's (template, row, column, value, source) rows and their
'     count, read only
'
'   AmountScale
'     Units of the reporting currency per bank amount (> 0)
'
'   AbsTolerance, RelTolerance, FactorTolerance
'     Tolerances, >= 0
'
'   HasBenchmark
'     True when the positions run is VALID; RowAmounts, RowStable and the
'     benchmark totals are then compared
'
'   RowAmounts, RowStable
'     The benchmark's amount and stable funding per factor-table row and
'     bucket
'
'   BenchmarkAsf, BenchmarkRsf
'     The benchmark's ASF and RSF totals
'
'   Lines, LineCount
'     Output: the reconciliation lines and their count
'
'   Headline
'     Output: the levels and effects
'
'   Diag, DiagCount
'     The diagnostic list, appended to here
'
' RETURNS
'   True when at least one bank figure was usable. Lines: a 2-D array
'   (1 To LineCount, 1 To RL_COLUMNS), sorted by check, template, row and
'   column. Headline (1 To 8, 1 To 3): ASF, RSF and NSFR of the four levels
'   (rows 1-4) and of the effects arithmetic, factor, amounts, total
'   (rows 5-8); Empty where a level is not available
'
' BEHAVIOR
'   Steps:
'     - reads the bank cells into the grid of rows and buckets, with a
'       diagnostic for a value that is not a number, a cell listed twice, a
'       factor outside 0-1 and the figures not compared; returns no lines
'       when no figure is usable
'     - processes the rows deepest first: arithmetic, factors,
'       re-performance, restatement
'     - checks C 84.00 against the bank's own rows
'     - compares amounts and stable funding with the benchmark when
'       HasBenchmark is True
'     - builds the headline and the attribution, then sorts the lines
'
' ERROR POLICY
'   No On Error. Unusable bank values are diagnostics; unexpected runtime
'   errors propagate
'
' DEPENDENCIES
'   NSFR_Core_Util: TextOf, NumberOf, IsBlankValue, FindRow1, FindRow2,
'   AddDiagnostic, IdDepth, HasBucket, ToleranceResult, ListToMatrix.
'   NSFR_Core_Engine: FactorValue, ComputeSummary.
'   ColumnRole, AmountColumn, FactorColumn, StableColumn, EffectiveFactor,
'   SummaryHasCells, RatioOf, DifferenceOf, AddLine, AddSummaryLine,
'   SortLines
'
' NOTES
'   A bank ratio (C 84.00 column 0040) is not scaled by AmountScale
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Core_Reconcile.ReconcileBank"

    Dim BankAmount()            As Double       'Bank amount per row and bucket, scaled
    Dim AmountSeen()            As Boolean      'Amount cell present in the bank file
    Dim BankFactor()            As Double       'Bank applicable factor per row and bucket
    Dim FactorSeen()            As Boolean      'Factor cell present
    Dim BankStable()            As Double       'Bank stable funding per row, scaled
    Dim StableSeen()            As Boolean      'Stable-funding cell present
    Dim Bank84()                As Double       'Bank C 84.00 per line and column
    Dim Seen84()                As Boolean      'C 84.00 cell present
    Dim Reperformed()           As Double       'Bank amounts x bank factors, per row
    Dim Restated()              As Double       'Bank amounts x regulatory factors, per row
    Dim DerivedAmt()            As Double       'Bank amount, or the children's where a parent has no cell
    Dim RowTotalAmt()           As Double       'Bank amount over all buckets, per row
    Dim RowTotalVal()           As Double       'Bank stable funding (reported, else re-performed)
    Dim Derived()               As Variant      'C 84.00 derived from the bank's rows
    Dim List()                  As Variant      'Reconciliation lines
    Dim Keys()                  As String       'Sort key per line
    Dim Usable                  As Long         'Usable bank figures
    Dim Ignored                 As Long         'Figures of rows or templates not compared
    Dim TemplateCode            As String       'Template of a bank cell
    Dim RowCode                 As String       'Row code of a bank cell
    Dim ColumnCode              As String       'Column code of a bank cell
    Dim Role                    As Long         'ROLE_* of the column
    Dim Bucket                  As Long         'Bucket of the column
    Dim Value                   As Double       'Bank value
    Dim Valid                   As Boolean      'Conversion outcome
    Dim TargetRow               As Long         'Row of the factor table or summary map
    Dim Kind                    As String       'Kind of a factor-table row
    Dim MaxDepth                As Long         'Deepest row ID
    Dim Depth                   As Long         'Depth being processed
    Dim Children                As Variant      'Row codes of the children
    Dim ChildRow                As Long         'Row of a child
    Dim Expected                As Double       'Expected figure of an arithmetic check
    Dim Amount                  As Double       'Bank amount of a bucket
    Dim Effective               As Double       'Bank factor as the template applies it
    Dim HasFactor               As Boolean      'Effective factor known
    Dim RegFactor               As Double       'Regulatory factor
    Dim RegPresent              As Boolean      'Regulatory factor defined
    Dim Total                   As Double       'Re-performed stable funding of a row
    Dim RestTotal               As Double       'Restated stable funding of a row
    Dim DerivedAsf              As Double       'C 84.00 ASF from the bank's rows
    Dim DerivedRsf              As Double       'C 84.00 RSF from the bank's rows
    Dim DerivedOk               As Boolean      'Derived ratio defined
    Dim Rsf80                   As Long         'Factor-table row of C 80.00 ID 1
    Dim Asf81                   As Long         'Factor-table row of C 81.00 ID 2
    Dim Levels(1 To 8, 1 To 3)  As Variant      'Headline
    Dim BucketsOf               As Long         'Amount buckets of a template
    Dim EngineStable            As Double       'Benchmark stable funding of a row
    Dim BankRowStable           As Double       'Bank stable funding of a row, 0 if absent
    Dim Measure                 As String       'Summary measure
    Dim ValueColumn             As Long         'C 84.00 column of a summary value
    Dim RowIndex                As Long         'Bank-cell row, factor-table row
    Dim InnerIndex              As Long         'Child index, headline column
    Dim BucketIndex             As Long         'Bucket being processed
    Dim SummaryIndex            As Long         'Summary-map row

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Size the grids of rows and buckets
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
    'Size the C 84.00 grid, at least one line
        ReDim Bank84(1 To IIf(SummaryRows > 0, SummaryRows, 1), 1 To 4)
        ReDim Seen84(1 To IIf(SummaryRows > 0, SummaryRows, 1), 1 To 4)
    'Start with no line
        LineCount = 0

'------------------------------------------------------------------------------
' 1. READ THE BANK CELLS INTO THE GRID
'------------------------------------------------------------------------------
    'Place each bank cell by its role, or report why it is not used
        For RowIndex = 1 To BankRows
            TemplateCode = TextOf(BankCells(RowIndex, BC_TEMPLATE))
            RowCode = TextOf(BankCells(RowIndex, BC_ROW))
            ColumnCode = TextOf(BankCells(RowIndex, BC_COLUMN))
            If Len(TemplateCode) = 0 And Len(RowCode) = 0 And Len(ColumnCode) = 0 Then
                'blank line: skipped
            Else
                Role = ColumnRole(TemplateCode, ColumnCode, Bucket)
                If Role = 0 Then
                    Ignored = Ignored + 1
                ElseIf Role = ROLE_IGNORED Then
                    'standard factors of the template are not compared
                Else
                    Value = NumberOf(BankCells(RowIndex, BC_VALUE), Valid)
                    If Role = ROLE_SUMMARY Then
                        TargetRow = FindRow1(Summary, SummaryRows, SC_ROW, RowCode)
                    Else
                        TargetRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ROW, RowCode)
                    End If
                    If TargetRow = 0 Then
                        Ignored = Ignored + 1
                    ElseIf Not Valid Then
                        If Not IsBlankValue(BankCells(RowIndex, BC_VALUE)) Then
                            AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_BANK, RowIndex, _
                                          TemplateCode & " r" & RowCode & " c" & ColumnCode, _
                                          "Bank value is not a number: the cell is not compared."
                        End If
                    ElseIf Role = ROLE_SUMMARY Then
                        If Seen84(TargetRow, Bucket) Then
                            AddDiagnostic Diag, DiagCount, SEV_WARNING, SHEET_BANK, RowIndex, _
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
                            AddDiagnostic Diag, DiagCount, SEV_WARNING, SHEET_BANK, RowIndex, _
                                          TemplateCode & " r" & RowCode & " c" & ColumnCode, _
                                          "Cell listed twice: the first value is used."
                        Else
                            BankAmount(TargetRow, Bucket) = Value * AmountScale
                            AmountSeen(TargetRow, Bucket) = True
                            Usable = Usable + 1
                        End If
                    ElseIf Role = ROLE_FACTOR Then
                        If FactorSeen(TargetRow, Bucket) Then
                            AddDiagnostic Diag, DiagCount, SEV_WARNING, SHEET_BANK, RowIndex, _
                                          TemplateCode & " r" & RowCode & " c" & ColumnCode, _
                                          "Cell listed twice: the first value is used."
                        ElseIf Value < 0 Or Value > 1 Then
                            AddDiagnostic Diag, DiagCount, SEV_ERROR, SHEET_BANK, RowIndex, _
                                          TemplateCode & " r" & RowCode & " c" & ColumnCode, _
                                          "Bank factor outside 0-1: treated as missing."
                        Else
                            BankFactor(TargetRow, Bucket) = Value
                            FactorSeen(TargetRow, Bucket) = True
                            Usable = Usable + 1
                        End If
                    ElseIf Role = ROLE_STABLE Then
                        If StableSeen(TargetRow) Then
                            AddDiagnostic Diag, DiagCount, SEV_WARNING, SHEET_BANK, RowIndex, _
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
        Next RowIndex
    'Report the figures that are not compared
        If Ignored > 0 Then
            AddDiagnostic Diag, DiagCount, SEV_INFO, SHEET_BANK, 0, "Bank cells", _
                          Ignored & " bank figure(s) of other templates, rows or columns are not compared " _
                          & "(simplified NSFR C 82.00 / C 83.00 is out of scope)."
        End If
    'Return no line and an empty headline when no figure is usable
        ReconcileBank = (Usable > 0)
        If Usable = 0 Then
            Lines = ListToMatrix(List, 0, RL_COLUMNS)
            Headline = Levels
            Exit Function
        End If

'------------------------------------------------------------------------------
' 2. ROWS, DEEPEST FIRST: ARITHMETIC, FACTORS, RE-PERFORMANCE, RESTATEMENT
'------------------------------------------------------------------------------
    'Find the deepest row ID
        For RowIndex = 1 To FactorRows
            If IdDepth(TextOf(Factors(RowIndex, FC_ID))) > MaxDepth Then _
                MaxDepth = IdDepth(TextOf(Factors(RowIndex, FC_ID)))
        Next RowIndex
    'Process the rows of each depth, children before their parents
        For Depth = MaxDepth To 0 Step -1
            For RowIndex = 1 To FactorRows
                If IdDepth(TextOf(Factors(RowIndex, FC_ID))) = Depth Then
                    Kind = TextOf(Factors(RowIndex, FC_KIND))
                    TemplateCode = TextOf(Factors(RowIndex, FC_TEMPLATE))
                    If Kind = KIND_AGG Then
                        Children = Split(TextOf(Factors(RowIndex, FC_CHILDREN)), ",")
                        Total = 0
                        RestTotal = 0
                        Expected = 0
                        For InnerIndex = LBound(Children) To UBound(Children)
                            ChildRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ROW, _
                                                Trim$(Children(InnerIndex)))
                            If ChildRow > 0 Then
                                Total = Total + Reperformed(ChildRow)
                                RestTotal = RestTotal + Restated(ChildRow)
                                If StableSeen(ChildRow) Then
                                    Expected = Expected + BankStable(ChildRow)
                                Else
                                    Expected = Expected + Reperformed(ChildRow)
                                End If
                            End If
                        Next InnerIndex
                        Reperformed(RowIndex) = Total
                        Restated(RowIndex) = RestTotal
                        For BucketIndex = 1 To BUCKET_COUNT
                            'A parent without a cell in a bucket carries its children's
                            'amounts upwards (C 80.00 row 1.7 has no HQLA cell, 1.7.3 has).
                            Amount = 0
                            For InnerIndex = LBound(Children) To UBound(Children)
                                ChildRow = FindRow2(Factors, FactorRows, FC_TEMPLATE, TemplateCode, FC_ROW, _
                                                    Trim$(Children(InnerIndex)))
                                If ChildRow > 0 Then Amount = Amount + DerivedAmt(ChildRow, BucketIndex)
                            Next InnerIndex
                            If Not AmountSeen(RowIndex, BucketIndex) Then
                                DerivedAmt(RowIndex, BucketIndex) = Amount
                            Else
                                DerivedAmt(RowIndex, BucketIndex) = BankAmount(RowIndex, BucketIndex)
                                AddLine List, LineCount, Keys, CHECK_ARITHMETIC, Factors, RowIndex, _
                                        AmountColumn(TemplateCode, BucketIndex), _
                                        BankAmount(RowIndex, BucketIndex), Amount, _
                                        ToleranceResult(Amount - BankAmount(RowIndex, BucketIndex), _
                                                        BankAmount(RowIndex, BucketIndex), _
                                                        AbsTolerance, RelTolerance), Empty
                            End If
                        Next BucketIndex
                    Else
                        Total = 0
                        RestTotal = 0
                        For BucketIndex = 1 To BUCKET_COUNT
                            DerivedAmt(RowIndex, BucketIndex) = BankAmount(RowIndex, BucketIndex)
                            Amount = BankAmount(RowIndex, BucketIndex)
                            If Amount <> 0 Then
                                Effective = EffectiveFactor(Factors, RowIndex, BucketIndex, BankFactor, _
                                                            FactorSeen, HasFactor)
                                RegFactor = FactorValue(Factors, RowIndex, FC_REG, BucketIndex, RegPresent)
                                If Kind = KIND_MEMO Then
                                    If Not HasFactor Then
                                        'A memo row's factor feeds C 84.00 (2.3 and 2.4).
                                        Effective = 0
                                        AddLine List, LineCount, Keys, CHECK_FACTOR, Factors, RowIndex, _
                                                FactorColumn(TemplateCode, BucketIndex), Empty, Empty, _
                                                RESULT_MISSING, Empty
                                    End If
                                    RestTotal = RestTotal + Amount * Effective
                                ElseIf Not HasFactor Then
                                    Effective = 0
                                    If RegPresent Then
                                        AddLine List, LineCount, Keys, CHECK_FACTOR, Factors, RowIndex, _
                                                FactorColumn(TemplateCode, BucketIndex), Empty, RegFactor, _
                                                RESULT_MISSING, Amount * RegFactor
                                        RestTotal = RestTotal + Amount * RegFactor
                                    Else
                                        AddLine List, LineCount, Keys, CHECK_FACTOR, Factors, RowIndex, _
                                                FactorColumn(TemplateCode, BucketIndex), Empty, Empty, _
                                                RESULT_MISSING, Empty
                                    End If
                                ElseIf Not RegPresent Then
                                    AddLine List, LineCount, Keys, CHECK_FACTOR, Factors, RowIndex, _
                                            FactorColumn(TemplateCode, BucketIndex), Effective, Empty, _
                                            RESULT_INFO, Empty
                                    RestTotal = RestTotal + Amount * Effective
                                Else
                                    If HasBucket(TextOf(Factors(RowIndex, FC_FACTOR_CELLS)), BucketIndex) Then
                                        ColumnCode = FactorColumn(TemplateCode, BucketIndex)
                                    Else
                                        ColumnCode = AmountColumn(TemplateCode, BucketIndex)
                                    End If
                                    AddLine List, LineCount, Keys, CHECK_FACTOR, Factors, RowIndex, ColumnCode, _
                                            Effective, RegFactor, _
                                            IIf(Abs(Effective - RegFactor) <= FactorTolerance, _
                                                RESULT_MATCH, RESULT_DIFF), _
                                            Amount * (RegFactor - Effective)
                                    RestTotal = RestTotal + Amount * RegFactor
                                End If
                                Total = Total + Amount * Effective
                            End If
                        Next BucketIndex
                        Reperformed(RowIndex) = Total
                        Restated(RowIndex) = RestTotal
                        Expected = Total
                    End If
                    If StableSeen(RowIndex) Then
                        AddLine List, LineCount, Keys, CHECK_ARITHMETIC, Factors, RowIndex, _
                                StableColumn(TemplateCode), BankStable(RowIndex), Expected, _
                                ToleranceResult(Expected - BankStable(RowIndex), BankStable(RowIndex), _
                                                AbsTolerance, RelTolerance), _
                                Empty
                    End If
                End If
            Next RowIndex
        Next Depth

'------------------------------------------------------------------------------
' 3. C 84.00 AGAINST THE BANK'S OWN ROWS
'------------------------------------------------------------------------------
    'Total each row's amounts and take its stable funding, reported else re-performed
        For RowIndex = 1 To FactorRows
            RowTotalAmt(RowIndex) = 0
            For BucketIndex = 1 To BUCKET_COUNT
                RowTotalAmt(RowIndex) = RowTotalAmt(RowIndex) + DerivedAmt(RowIndex, BucketIndex)
            Next BucketIndex
            If StableSeen(RowIndex) Then RowTotalVal(RowIndex) = BankStable(RowIndex) _
                Else RowTotalVal(RowIndex) = Reperformed(RowIndex)
        Next RowIndex
    'Derive C 84.00 from the bank's rows
        DerivedOk = ComputeSummary(Factors, FactorRows, Summary, SummaryRows, RowTotalAmt, RowTotalVal, _
                                   Derived, DerivedAsf, DerivedRsf)
    'Compare each reported C 84.00 cell with the derived one
        For SummaryIndex = 1 To SummaryRows
            Measure = TextOf(Summary(SummaryIndex, SC_MEASURE))
            If Measure = "NSFR" Then
                If Seen84(SummaryIndex, 4) And DerivedOk Then
                    AddSummaryLine List, LineCount, Keys, CHECK_ARITHMETIC, Summary, SummaryIndex, "0040", _
                                   Bank84(SummaryIndex, 4), Derived(SummaryIndex, 4), _
                                   IIf(Abs(Derived(SummaryIndex, 4) - Bank84(SummaryIndex, 4)) <= FactorTolerance, _
                                       RESULT_MATCH, RESULT_DIFF)
                End If
            Else
                If Seen84(SummaryIndex, 1) Then
                    AddSummaryLine List, LineCount, Keys, CHECK_ARITHMETIC, Summary, SummaryIndex, "0010", _
                                   Bank84(SummaryIndex, 1), Derived(SummaryIndex, 1), _
                                   ToleranceResult(Derived(SummaryIndex, 1) - Bank84(SummaryIndex, 1), _
                                                   Bank84(SummaryIndex, 1), AbsTolerance, RelTolerance)
                End If
                If Measure = "RSF" Then ValueColumn = 2 Else ValueColumn = 3
                If Seen84(SummaryIndex, ValueColumn) Then
                    AddSummaryLine List, LineCount, Keys, CHECK_ARITHMETIC, Summary, SummaryIndex, _
                                   "00" & ValueColumn & "0", _
                                   Bank84(SummaryIndex, ValueColumn), Derived(SummaryIndex, ValueColumn), _
                                   ToleranceResult(Derived(SummaryIndex, ValueColumn) - _
                                                   Bank84(SummaryIndex, ValueColumn), _
                                                   Bank84(SummaryIndex, ValueColumn), AbsTolerance, RelTolerance)
                End If
            End If
        Next SummaryIndex

'------------------------------------------------------------------------------
' 4. BENCHMARK COMPARISON
'------------------------------------------------------------------------------
    'Compare each row's amounts and stable funding with the positions run
        If HasBenchmark Then
            For RowIndex = 1 To FactorRows
                TemplateCode = TextOf(Factors(RowIndex, FC_TEMPLATE))
                If TemplateCode = TEMPLATE_RSF Then BucketsOf = 4 Else BucketsOf = 3
                For BucketIndex = 1 To BucketsOf
                    If HasBucket(TextOf(Factors(RowIndex, FC_AMOUNT_CELLS)), BucketIndex) Then
                        If BankAmount(RowIndex, BucketIndex) <> 0 Or RowAmounts(RowIndex, BucketIndex) <> 0 Then
                            AddLine List, LineCount, Keys, CHECK_AMOUNT, Factors, RowIndex, _
                                    AmountColumn(TemplateCode, BucketIndex), _
                                    BankAmount(RowIndex, BucketIndex), RowAmounts(RowIndex, BucketIndex), _
                                    ToleranceResult(RowAmounts(RowIndex, BucketIndex) - _
                                                    BankAmount(RowIndex, BucketIndex), _
                                                    BankAmount(RowIndex, BucketIndex), _
                                                    AbsTolerance, RelTolerance), Empty
                        End If
                    End If
                Next BucketIndex
                EngineStable = 0
                For BucketIndex = 1 To BUCKET_COUNT
                    EngineStable = EngineStable + RowStable(RowIndex, BucketIndex)
                Next BucketIndex
                If StableSeen(RowIndex) Then BankRowStable = BankStable(RowIndex) Else BankRowStable = 0
                If BankRowStable <> 0 Or EngineStable <> 0 Then
                    AddLine List, LineCount, Keys, CHECK_STABLE, Factors, RowIndex, StableColumn(TemplateCode), _
                            BankRowStable, EngineStable, _
                            ToleranceResult(EngineStable - BankRowStable, BankRowStable, AbsTolerance, RelTolerance), _
                            Empty
                End If
            Next RowIndex
        End If

'------------------------------------------------------------------------------
' 5. HEADLINE AND ATTRIBUTION
'------------------------------------------------------------------------------
    'Find the total rows of C 80.00 and C 81.00
        Rsf80 = FindRow2(Factors, FactorRows, FC_TEMPLATE, TEMPLATE_RSF, FC_ID, "1")
        Asf81 = FindRow2(Factors, FactorRows, FC_TEMPLATE, TEMPLATE_ASF, FC_ID, "2")
    'Take the reported level from the bank's C 84.00
        For SummaryIndex = 1 To SummaryRows
            If TextOf(Summary(SummaryIndex, SC_ID)) = "1" And Seen84(SummaryIndex, 2) Then _
                Levels(HL_REPORTED, HL_RSF) = Bank84(SummaryIndex, 2)
            If TextOf(Summary(SummaryIndex, SC_ID)) = "2" And Seen84(SummaryIndex, 3) Then _
                Levels(HL_REPORTED, HL_ASF) = Bank84(SummaryIndex, 3)
            If TextOf(Summary(SummaryIndex, SC_MEASURE)) = "NSFR" And Seen84(SummaryIndex, 4) Then _
                Levels(HL_REPORTED, HL_NSFR) = Bank84(SummaryIndex, 4)
        Next SummaryIndex
    'Fall back on the C 80.00 / C 81.00 totals
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
    'Set the re-performed and restated levels
        If Rsf80 > 0 And Asf81 > 0 Then
            Levels(HL_REPERFORMED, HL_ASF) = Reperformed(Asf81)
            Levels(HL_REPERFORMED, HL_RSF) = Reperformed(Rsf80)
            Levels(HL_REPERFORMED, HL_NSFR) = RatioOf(Reperformed(Asf81), Reperformed(Rsf80))
            Levels(HL_RESTATED, HL_ASF) = Restated(Asf81)
            Levels(HL_RESTATED, HL_RSF) = Restated(Rsf80)
            Levels(HL_RESTATED, HL_NSFR) = RatioOf(Restated(Asf81), Restated(Rsf80))
        End If
    'Set the benchmark level
        If HasBenchmark Then
            Levels(HL_BENCHMARK, HL_ASF) = BenchmarkAsf
            Levels(HL_BENCHMARK, HL_RSF) = BenchmarkRsf
            Levels(HL_BENCHMARK, HL_NSFR) = RatioOf(BenchmarkAsf, BenchmarkRsf)
        End If
    'Set the effects: arithmetic, factor, amounts, total
        For InnerIndex = 1 To 3
            Levels(5, InnerIndex) = DifferenceOf(Levels(HL_REPERFORMED, InnerIndex), Levels(HL_REPORTED, InnerIndex))
            Levels(6, InnerIndex) = DifferenceOf(Levels(HL_RESTATED, InnerIndex), Levels(HL_REPERFORMED, InnerIndex))
            Levels(7, InnerIndex) = DifferenceOf(Levels(HL_BENCHMARK, InnerIndex), Levels(HL_RESTATED, InnerIndex))
            Levels(8, InnerIndex) = DifferenceOf(Levels(HL_BENCHMARK, InnerIndex), Levels(HL_REPORTED, InnerIndex))
        Next InnerIndex
        Headline = Levels

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the lines sorted by check, template, row and column
        SortLines List, Keys, LineCount
        Lines = ListToMatrix(List, LineCount, RL_COLUMNS)

End Function

'
'------------------------------------------------------------------------------
'
'                                   COLUMNS
'
'------------------------------------------------------------------------------
'

Private Function ColumnRole( _
    ByVal TemplateCode As String, _
    ByVal ColumnCode As String, _
    ByRef Bucket As Long) _
    As Long
'
'------------------------------------------------------------------------------
'                          CLASSIFY A TEMPLATE COLUMN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the role and bucket of a column code of the reporting templates
'
' WHY THIS EXISTS
'   The bank's cells are compared by role: amount, applicable factor,
'   stable funding or C 84.00 value
'
' INPUTS
'   TemplateCode
'     The template of the cell
'
'   ColumnCode
'     The four-digit column code
'
'   Bucket
'     Output: the bucket of the column, 0 when it has none
'
' RETURNS
'   ROLE_* constant, or 0 for a template or column not compared.
'   C 80.00: 0010-0040 amounts, 0050-0080 standard factors, 0090-0120
'   applicable factors, 0130 RSF. C 81.00: 0010-0030 amounts, 0040-0060
'   standard, 0070-0090 applicable, 0100 ASF. C 84.00: 0010-0040,
'   Bucket = column number
'
' BEHAVIOR
'   Rejects a code that is not four digits or not a multiple of 10, then
'   maps the column number (code / 10) by template
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   NSFR_Core_Util: IsAllDigits
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
    Const PROC_NAME             As String = "NSFR_Core_Reconcile.ColumnRole"

    Dim Number                  As Long         'Column code / 10

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Start without a bucket
        Bucket = 0
    'Reject a code that is not four digits or not a multiple of 10
        If Not IsAllDigits(ColumnCode) Or Len(ColumnCode) <> 4 Then Exit Function
        Number = CLng(ColumnCode) \ 10
        If CLng(ColumnCode) Mod 10 <> 0 Then Exit Function

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Map the column number by template
        Select Case TemplateCode
            Case TEMPLATE_RSF
                Select Case Number
                    Case 1 To 4
                        ColumnRole = ROLE_AMOUNT
                        Bucket = Number
                    Case 5 To 8
                        ColumnRole = ROLE_IGNORED
                    Case 9 To 12
                        ColumnRole = ROLE_FACTOR
                        Bucket = Number - 8
                    Case 13
                        ColumnRole = ROLE_STABLE
                    Case Else
                        'No role for other columns
                End Select
            Case TEMPLATE_ASF
                Select Case Number
                    Case 1 To 3
                        ColumnRole = ROLE_AMOUNT
                        Bucket = Number
                    Case 4 To 6
                        ColumnRole = ROLE_IGNORED
                    Case 7 To 9
                        ColumnRole = ROLE_FACTOR
                        Bucket = Number - 6
                    Case 10
                        ColumnRole = ROLE_STABLE
                    Case Else
                        'No role for other columns
                End Select
            Case TEMPLATE_SUMMARY
                If Number >= 1 And Number <= 4 Then
                    ColumnRole = ROLE_SUMMARY
                    Bucket = Number
                End If
            Case Else
                'No role for other templates
        End Select

End Function

Private Function AmountColumn( _
    ByVal TemplateCode As String, _
    ByVal Bucket As Long) _
    As String
'
'------------------------------------------------------------------------------
'                          AMOUNT COLUMN OF A BUCKET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the column code of the amount of a bucket (0010-0040)
'
' NOTES
'   TemplateCode is not used: both templates number their amounts alike
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Format ten times the bucket on four digits
        AmountColumn = Format$(10 * Bucket, "0000")

End Function

Private Function FactorColumn( _
    ByVal TemplateCode As String, _
    ByVal Bucket As Long) _
    As String
'
'------------------------------------------------------------------------------
'                     APPLICABLE-FACTOR COLUMN OF A BUCKET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the column code of the applicable factor of a bucket: C 80.00
'   0090-0120, C 81.00 0070-0090
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Offset the bucket by the template's first applicable-factor column
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
'------------------------------------------------------------------------------
'                     STABLE-FUNDING COLUMN OF A TEMPLATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the column code of the stable funding: 0130 or 0100
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Pick the C 80.00 or C 81.00 column
        If TemplateCode = TEMPLATE_RSF Then StableColumn = "0130" Else StableColumn = "0100"

End Function

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function SummaryHasCells( _
    ByRef Seen84() As Boolean, _
    ByVal SummaryRows As Long) _
    As Boolean
'
'------------------------------------------------------------------------------
'                    TELL WHETHER A C 84.00 CELL WAS USABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when at least one C 84.00 cell was usable
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Core_Reconcile.SummaryHasCells"

    Dim SummaryIndex            As Long         'Summary-map row
    Dim ColumnIndex             As Long         'C 84.00 column

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Stop at the first usable cell
        For SummaryIndex = 1 To SummaryRows
            For ColumnIndex = 1 To 4
                If Seen84(SummaryIndex, ColumnIndex) Then
                    SummaryHasCells = True
                    Exit Function
                End If
            Next ColumnIndex
        Next SummaryIndex

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
'------------------------------------------------------------------------------
'                    BANK FACTOR AS THE TEMPLATE APPLIES IT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the bank's factor as the template applies it: the entered
'   applicable factor where the template has a factor cell, 0 where it has
'   none
'
' RETURNS
'   The factor; Known = False when the template has a factor cell for the
'   bucket and the bank left it blank
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Use the entered factor where the template has a factor cell, else 0
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
'------------------------------------------------------------------------------
'                         DIVIDE TWO OPTIONAL FIGURES
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Numerator / Denominator, Empty when either is Empty or the
'   denominator is not positive
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Leave the result Empty when the ratio is not defined
        If IsEmpty(Numerator) Or IsEmpty(Denominator) Then Exit Function
        If CDbl(Denominator) <= 0 Then Exit Function

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Divide
        RatioOf = CDbl(Numerator) / CDbl(Denominator)

End Function

Private Function DifferenceOf( _
    ByVal Minuend As Variant, _
    ByVal Subtrahend As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'                        SUBTRACT TWO OPTIONAL FIGURES
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Minuend - Subtrahend, Empty when either is Empty
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Leave the result Empty when a figure is missing
        If IsEmpty(Minuend) Or IsEmpty(Subtrahend) Then Exit Function

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Subtract
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
'------------------------------------------------------------------------------
'                   APPEND A LINE OF A C 80.00 / C 81.00 ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one reconciliation line of a C 80.00 / C 81.00 row
'
' DEPENDENCIES
'   AppendLine; NSFR_Core_Util: TextOf
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPEND LINE
'------------------------------------------------------------------------------
    'Describe the line by the template, row, ID and item of the factor-table row
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
'------------------------------------------------------------------------------
'                       APPEND A LINE OF A C 84.00 LINE
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one reconciliation line of a C 84.00 line
'
' DEPENDENCIES
'   AppendLine; NSFR_Core_Util: TextOf
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPEND LINE
'------------------------------------------------------------------------------
    'Describe the line by the row, ID and item of the summary map, without impact
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
'------------------------------------------------------------------------------
'                    BUILD AND APPEND A RECONCILIATION LINE
'------------------------------------------------------------------------------
' PURPOSE
'   Builds a line with its difference and relative difference, appends it
'   and its sort key
'
' WHY THIS EXISTS
'   AddLine and AddSummaryLine share one layout of the line and of its key
'
' INPUTS
'   List, LineCount
'     The lines so far and their count, appended to here
'
'   Keys
'     The sort key per line, resized to follow List
'
'   CheckName
'     The check: ARITHMETIC, FACTOR, AMOUNT or STABLE FUNDING
'
'   TemplateCode, RowCode, RowId, Item, ColumnCode
'     The cell the line describes
'
'   BankValue, OtherValue
'     The bank's figure and the expected or benchmark figure, Empty when
'     absent
'
'   Outcome
'     The result: MATCH, DIFF, MISSING or INFO
'
'   Impact
'     The stable-funding impact of a factor difference, else Empty
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Sets the difference (other - bank) when both figures are present, and
'   the relative difference when the bank's figure is not 0. The key is
'   the check rank (ARITHMETIC 1, FACTOR 2, AMOUNT 3, other 4), template,
'   row and column, separated by "|"
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   NSFR_Core_Util: NewRow, ListAdd
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
    Const PROC_NAME             As String = "NSFR_Core_Reconcile.AppendLine"

    Dim RowItem                 As Variant      'Line
    Dim Rank                    As String       'Check order

'------------------------------------------------------------------------------
' BUILD LINE
'------------------------------------------------------------------------------
    'Fill the line
        RowItem = NewRow(RL_COLUMNS)
        RowItem(RL_CHECK) = CheckName
        RowItem(RL_TEMPLATE) = TemplateCode
        RowItem(RL_ROW) = RowCode
        RowItem(RL_ID) = RowId
        RowItem(RL_ITEM) = Item
        RowItem(RL_COLUMN) = ColumnCode
        RowItem(RL_BANK) = BankValue
        RowItem(RL_OTHER) = OtherValue
    'Set the difference and relative difference when both figures are present
        If Not IsEmpty(BankValue) And Not IsEmpty(OtherValue) Then
            RowItem(RL_DIFFERENCE) = CDbl(OtherValue) - CDbl(BankValue)
            If CDbl(BankValue) <> 0 Then RowItem(RL_RELATIVE) = RowItem(RL_DIFFERENCE) / Abs(CDbl(BankValue))
        End If
        RowItem(RL_RESULT) = Outcome
        RowItem(RL_IMPACT) = Impact

'------------------------------------------------------------------------------
' APPEND LINE
'------------------------------------------------------------------------------
    'Append the line
        ListAdd List, LineCount, RowItem
    'Rank the check
        Select Case CheckName
            Case CHECK_ARITHMETIC
                Rank = "1"
            Case CHECK_FACTOR
                Rank = "2"
            Case CHECK_AMOUNT
                Rank = "3"
            Case Else
                Rank = "4"
        End Select
    'Keep the keys as long as the list
        If LineCount = 1 Then
            ReDim Keys(1 To UBound(List))
        ElseIf UBound(Keys) < UBound(List) Then
            ReDim Preserve Keys(1 To UBound(List))
        End If
    'Store the sort key
        Keys(LineCount) = Rank & "|" & TemplateCode & "|" & RowCode & "|" & ColumnCode

End Sub

Private Sub SortLines( _
    ByRef List() As Variant, _
    ByRef Keys() As String, _
    ByVal LineCount As Long)
'
'------------------------------------------------------------------------------
'                         SORT THE LINES BY THEIR KEYS
'------------------------------------------------------------------------------
' PURPOSE
'   Stable insertion sort of the lines by their keys (check, template, row,
'   column)
'
' WHY THIS EXISTS
'   Lines with equal keys keep the order in which the checks added them
'
' INPUTS
'   List
'     The lines, sorted in place
'
'   Keys
'     The sort key per line, sorted with the lines
'
'   LineCount
'     The number of lines to sort
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Inserts each line, from the second, after the last earlier line whose
'   key is not greater (binary comparison)
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
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
    Const PROC_NAME             As String = "NSFR_Core_Reconcile.SortLines"

    Dim KeyValue                As String       'Key being inserted
    Dim Item                    As Variant      'Line being inserted
    Dim InsertIndex             As Long         'Line being inserted, from the second
    Dim CompareIndex            As Long         'Line compared with it, moving down

'------------------------------------------------------------------------------
' SORT
'------------------------------------------------------------------------------
    'Insert each line after the last line whose key is not greater
        For InsertIndex = 2 To LineCount
            KeyValue = Keys(InsertIndex)
            Item = List(InsertIndex)
            CompareIndex = InsertIndex - 1
            Do While CompareIndex >= 1
                If StrComp(Keys(CompareIndex), KeyValue, vbBinaryCompare) <= 0 Then Exit Do
                Keys(CompareIndex + 1) = Keys(CompareIndex)
                List(CompareIndex + 1) = List(CompareIndex)
                CompareIndex = CompareIndex - 1
            Loop
            Keys(CompareIndex + 1) = KeyValue
            List(CompareIndex + 1) = Item
        Next InsertIndex

End Sub
