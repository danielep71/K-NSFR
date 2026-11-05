Attribute VB_Name = "Test_NSFR_Harness"
'==============================================================================
' MODULE        Test_NSFR_Harness
' PURPOSE       Neutral test harness: records the outcome of each declared
'               case and builds a deterministic text summary. It implements
'               no NSFR calculation and asserts nothing about one.
' PUBLIC        NsfrTestRun, BeginRun, CheckExact, CheckText, CheckTolerance,
'               CheckDiagnostic, SkipCase, RecordCleanup, Summary, RunPassed
' DEPENDENCIES  VBA only; no Excel object model, file, clock or random number.
' STATE         None at module level: the caller owns each NsfrTestRun.
' ERRORS        Validates instead of trapping: an invalid case ID, number or
'               tolerance becomes a failure or a problem in the summary.
' COMPATIBILITY Windows 11, Microsoft 365 desktop Excel, VBA7, 64-bit (#14).
' RULES         tests/README.md; numbers follow docs/validation/VALIDATION.md.
' UPDATED       2026-11-05
' AUTHOR        danielep71
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
Option Explicit
Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
Public Const OUTCOME_PASS As String = "PASS"
Public Const OUTCOME_FAIL As String = "FAIL"
Public Const OUTCOME_SKIPPED As String = "SKIPPED"
Public Const OUTCOME_MISSING As String = "MISSING"
Public Const MODE_ABSOLUTE As String = "absolute"
Public Const MODE_RELATIVE As String = "relative"
Private Const MAX_DIGITS As Long = 28           'Significant digits a Decimal holds exactly

'One test run. The caller declares it, passes it ByRef to every procedure
'of this module and never changes its members directly.
Public Type NsfrTestRun
    SuiteName       As String                   'Name on the first summary line
    DeclaredIds()   As String                   'Declared case IDs, in declaration order
    DeclaredCount   As Long                     'Number of declared case IDs
    CaseIds()       As String                   'Recorded case IDs, in recording order
    Outcomes()      As String                   'PASS, FAIL or SKIPPED per recorded case
    Details()       As String                   'Failure or skip detail per recorded case
    CaseCount       As Long                     'Number of recorded cases
    CleanupLines()  As String                   'One "ID detail" text per cleanup failure
    CleanupCount    As Long                     'Number of cleanup failures
    Problems()      As String                   'TestRun-level problems
    ProblemCount    As Long                     'Number of run-level problems
    UndeclaredCount As Long                     'Outcomes recorded for undeclared IDs
End Type


'------------------------------------------------------------------------------
'                              RUN LIFECYCLE
'------------------------------------------------------------------------------

Public Sub BeginRun( _
    ByRef TestRun As NsfrTestRun, _
    ByVal SuiteName As String, _
    ByVal DeclaredIdList As String)
'PURPOSE        Start a run and declare every case it must execute.
'INPUTS         TestRun: the run to reset (output). SuiteName: printed name.
'               DeclaredIdList: comma-separated case IDs, each of upper-case
'               letters, digits and hyphens, for example "FX-0001,FX-0002".
'ERROR POLICY   An invalid or repeated ID, or an empty list, is recorded as
'               a run-level problem; nothing is raised.
'UPDATED        2026-11-05
    Dim Fresh       As NsfrTestRun              'An empty run to reset with
    Dim Parts()     As String                   'The split ID list
    Dim CaseId      As String                   'One declared ID
    Dim i           As Long                     'Index into Parts

        TestRun = Fresh
        TestRun.SuiteName = SuiteName
        If Len(DeclaredIdList) > 0 Then
            Parts = Split(DeclaredIdList, ",")
            For i = LBound(Parts) To UBound(Parts)
                CaseId = Parts(i)
                If Not IsValidId(CaseId) Then
                    AddProblem TestRun, "invalid case ID " & Quoted(CaseId)
                ElseIf IsDeclared(TestRun, CaseId) Then
                    AddProblem TestRun, "case " & CaseId & " declared twice"
                Else
                    TestRun.DeclaredCount = TestRun.DeclaredCount + 1
                    ReDim Preserve TestRun.DeclaredIds(1 To TestRun.DeclaredCount)
                    TestRun.DeclaredIds(TestRun.DeclaredCount) = CaseId
                End If
            Next i
        End If
        If TestRun.DeclaredCount = 0 Then AddProblem TestRun, "no case declared"
End Sub


'------------------------------------------------------------------------------
'                                 OUTCOMES
'------------------------------------------------------------------------------

Public Sub CheckExact( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String, _
    ByVal Expected As String, _
    ByVal Actual As String)
'PURPOSE        Record PASS when Actual equals Expected character for
'               character (binary comparison), otherwise FAIL.
'INPUTS         Expected and Actual as text, numbers at their reported scale.
'UPDATED        2026-11-05

        If StrComp(Expected, Actual, vbBinaryCompare) = 0 Then
            RecordOutcome TestRun, CaseId, OUTCOME_PASS, vbNullString
        Else
            RecordOutcome TestRun, CaseId, OUTCOME_FAIL, _
                          "expected " & Quoted(Expected) & ", got " & Quoted(Actual)
        End If
End Sub


Public Sub CheckText( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String, _
    ByVal Expected As String, _
    ByVal Actual As String)
'PURPOSE        Record PASS when two multi-line texts (CR LF line ends) are
'               equal, otherwise FAIL naming the first line that differs.
'UPDATED        2026-11-05
    Dim ExpectedLines() As String               'Expected split into lines
    Dim ActualLines()   As String               'Actual split into lines
    Dim ExpectedLine    As String               'Expected line at LineNumber, or empty
    Dim ActualLine      As String               'Actual line at LineNumber, or empty
    Dim LineNumber      As Long                 'First line that differs, from 1

        If StrComp(Expected, Actual, vbBinaryCompare) = 0 Then
            RecordOutcome TestRun, CaseId, OUTCOME_PASS, vbNullString
            Exit Sub
        End If
        ExpectedLines = Split(Expected, vbCrLf)
        ActualLines = Split(Actual, vbCrLf)
        LineNumber = 1
        Do
            ExpectedLine = LineAt(ExpectedLines, LineNumber)
            ActualLine = LineAt(ActualLines, LineNumber)
            If StrComp(ExpectedLine, ActualLine, vbBinaryCompare) <> 0 Then Exit Do
            If LineNumber > UBound(ExpectedLines) + 1 And LineNumber > UBound(ActualLines) + 1 Then
                Exit Do
            End If
            LineNumber = LineNumber + 1
        Loop
        RecordOutcome TestRun, CaseId, OUTCOME_FAIL, _
                      "line " & LineNumber & ": expected " & Quoted(ExpectedLine) & _
                      ", got " & Quoted(ActualLine)
End Sub


Public Sub CheckTolerance( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String, _
    ByVal Expected As String, _
    ByVal Actual As String, _
    ByVal Mode As String, _
    ByVal Tolerance As String, _
    ByVal Floor As String)
'PURPOSE        Record PASS when Actual is within an approved tolerance of
'               Expected, otherwise FAIL.
'INPUTS         Expected, Actual, Tolerance, Floor: non-negative decimal text
'               with a point, at most 28 digits. Mode: "absolute" (difference
'               <= Tolerance) or "relative" (difference <= the larger of
'               Tolerance x Expected and Floor; Tolerance below 1; Floor
'               required). Floor is ignored for "absolute".
'ERROR POLICY   Text that is not such a decimal, an unknown mode, a missing
'               floor, or values whose difference or relative allowance
'               Decimal cannot hold exactly, is a FAIL with its reason; the
'               comparison is never made on a rounded value.
'UPDATED        2026-11-05
    Dim ExpectedValue   As Variant              'Decimal subtype
    Dim ActualValue     As Variant              'Decimal subtype
    Dim ToleranceValue  As Variant              'Decimal subtype
    Dim FloorValue      As Variant              'Decimal subtype
    Dim Allowed         As Variant              'Largest admitted difference, Decimal
    Dim Reason          As String               'Why the inputs are unusable, if they are

        If Not ParseDecimal(Expected, ExpectedValue) Then
            Reason = "not a decimal: " & Quoted(Expected)
        ElseIf Not ParseDecimal(Actual, ActualValue) Then
            Reason = "not a decimal: " & Quoted(Actual)
        ElseIf Not ParseDecimal(Tolerance, ToleranceValue) Then
            Reason = "not a decimal: " & Quoted(Tolerance)
        ElseIf Not SubtractsExactly(Expected, Actual) Then
            Reason = "expected and actual cannot be subtracted exactly in Decimal"
        ElseIf Mode = MODE_ABSOLUTE Then
            Allowed = ToleranceValue
        ElseIf Mode <> MODE_RELATIVE Then
            Reason = "unknown tolerance mode " & Quoted(Mode)
        ElseIf ToleranceValue >= 1 Then
            Reason = "a relative tolerance must be below 1"
        ElseIf Len(Floor) = 0 Then
            Reason = "a relative tolerance needs a floor"
        ElseIf Not ParseDecimal(Floor, FloorValue) Then
            Reason = "not a decimal: " & Quoted(Floor)
        ElseIf Not MultipliesExactly(Tolerance, Expected) Then
            Reason = "tolerance x expected cannot be computed exactly in Decimal"
        Else
            Allowed = ToleranceValue * ExpectedValue
            If Allowed < FloorValue Then Allowed = FloorValue
        End If

        If Len(Reason) > 0 Then
            RecordOutcome TestRun, CaseId, OUTCOME_FAIL, Reason
        ElseIf Abs(ExpectedValue - ActualValue) <= Allowed Then
            RecordOutcome TestRun, CaseId, OUTCOME_PASS, vbNullString
        Else
            RecordOutcome TestRun, CaseId, OUTCOME_FAIL, _
                          "got " & Quoted(Actual) & ", outside " & Mode & " " & _
                          Tolerance & " of " & Quoted(Expected)
        End If
End Sub


Public Sub CheckDiagnostic( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String, _
    ByVal ExpectedCodes As String, _
    ByVal ActualCodes As String)
'PURPOSE        Expected-error outcome: record PASS when the subject reported
'               exactly the expected diagnostics, otherwise FAIL.
'INPUTS         ExpectedCodes, ActualCodes: the diagnostics as text in the
'               subject's documented order, for example "NSFR-W001". An
'               empty ActualCodes means the subject reported none.
'UPDATED        2026-11-05

        If Len(ExpectedCodes) = 0 Then
            RecordOutcome TestRun, CaseId, OUTCOME_FAIL, "no expected diagnostic given"
        ElseIf Len(ActualCodes) = 0 Then
            RecordOutcome TestRun, CaseId, OUTCOME_FAIL, _
                          "expected diagnostic " & Quoted(ExpectedCodes) & ", none reported"
        ElseIf StrComp(ExpectedCodes, ActualCodes, vbBinaryCompare) = 0 Then
            RecordOutcome TestRun, CaseId, OUTCOME_PASS, vbNullString
        Else
            RecordOutcome TestRun, CaseId, OUTCOME_FAIL, _
                          "expected diagnostic " & Quoted(ExpectedCodes) & ", got " & _
                          Quoted(ActualCodes)
        End If
End Sub


Public Sub SkipCase( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String, _
    ByVal Reason As String)
'PURPOSE        Record a declared case as SKIPPED. A skipped case fails the run.
'UPDATED        2026-11-05

        RecordOutcome TestRun, CaseId, OUTCOME_SKIPPED, Reason
End Sub


Public Sub RecordCleanup( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String, _
    ByVal FailureText As String)
'PURPOSE        Record the cleanup result of a case, separately from its
'               outcome. An empty FailureText means cleanup succeeded.
'UPDATED        2026-11-05

        If Len(FailureText) = 0 Then Exit Sub
        TestRun.CleanupCount = TestRun.CleanupCount + 1
        ReDim Preserve TestRun.CleanupLines(1 To TestRun.CleanupCount)
        TestRun.CleanupLines(TestRun.CleanupCount) = CaseId & " " & FailureText
End Sub


'------------------------------------------------------------------------------
'                                  RESULTS
'------------------------------------------------------------------------------

Public Function Summary( _
    ByRef TestRun As NsfrTestRun) _
    As String
'PURPOSE        Build the run's summary: one line per declared case in ID
'               order, then cleanup failures and problems in text order, then
'               the counts and the result. Lines end with CR LF.
'RETURNS        The same text for the same recorded outcomes, whatever the
'               order in which they were recorded.
'UPDATED        2026-11-05
    Dim Ids()       As String                   'Declared IDs, sorted
    Dim Lines()     As String                   'Cleanup lines or problems, sorted
    Dim SummaryText As String                   'The summary being built
    Dim Index       As Long                     'Recorded index of a case, 0 if none
    Dim Counts(1 To 4) As Long                  'Passed, failed, skipped, missing
    Dim i           As Long                     'Line counter

        SummaryText = "SUITE " & TestRun.SuiteName & vbCrLf
        If TestRun.DeclaredCount > 0 Then
            Ids = SortedCopy(TestRun.DeclaredIds, TestRun.DeclaredCount)
            For i = 1 To TestRun.DeclaredCount
                Index = RecordedIndex(TestRun, Ids(i))
                If Index = 0 Then
                    SummaryText = SummaryText & Ids(i) & " " & OUTCOME_MISSING & vbCrLf
                ElseIf Len(TestRun.Details(Index)) = 0 Then
                    SummaryText = SummaryText & Ids(i) & " " & TestRun.Outcomes(Index) & vbCrLf
                Else
                    SummaryText = SummaryText & Ids(i) & " " & TestRun.Outcomes(Index) & " " & _
                           TestRun.Details(Index) & vbCrLf
                End If
            Next i
        End If
        If TestRun.CleanupCount > 0 Then
            Lines = SortedCopy(TestRun.CleanupLines, TestRun.CleanupCount)
            For i = 1 To TestRun.CleanupCount
                SummaryText = SummaryText & "CLEANUP FAILED " & Lines(i) & vbCrLf
            Next i
        End If
        If TestRun.ProblemCount > 0 Then
            Lines = SortedCopy(TestRun.Problems, TestRun.ProblemCount)
            For i = 1 To TestRun.ProblemCount
                SummaryText = SummaryText & "PROBLEM " & Lines(i) & vbCrLf
            Next i
        End If

        CountOutcomes TestRun, Counts
        SummaryText = SummaryText & "COUNTS declared=" & TestRun.DeclaredCount & _
               " executed=" & (Counts(1) + Counts(2)) & _
               " passed=" & Counts(1) & " failed=" & Counts(2) & _
               " skipped=" & Counts(3) & " missing=" & Counts(4) & _
               " undeclared=" & TestRun.UndeclaredCount & _
               " cleanup_failed=" & TestRun.CleanupCount & _
               " problems=" & TestRun.ProblemCount & vbCrLf
        If RunPassed(TestRun) Then
            Summary = SummaryText & "RESULT " & OUTCOME_PASS & vbCrLf
        Else
            Summary = SummaryText & "RESULT " & OUTCOME_FAIL & vbCrLf
        End If
End Function


Public Function RunPassed( _
    ByRef TestRun As NsfrTestRun) _
    As Boolean
'PURPOSE        True only when every declared case passed: none failed,
'               skipped or missing, no cleanup failed and no problem arose.
'UPDATED        2026-11-05
    Dim Counts(1 To 4) As Long                  'Passed, failed, skipped, missing

        CountOutcomes TestRun, Counts
        RunPassed = TestRun.DeclaredCount > 0 And Counts(1) = TestRun.DeclaredCount _
                    And TestRun.CleanupCount = 0 And TestRun.ProblemCount = 0
End Function


'------------------------------------------------------------------------------
'                                  HELPERS
'------------------------------------------------------------------------------

Private Sub RecordOutcome( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String, _
    ByVal Outcome As String, _
    ByVal Detail As String)
'PURPOSE        Record one outcome of a declared case, once.
'ERROR POLICY   An invalid, undeclared or repeated ID is a run-level problem
'               and the outcome is not recorded.

        If Not IsValidId(CaseId) Then
            AddProblem TestRun, "invalid case ID " & Quoted(CaseId)
        ElseIf Not IsDeclared(TestRun, CaseId) Then
            TestRun.UndeclaredCount = TestRun.UndeclaredCount + 1
            AddProblem TestRun, "case " & CaseId & " is not declared"
        ElseIf RecordedIndex(TestRun, CaseId) > 0 Then
            AddProblem TestRun, "case " & CaseId & " recorded twice"
        Else
            TestRun.CaseCount = TestRun.CaseCount + 1
            ReDim Preserve TestRun.CaseIds(1 To TestRun.CaseCount)
            ReDim Preserve TestRun.Outcomes(1 To TestRun.CaseCount)
            ReDim Preserve TestRun.Details(1 To TestRun.CaseCount)
            TestRun.CaseIds(TestRun.CaseCount) = CaseId
            TestRun.Outcomes(TestRun.CaseCount) = Outcome
            TestRun.Details(TestRun.CaseCount) = Detail
        End If
End Sub


Private Sub AddProblem( _
    ByRef TestRun As NsfrTestRun, _
    ByVal Problem As String)

        TestRun.ProblemCount = TestRun.ProblemCount + 1
        ReDim Preserve TestRun.Problems(1 To TestRun.ProblemCount)
        TestRun.Problems(TestRun.ProblemCount) = Problem
End Sub


Private Sub CountOutcomes( _
    ByRef TestRun As NsfrTestRun, _
    ByRef Counts() As Long)
'PURPOSE        Count declared cases: Counts(1) passed, (2) failed,
'               (3) skipped, (4) missing.
    Dim Index       As Long                     'Recorded index of a case, 0 if none
    Dim i           As Long                     'Declared case counter

        For i = 1 To 4
            Counts(i) = 0
        Next i
        For i = 1 To TestRun.DeclaredCount
            Index = RecordedIndex(TestRun, TestRun.DeclaredIds(i))
            If Index = 0 Then
                Counts(4) = Counts(4) + 1
            ElseIf TestRun.Outcomes(Index) = OUTCOME_PASS Then
                Counts(1) = Counts(1) + 1
            ElseIf TestRun.Outcomes(Index) = OUTCOME_FAIL Then
                Counts(2) = Counts(2) + 1
            Else
                Counts(3) = Counts(3) + 1
            End If
        Next i
End Sub


Private Function IsValidId( _
    ByVal CaseId As String) _
    As Boolean
'PURPOSE        True for 1 to 32 upper-case ASCII letters, digits and hyphens.
    Dim i           As Long                     'Character position

        If Len(CaseId) = 0 Or Len(CaseId) > 32 Then Exit Function
        For i = 1 To Len(CaseId)
            If Not Mid$(CaseId, i, 1) Like "[A-Z0-9-]" Then Exit Function
        Next i
        IsValidId = True
End Function


Private Function IsDeclared( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String) _
    As Boolean
    Dim i           As Long                     'Declared case counter

        For i = 1 To TestRun.DeclaredCount
            If StrComp(TestRun.DeclaredIds(i), CaseId, vbBinaryCompare) = 0 Then
                IsDeclared = True
                Exit Function
            End If
        Next i
End Function


Private Function RecordedIndex( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String) _
    As Long
'RETURNS        The index of the recorded case with this ID, or 0.
    Dim i           As Long                     'Recorded case counter

        For i = 1 To TestRun.CaseCount
            If StrComp(TestRun.CaseIds(i), CaseId, vbBinaryCompare) = 0 Then
                RecordedIndex = i
                Exit Function
            End If
        Next i
End Function


Private Function SortedCopy( _
    ByRef Values() As String, _
    ByVal Count As Long) _
    As String()
'PURPOSE        A copy of Values(1 To Count) in binary text order.
    Dim Sorted()    As String                   'The copy being sorted
    Dim Current     As String                   'The value being inserted
    Dim i           As Long                     'Outer position
    Dim j           As Long                     'Insertion position

        ReDim Sorted(1 To Count)
        For i = 1 To Count
            Current = Values(i)
            j = i - 1
            Do While j >= 1
                If StrComp(Sorted(j), Current, vbBinaryCompare) <= 0 Then Exit Do
                Sorted(j + 1) = Sorted(j)
                j = j - 1
            Loop
            Sorted(j + 1) = Current
        Next i
        SortedCopy = Sorted
End Function


Private Function ParseDecimal( _
    ByVal Source As String, _
    ByRef Parsed As Variant) _
    As Boolean
'PURPOSE        Convert non-negative decimal text with a point, such as
'               "6.766915", to a Decimal without using the system locale.
'INPUTS         Source: digits, optionally a point and more digits; at most
'               28 digits in all.
'RETURNS        True and Parsed (Decimal) when Source has that form; False
'               otherwise, leaving Parsed unchanged.
    Dim PointAt     As Long                     'Position of the point, 0 if none
    Dim WholeDigits As String                   'Digits before the point
    Dim Fraction    As String                   'Digits after the point
    Dim PowerOfTen  As Variant                  'Ten to the number of fraction digits, Decimal
    Dim i           As Long                     'Digit counter

        PointAt = InStr(1, Source, ".", vbBinaryCompare)
        If PointAt = 0 Then
            WholeDigits = Source
        Else
            WholeDigits = Left$(Source, PointAt - 1)
            Fraction = Mid$(Source, PointAt + 1)
            If Len(Fraction) = 0 Then Exit Function
        End If
        If Len(WholeDigits) = 0 Then Exit Function
        If Len(WholeDigits) + Len(Fraction) > MAX_DIGITS Then Exit Function
        If Not AllDigits(WholeDigits & Fraction) Then Exit Function

        PowerOfTen = CDec(1)
        For i = 1 To Len(Fraction)
            PowerOfTen = PowerOfTen * 10
        Next i
        Parsed = CDec(WholeDigits & Fraction) / PowerOfTen
        ParseDecimal = True
End Function


Private Function LineAt( _
    ByRef Lines() As String, _
    ByVal LineNumber As Long) _
    As String
'RETURNS        Line LineNumber (from 1) of a Split result, or empty past its end.

        If LineNumber - 1 <= UBound(Lines) Then LineAt = Lines(LineNumber - 1)
End Function


Private Function SubtractsExactly( _
    ByVal First As String, _
    ByVal Second As String) _
    As Boolean
'PURPOSE        True when two parsed decimals, aligned on the point, need at
'               most 28 digits, so that their difference is exact.

        SubtractsExactly = MaxOf(WholeDigits(First), WholeDigits(Second)) + _
                           MaxOf(FractionDigits(First), FractionDigits(Second)) <= MAX_DIGITS
End Function


Private Function MultipliesExactly( _
    ByVal First As String, _
    ByVal Second As String) _
    As Boolean
'PURPOSE        True when the product of two parsed decimals has at most 28
'               significant digits and 28 decimals, so that it is exact.

        MultipliesExactly = SignificantDigits(First) + SignificantDigits(Second) <= MAX_DIGITS _
                            And FractionDigits(First) + FractionDigits(Second) <= MAX_DIGITS
End Function


Private Function SignificantDigits( _
    ByVal Source As String) _
    As Long
'RETURNS        The digits of a parsed decimal without its point and leading
'               zeros; 1 for zero.
    Dim Digits      As String                   'Source without its point
    Dim i           As Long                     'First significant position

        Digits = Replace(Source, ".", vbNullString)
        i = 1
        Do While i < Len(Digits)
            If Mid$(Digits, i, 1) <> "0" Then Exit Do
            i = i + 1
        Loop
        SignificantDigits = Len(Digits) - i + 1
End Function


Private Function FractionDigits( _
    ByVal Source As String) _
    As Long
'RETURNS        The digits after the point of a parsed decimal, 0 if none.
    Dim PointAt     As Long                     'Position of the point, 0 if none

        PointAt = InStr(1, Source, ".", vbBinaryCompare)
        If PointAt > 0 Then FractionDigits = Len(Source) - PointAt
End Function


Private Function WholeDigits( _
    ByVal Source As String) _
    As Long
'RETURNS        The significant digits before the point of a parsed decimal.

        WholeDigits = MaxOf(SignificantDigits(Source) - FractionDigits(Source), 0)
End Function


Private Function MaxOf( _
    ByVal First As Long, _
    ByVal Second As Long) _
    As Long

        If First > Second Then
            MaxOf = First
        Else
            MaxOf = Second
        End If
End Function


Private Function AllDigits( _
    ByVal Candidate As String) _
    As Boolean
    Dim i           As Long                     'Character position

        For i = 1 To Len(Candidate)
            If Not Mid$(Candidate, i, 1) Like "[0-9]" Then Exit Function
        Next i
        AllDigits = True
End Function


Private Function Quoted( _
    ByVal Unquoted As String) _
    As String

        Quoted = """" & Unquoted & """"
End Function
