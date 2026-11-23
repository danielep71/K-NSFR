Attribute VB_Name = "Test_NSFR_Harness"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: Test_NSFR_Harness
'==============================================================================
' PURPOSE
'   Neutral test harness: records the outcome of each declared case and
'   builds a deterministic text summary. It implements no NSFR calculation
'   and asserts nothing about one
'
' WHY THIS EXISTS
'   A suite needs one place that declares its cases, records one outcome per
'   case and prints a summary that depends only on the recorded outcomes
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - BeginRun: reset a run and declare its case IDs
'     - CheckExact, CheckText, CheckTolerance, CheckDiagnostic, SkipCase:
'       record one outcome per declared case
'     - RecordCleanup: record a cleanup failure apart from the outcome
'     - Summary, RunPassed: the summary text and whether the run passed
'
' ERROR POLICY
'   Validates instead of trapping: an invalid case ID, number or tolerance
'   becomes a failure or a problem in the summary
'
' DEPENDENCIES
'   VBA only; no Excel object model, file, clock or random number
'
' NOTES
'   State: none at module level; the caller owns each NsfrTestRun.
'   Compatibility: Windows 11, Microsoft 365 desktop Excel, VBA7, 64-bit
'   (#14).
'   Rules: tests/README.md; numbers follow docs/validation/VALIDATION.md.
'
' PUBLIC SURFACE
'   NsfrTestRun, BeginRun, CheckExact, CheckText, CheckTolerance,
'   CheckDiagnostic, SkipCase, RecordCleanup, Summary, RunPassed
'
' UPDATED
'   2026-11-05
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const MAX_DIGITS            As Long = 28            'Significant digits a Decimal holds exactly

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'Case outcomes; PASS and FAIL are also the run result
    Public Const OUTCOME_PASS           As String = "PASS"      'Case passed; run passed
    Public Const OUTCOME_FAIL           As String = "FAIL"      'Case failed; run failed
    Public Const OUTCOME_SKIPPED        As String = "SKIPPED"   'Case skipped, with a reason
    Public Const OUTCOME_MISSING        As String = "MISSING"   'Case declared but never recorded

    'Tolerance modes of CheckTolerance
    Public Const MODE_ABSOLUTE          As String = "absolute"  'Difference <= Tolerance
    Public Const MODE_RELATIVE          As String = "relative"  'Difference <= max(Tolerance x Expected, Floor)

'------------------------------------------------------------------------------
' PUBLIC TYPES
'------------------------------------------------------------------------------
    'One test run. The caller declares it, passes it ByRef to every procedure
    'of this module and never changes its members directly.
    Public Type NsfrTestRun
        SuiteName           As String           'Name on the first summary line
        DeclaredIds()       As String           'Declared case IDs, in declaration order
        DeclaredCount       As Long             'Number of declared case IDs
        CaseIds()           As String           'Recorded case IDs, in recording order
        Outcomes()          As String           'PASS, FAIL or SKIPPED per recorded case
        Details()           As String           'Failure or skip detail per recorded case
        CaseCount           As Long             'Number of recorded cases
        CleanupLines()      As String           'One "ID detail" text per cleanup failure
        CleanupCount        As Long             'Number of cleanup failures
        Problems()          As String           'TestRun-level problems
        ProblemCount        As Long             'Number of run-level problems
        UndeclaredCount     As Long             'Outcomes recorded for undeclared IDs
    End Type


'
'------------------------------------------------------------------------------
'
'                                RUN LIFECYCLE
'
'------------------------------------------------------------------------------
'

Public Sub BeginRun( _
    ByRef TestRun As NsfrTestRun, _
    ByVal SuiteName As String, _
    ByVal DeclaredIdList As String)
'
'------------------------------------------------------------------------------
'                                 START A RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Starts a run and declares every case it must execute
'
' WHY THIS EXISTS
'   A case that is declared but never recorded must show as MISSING, so the
'   run knows its cases before any outcome is recorded
'
' INPUTS
'   TestRun
'     The run to reset (output)
'
'   SuiteName
'     The name printed on the first summary line
'
'   DeclaredIdList
'     Comma-separated case IDs, each of upper-case letters, digits and
'     hyphens, for example "FX-0001,FX-0002"
'
' RETURNS
'   Nothing; resets TestRun
'
' BEHAVIOR
'   Resets TestRun to an empty run, stores SuiteName, then declares each ID
'   of the list in order
'
' ERROR POLICY
'   An invalid or repeated ID, or an empty list, is recorded as a run-level
'   problem; nothing is raised
'
' DEPENDENCIES
'   IsValidId, IsDeclared, AddProblem, Quoted
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
    Const PROC_NAME             As String = "Test_NSFR_Harness.BeginRun"

    Dim Fresh                   As NsfrTestRun  'An empty run to reset with
    Dim Parts()                 As String       'The split ID list
    Dim CaseId                  As String       'One declared ID
    Dim PartIndex               As Long         'Index into Parts

'------------------------------------------------------------------------------
' RESET THE RUN
'------------------------------------------------------------------------------
    'Start from an empty run named after the suite
        TestRun = Fresh
        TestRun.SuiteName = SuiteName

'------------------------------------------------------------------------------
' REGISTER THE CASE IDS
'------------------------------------------------------------------------------
    'Declare each valid ID once; record invalid and repeated IDs as problems
        If VBA.Len(DeclaredIdList) > 0 Then
            Parts = Split(DeclaredIdList, ",")
            For PartIndex = LBound(Parts) To UBound(Parts)
                CaseId = Parts(PartIndex)
                If Not IsValidId(CaseId) Then
                    AddProblem TestRun, "invalid case ID " & Quoted(CaseId)
                ElseIf IsDeclared(TestRun, CaseId) Then
                    AddProblem TestRun, "case " & CaseId & " declared twice"
                Else
                    TestRun.DeclaredCount = TestRun.DeclaredCount + 1
                    ReDim Preserve TestRun.DeclaredIds(1 To TestRun.DeclaredCount)
                    TestRun.DeclaredIds(TestRun.DeclaredCount) = CaseId
                End If
            Next PartIndex
        End If
    'Record a run without any declared case as a problem
        If TestRun.DeclaredCount = 0 Then AddProblem TestRun, "no case declared"

End Sub

'
'------------------------------------------------------------------------------
'
'                                   OUTCOMES
'
'------------------------------------------------------------------------------
'

Public Sub CheckExact( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String, _
    ByVal Expected As String, _
    ByVal Actual As String)
'
'------------------------------------------------------------------------------
'                             CHECK AN EXACT VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Records PASS when Actual equals Expected character for character
'   (binary comparison), otherwise FAIL
'
' WHY THIS EXISTS
'   Comparisons are exact by default; numbers are compared as text at their
'   reported scale
'
' INPUTS
'   TestRun
'     The run that records the outcome
'
'   CaseId
'     The declared case ID
'
'   Expected
'     The expected value as text, numbers at their reported scale
'
'   Actual
'     The actual value as text, numbers at their reported scale
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares the two texts in binary order and records the outcome; a FAIL
'   detail quotes both values
'
' ERROR POLICY
'   Nothing is raised; an invalid, undeclared or repeated case ID is a
'   run-level problem (RecordOutcome)
'
' DEPENDENCIES
'   RecordOutcome, Quoted
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD THE OUTCOME
'------------------------------------------------------------------------------
    'Pass on an exact match, otherwise fail quoting both values
        If VBA.StrComp(Expected, Actual, vbBinaryCompare) = 0 Then
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
'
'------------------------------------------------------------------------------
'                           CHECK A MULTI-LINE TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Records PASS when two multi-line texts (CR LF line ends) are equal,
'   otherwise FAIL naming the first line that differs
'
' WHY THIS EXISTS
'   A summary is several lines long; naming the first line that differs
'   makes a failure readable
'
' INPUTS
'   TestRun
'     The run that records the outcome
'
'   CaseId
'     The declared case ID
'
'   Expected
'     The expected text, lines separated by CR LF
'
'   Actual
'     The actual text, lines separated by CR LF
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares the whole texts in binary order. When they differ, compares
'   them line by line and records FAIL with the number of the first line
'   that differs and both versions of it; a line past the end of a text is
'   empty
'
' ERROR POLICY
'   Nothing is raised; an invalid, undeclared or repeated case ID is a
'   run-level problem (RecordOutcome)
'
' DEPENDENCIES
'   RecordOutcome, LineAt, Quoted
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
    Const PROC_NAME             As String = "Test_NSFR_Harness.CheckText"

    Dim ExpectedLines()         As String       'Expected split into lines
    Dim ActualLines()           As String       'Actual split into lines
    Dim ExpectedLine            As String       'Expected line at LineNumber, or empty
    Dim ActualLine              As String       'Actual line at LineNumber, or empty
    Dim LineNumber              As Long         'First line that differs, from 1

'------------------------------------------------------------------------------
' COMPARE THE WHOLE TEXTS
'------------------------------------------------------------------------------
    'Pass when the texts are equal
        If VBA.StrComp(Expected, Actual, vbBinaryCompare) = 0 Then
            RecordOutcome TestRun, CaseId, OUTCOME_PASS, vbNullString
            Exit Sub
        End If

'------------------------------------------------------------------------------
' FIND THE FIRST DIFFERENT LINE
'------------------------------------------------------------------------------
    'Split both texts into lines
        ExpectedLines = Split(Expected, vbCrLf)
        ActualLines = Split(Actual, vbCrLf)
    'Advance until a line differs or both texts end
        LineNumber = 1
        Do
            ExpectedLine = LineAt(ExpectedLines, LineNumber)
            ActualLine = LineAt(ActualLines, LineNumber)
            If VBA.StrComp(ExpectedLine, ActualLine, vbBinaryCompare) <> 0 Then Exit Do
            If LineNumber > UBound(ExpectedLines) + 1 And LineNumber > UBound(ActualLines) + 1 Then
                Exit Do
            End If
            LineNumber = LineNumber + 1
        Loop

'------------------------------------------------------------------------------
' RECORD THE OUTCOME
'------------------------------------------------------------------------------
    'Fail naming the line and both versions of it
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
'
'------------------------------------------------------------------------------
'                       CHECK A VALUE WITHIN A TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   Records PASS when Actual is within an approved tolerance of Expected,
'   otherwise FAIL
'
' WHY THIS EXISTS
'   A fixture may record an approved tolerance; the comparison must then be
'   exact in Decimal and independent of the system locale
'
' INPUTS
'   TestRun
'     The run that records the outcome
'
'   CaseId
'     The declared case ID
'
'   Expected
'     Non-negative decimal text with a point, at most 28 digits
'
'   Actual
'     Non-negative decimal text with a point, at most 28 digits
'
'   Mode
'     "absolute" (difference <= Tolerance) or "relative" (difference <= the
'     larger of Tolerance x Expected and Floor; Tolerance below 1; Floor
'     required)
'
'   Tolerance
'     Non-negative decimal text with a point, at most 28 digits
'
'   Floor
'     Non-negative decimal text with a point, at most 28 digits; ignored for
'     "absolute"
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses the values, checks that the difference and, for "relative", the
'   allowance are exact in Decimal, resolves the allowance, then records
'   PASS when the absolute difference is within it
'
' ERROR POLICY
'   Text that is not such a decimal, an unknown mode, a missing floor, or
'   values whose difference or relative allowance Decimal cannot hold
'   exactly, is a FAIL with its reason; the comparison is never made on a
'   rounded value. Nothing is raised
'
' DEPENDENCIES
'   ParseDecimal, SubtractsExactly, MultipliesExactly, RecordOutcome, Quoted
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
    Const PROC_NAME             As String = "Test_NSFR_Harness.CheckTolerance"

    Dim ExpectedValue           As Variant      'Decimal subtype
    Dim ActualValue             As Variant      'Decimal subtype
    Dim ToleranceValue          As Variant      'Decimal subtype
    Dim FloorValue              As Variant      'Decimal subtype
    Dim Allowed                 As Variant      'Largest admitted difference, Decimal
    Dim Reason                  As String       'Why the inputs are unusable, if they are

'------------------------------------------------------------------------------
' VALIDATE INPUTS AND RESOLVE THE ALLOWANCE
'------------------------------------------------------------------------------
    'Give the first reason the inputs are unusable, or the allowed difference
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
        ElseIf VBA.Len(Floor) = 0 Then
            Reason = "a relative tolerance needs a floor"
        ElseIf Not ParseDecimal(Floor, FloorValue) Then
            Reason = "not a decimal: " & Quoted(Floor)
        ElseIf Not MultipliesExactly(Tolerance, Expected) Then
            Reason = "tolerance x expected cannot be computed exactly in Decimal"
        Else
            Allowed = ToleranceValue * ExpectedValue
            If Allowed < FloorValue Then Allowed = FloorValue
        End If

'------------------------------------------------------------------------------
' RECORD THE OUTCOME
'------------------------------------------------------------------------------
    'Fail with the reason, pass within the allowance, otherwise fail
        If VBA.Len(Reason) > 0 Then
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
'
'------------------------------------------------------------------------------
'                         CHECK AN EXPECTED DIAGNOSTIC
'------------------------------------------------------------------------------
' PURPOSE
'   Expected-error outcome: records PASS when the subject reported exactly
'   the expected diagnostics, otherwise FAIL
'
' WHY THIS EXISTS
'   An expected error is an outcome of its own: reporting none, or another
'   diagnostic, must fail
'
' INPUTS
'   TestRun
'     The run that records the outcome
'
'   CaseId
'     The declared case ID
'
'   ExpectedCodes
'     The expected diagnostics as text in the subject's documented order,
'     for example "NSFR-W001"
'
'   ActualCodes
'     The reported diagnostics in the same form. An empty ActualCodes means
'     the subject reported none
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Fails when no expected diagnostic is given or none was reported, passes
'   on an exact binary match, otherwise fails quoting both
'
' ERROR POLICY
'   Nothing is raised; an invalid, undeclared or repeated case ID is a
'   run-level problem (RecordOutcome)
'
' DEPENDENCIES
'   RecordOutcome, Quoted
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD THE OUTCOME
'------------------------------------------------------------------------------
    'Require both sides, then pass only on an exact match
        If VBA.Len(ExpectedCodes) = 0 Then
            RecordOutcome TestRun, CaseId, OUTCOME_FAIL, "no expected diagnostic given"
        ElseIf VBA.Len(ActualCodes) = 0 Then
            RecordOutcome TestRun, CaseId, OUTCOME_FAIL, _
                "expected diagnostic " & Quoted(ExpectedCodes) & ", none reported"
        ElseIf VBA.StrComp(ExpectedCodes, ActualCodes, vbBinaryCompare) = 0 Then
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
'
'------------------------------------------------------------------------------
'                                 SKIP A CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Records a declared case as SKIPPED
'
' WHY THIS EXISTS
'   A case that cannot run must still be reported, with its reason
'
' INPUTS
'   TestRun
'     The run that records the outcome
'
'   CaseId
'     The declared case ID
'
'   Reason
'     Why the case was skipped, printed after SKIPPED
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Records the outcome SKIPPED with Reason as its detail
'
' ERROR POLICY
'   Nothing is raised; an invalid, undeclared or repeated case ID is a
'   run-level problem (RecordOutcome)
'
' DEPENDENCIES
'   RecordOutcome
'
' NOTES
'   A skipped case fails the run
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD THE OUTCOME
'------------------------------------------------------------------------------
    'Record the skip with its reason
        RecordOutcome TestRun, CaseId, OUTCOME_SKIPPED, Reason

End Sub

Public Sub RecordCleanup( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String, _
    ByVal FailureText As String)
'
'------------------------------------------------------------------------------
'                           RECORD A CLEANUP RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   Records the cleanup result of a case, separately from its outcome
'
' WHY THIS EXISTS
'   A cleanup failure fails the run even when every case passed
'
' INPUTS
'   TestRun
'     The run that records the cleanup result
'
'   CaseId
'     The case whose cleanup ran
'
'   FailureText
'     What failed. An empty FailureText means cleanup succeeded
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Does nothing for an empty FailureText; otherwise appends the line
'   "<CaseId> <FailureText>" to the cleanup failures
'
' ERROR POLICY
'   None; CaseId is not validated
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
' RECORD THE CLEANUP FAILURE
'------------------------------------------------------------------------------
    'Ignore a successful cleanup
        If VBA.Len(FailureText) = 0 Then Exit Sub
    'Append the failure line
        TestRun.CleanupCount = TestRun.CleanupCount + 1
        ReDim Preserve TestRun.CleanupLines(1 To TestRun.CleanupCount)
        TestRun.CleanupLines(TestRun.CleanupCount) = CaseId & " " & FailureText

End Sub

'
'------------------------------------------------------------------------------
'
'                                   RESULTS
'
'------------------------------------------------------------------------------
'

Public Function Summary( _
    ByRef TestRun As NsfrTestRun) As String
'
'------------------------------------------------------------------------------
'                            BUILD THE RUN SUMMARY
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the run's summary: one line per declared case in ID order, then
'   cleanup failures and problems in text order, then the counts and the
'   result
'
' WHY THIS EXISTS
'   The summary is the evidence of a run; it must be the same text for the
'   same outcomes
'
' INPUTS
'   TestRun
'     The run to summarize
'
' RETURNS
'   The summary text, lines ending with CR LF. The same text for the same
'   recorded outcomes, whatever the order in which they were recorded
'
' BEHAVIOR
'   Writes, in order:
'     - SUITE <name>
'     - <ID> <outcome> [<detail>] per declared case, MISSING if unrecorded
'     - CLEANUP FAILED <line> per cleanup failure
'     - PROBLEM <text> per run-level problem
'     - COUNTS declared=... executed=... passed=... failed=... skipped=...
'       missing=... undeclared=... cleanup_failed=... problems=...
'     - RESULT PASS or RESULT FAIL
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   SortedCopy, RecordedIndex, CountOutcomes, RunPassed
'
' NOTES
'   tests/README.md documents the exact text of the self-test's summary
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_Harness.Summary"

    Dim Ids()                   As String       'Declared IDs, sorted
    Dim Lines()                 As String       'Cleanup lines or problems, sorted
    Dim SummaryText             As String       'The summary being built
    Dim Index                   As Long         'Recorded index of a case, 0 if none
    Dim Counts(1 To 4)          As Long         'Passed, failed, skipped, missing
    Dim LineIndex               As Long         'Line counter

'------------------------------------------------------------------------------
' WRITE THE CASE LINES
'------------------------------------------------------------------------------
    'Name the suite
        SummaryText = "SUITE " & TestRun.SuiteName & vbCrLf
    'Write one line per declared case, in ID order
        If TestRun.DeclaredCount > 0 Then
            Ids = SortedCopy(TestRun.DeclaredIds, TestRun.DeclaredCount)
            For LineIndex = 1 To TestRun.DeclaredCount
                Index = RecordedIndex(TestRun, Ids(LineIndex))
                If Index = 0 Then
                    SummaryText = SummaryText & Ids(LineIndex) & " " & OUTCOME_MISSING & vbCrLf
                ElseIf VBA.Len(TestRun.Details(Index)) = 0 Then
                    SummaryText = SummaryText & Ids(LineIndex) & " " & _
                        TestRun.Outcomes(Index) & vbCrLf
                Else
                    SummaryText = SummaryText & Ids(LineIndex) & " " & _
                        TestRun.Outcomes(Index) & " " & TestRun.Details(Index) & vbCrLf
                End If
            Next LineIndex
        End If

'------------------------------------------------------------------------------
' WRITE THE CLEANUP FAILURES AND PROBLEMS
'------------------------------------------------------------------------------
    'Write the cleanup failures in text order
        If TestRun.CleanupCount > 0 Then
            Lines = SortedCopy(TestRun.CleanupLines, TestRun.CleanupCount)
            For LineIndex = 1 To TestRun.CleanupCount
                SummaryText = SummaryText & "CLEANUP FAILED " & Lines(LineIndex) & vbCrLf
            Next LineIndex
        End If
    'Write the run-level problems in text order
        If TestRun.ProblemCount > 0 Then
            Lines = SortedCopy(TestRun.Problems, TestRun.ProblemCount)
            For LineIndex = 1 To TestRun.ProblemCount
                SummaryText = SummaryText & "PROBLEM " & Lines(LineIndex) & vbCrLf
            Next LineIndex
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Write the counts
        CountOutcomes TestRun, Counts
        SummaryText = SummaryText & "COUNTS declared=" & TestRun.DeclaredCount & _
            " executed=" & (Counts(1) + Counts(2)) & _
            " passed=" & Counts(1) & " failed=" & Counts(2) & _
            " skipped=" & Counts(3) & " missing=" & Counts(4) & _
            " undeclared=" & TestRun.UndeclaredCount & _
            " cleanup_failed=" & TestRun.CleanupCount & _
            " problems=" & TestRun.ProblemCount & vbCrLf
    'Close the summary with the result
        If RunPassed(TestRun) Then
            Summary = SummaryText & "RESULT " & OUTCOME_PASS & vbCrLf
        Else
            Summary = SummaryText & "RESULT " & OUTCOME_FAIL & vbCrLf
        End If

End Function

Public Function RunPassed( _
    ByRef TestRun As NsfrTestRun) As Boolean
'
'------------------------------------------------------------------------------
'                         TELL WHETHER THE RUN PASSED
'------------------------------------------------------------------------------
' PURPOSE
'   True only when every declared case passed: none failed, skipped or
'   missing, no cleanup failed and no problem arose
'
' WHY THIS EXISTS
'   The result line of the summary and the caller need the same verdict
'
' INPUTS
'   TestRun
'     The run to judge
'
' RETURNS
'   True when at least one case is declared, every declared case passed,
'   and there is no cleanup failure and no run-level problem; False
'   otherwise
'
' BEHAVIOR
'   Counts the outcomes of the declared cases and applies the rule above
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CountOutcomes
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
    Const PROC_NAME             As String = "Test_NSFR_Harness.RunPassed"

    Dim Counts(1 To 4)          As Long         'Passed, failed, skipped, missing

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Pass only when every declared case passed and nothing else went wrong
        CountOutcomes TestRun, Counts
        RunPassed = TestRun.DeclaredCount > 0 And Counts(1) = TestRun.DeclaredCount _
            And TestRun.CleanupCount = 0 And TestRun.ProblemCount = 0

End Function

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub RecordOutcome( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String, _
    ByVal Outcome As String, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                              RECORD ONE OUTCOME
'------------------------------------------------------------------------------
' PURPOSE
'   Records one outcome of a declared case, once
'
' WHY THIS EXISTS
'   Every Check procedure and SkipCase record through one place that
'   enforces declared, unique case IDs
'
' INPUTS
'   TestRun
'     The run that records the outcome
'
'   CaseId
'     The declared case ID
'
'   Outcome
'     OUTCOME_PASS, OUTCOME_FAIL or OUTCOME_SKIPPED
'
'   Detail
'     The failure or skip detail, empty for none
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Appends CaseId, Outcome and Detail to the recorded cases
'
' ERROR POLICY
'   An invalid, undeclared or repeated ID is a run-level problem and the
'   outcome is not recorded; an undeclared ID is also counted
'
' DEPENDENCIES
'   IsValidId, IsDeclared, RecordedIndex, AddProblem, Quoted
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD
'------------------------------------------------------------------------------
    'Reject an invalid, undeclared or repeated ID; otherwise append the outcome
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
'
'------------------------------------------------------------------------------
'                          RECORD A RUN-LEVEL PROBLEM
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one run-level problem to the run
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD
'------------------------------------------------------------------------------
    'Append the problem
        TestRun.ProblemCount = TestRun.ProblemCount + 1
        ReDim Preserve TestRun.Problems(1 To TestRun.ProblemCount)
        TestRun.Problems(TestRun.ProblemCount) = Problem

End Sub

Private Sub CountOutcomes( _
    ByRef TestRun As NsfrTestRun, _
    ByRef Counts() As Long)
'
'------------------------------------------------------------------------------
'                              COUNT THE OUTCOMES
'------------------------------------------------------------------------------
' PURPOSE
'   Counts declared cases: Counts(1) passed, (2) failed, (3) skipped,
'   (4) missing
'
' WHY THIS EXISTS
'   Summary and RunPassed need the same counts
'
' INPUTS
'   TestRun
'     The run to count
'
'   Counts
'     Output: an array dimensioned 1 To 4
'
' RETURNS
'   Nothing; fills Counts
'
' BEHAVIOR
'   Resets the four counts, then classifies each declared case by its
'   recorded outcome; an unrecorded case is missing, and an outcome other
'   than PASS or FAIL counts as skipped
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   RecordedIndex
'
' NOTES
'   Outcomes recorded for undeclared IDs are not counted here
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_Harness.CountOutcomes"

    Dim Index                   As Long         'Recorded index of a case, 0 if none
    Dim ItemIndex               As Long         'Count index, then declared case counter

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Reset the counts
        For ItemIndex = 1 To 4
            Counts(ItemIndex) = 0
        Next ItemIndex
    'Classify each declared case
        For ItemIndex = 1 To TestRun.DeclaredCount
            Index = RecordedIndex(TestRun, TestRun.DeclaredIds(ItemIndex))
            If Index = 0 Then
                Counts(4) = Counts(4) + 1
            ElseIf TestRun.Outcomes(Index) = OUTCOME_PASS Then
                Counts(1) = Counts(1) + 1
            ElseIf TestRun.Outcomes(Index) = OUTCOME_FAIL Then
                Counts(2) = Counts(2) + 1
            Else
                Counts(3) = Counts(3) + 1
            End If
        Next ItemIndex

End Sub

Private Function IsValidId( _
    ByVal CaseId As String) As Boolean
'
'------------------------------------------------------------------------------
'                              VALIDATE A CASE ID
'------------------------------------------------------------------------------
' PURPOSE
'   True for 1 to 32 upper-case ASCII letters, digits and hyphens
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_Harness.IsValidId"

    Dim CharIndex               As Long         'Character position

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject an empty or too long ID
        If VBA.Len(CaseId) = 0 Or VBA.Len(CaseId) > 32 Then Exit Function
    'Reject any character outside the allowed set
        For CharIndex = 1 To VBA.Len(CaseId)
            If Not VBA.Mid$(CaseId, CharIndex, 1) Like "[A-Z0-9-]" Then Exit Function
        Next CharIndex
    'Accept the ID
        IsValidId = True

End Function

Private Function IsDeclared( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String) As Boolean
'
'------------------------------------------------------------------------------
'                       TELL WHETHER A CASE IS DECLARED
'------------------------------------------------------------------------------
' PURPOSE
'   True when CaseId is among the run's declared IDs (binary comparison)
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_Harness.IsDeclared"

    Dim DeclaredIndex           As Long         'Declared case counter

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Return True at the first declared ID that matches
        For DeclaredIndex = 1 To TestRun.DeclaredCount
            If VBA.StrComp(TestRun.DeclaredIds(DeclaredIndex), CaseId, vbBinaryCompare) = 0 Then
                IsDeclared = True
                Exit Function
            End If
        Next DeclaredIndex

End Function

Private Function RecordedIndex( _
    ByRef TestRun As NsfrTestRun, _
    ByVal CaseId As String) As Long
'
'------------------------------------------------------------------------------
'                             FIND A RECORDED CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Finds the recorded case with this ID (binary comparison)
'
' RETURNS
'   The index of the recorded case with this ID, or 0
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_Harness.RecordedIndex"

    Dim CaseIndex               As Long         'Recorded case counter

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Return the index of the first recorded ID that matches
        For CaseIndex = 1 To TestRun.CaseCount
            If VBA.StrComp(TestRun.CaseIds(CaseIndex), CaseId, vbBinaryCompare) = 0 Then
                RecordedIndex = CaseIndex
                Exit Function
            End If
        Next CaseIndex

End Function

Private Function SortedCopy( _
    ByRef Values() As String, _
    ByVal Count As Long) As String()
'
'------------------------------------------------------------------------------
'                          SORT A COPY OF A TEXT LIST
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a copy of Values(1 To Count) in binary text order
'
' WHY THIS EXISTS
'   The summary lists cases, cleanup failures and problems in a fixed order,
'   whatever the order in which they were recorded
'
' INPUTS
'   Values
'     The texts, indexed from 1
'
'   Count
'     The number of texts to copy, at least 1
'
' RETURNS
'   A new array (1 To Count) with the texts in binary order
'
' BEHAVIOR
'   Insertion sort into the copy; Values is not changed
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   Equal texts keep their relative order
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_Harness.SortedCopy"

    Dim Sorted()                As String       'The copy being sorted
    Dim Current                 As String       'The value being inserted
    Dim OuterIndex              As Long         'Outer position
    Dim InsertAt                As Long         'Insertion position

'------------------------------------------------------------------------------
' SORT
'------------------------------------------------------------------------------
    'Insert each value after the last smaller or equal one
        ReDim Sorted(1 To Count)
        For OuterIndex = 1 To Count
            Current = Values(OuterIndex)
            InsertAt = OuterIndex - 1
            Do While InsertAt >= 1
                If VBA.StrComp(Sorted(InsertAt), Current, vbBinaryCompare) <= 0 Then Exit Do
                Sorted(InsertAt + 1) = Sorted(InsertAt)
                InsertAt = InsertAt - 1
            Loop
            Sorted(InsertAt + 1) = Current
        Next OuterIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the sorted copy
        SortedCopy = Sorted

End Function

Private Function ParseDecimal( _
    ByVal Source As String, _
    ByRef Parsed As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                              PARSE DECIMAL TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Converts non-negative decimal text with a point, such as "6.766915", to
'   a Decimal without using the system locale
'
' WHY THIS EXISTS
'   The summary must not depend on the system locale, so decimal text is
'   converted digit by digit rather than by the locale's number parsing
'
' INPUTS
'   Source
'     Digits, optionally a point and more digits; at most 28 digits in all
'
'   Parsed
'     Output: the Decimal value
'
' RETURNS
'   True and Parsed (Decimal) when Source has that form; False otherwise,
'   leaving Parsed unchanged
'
' BEHAVIOR
'   Splits Source at the point, checks the digits, converts all digits as a
'   whole number and divides by ten to the number of fraction digits
'
' ERROR POLICY
'   None; text of another form returns False
'
' DEPENDENCIES
'   AllDigits
'
' NOTES
'   A point with no digit after it, or no digit before it, is rejected
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_Harness.ParseDecimal"

    Dim PointAt                 As Long         'Position of the point, 0 if none
    Dim WholeText               As String       'Digits before the point
    Dim Fraction                As String       'Digits after the point
    Dim PowerOfTen              As Variant      'Ten to the number of fraction digits, Decimal
    Dim DigitIndex              As Long         'Digit counter

'------------------------------------------------------------------------------
' SPLIT AT THE POINT
'------------------------------------------------------------------------------
    'Separate the whole and fraction digits; reject a point without fraction digits
        PointAt = VBA.InStr(1, Source, ".", vbBinaryCompare)
        If PointAt = 0 Then
            WholeText = Source
        Else
            WholeText = VBA.Left$(Source, PointAt - 1)
            Fraction = VBA.Mid$(Source, PointAt + 1)
            If VBA.Len(Fraction) = 0 Then Exit Function
        End If

'------------------------------------------------------------------------------
' VALIDATE THE DIGITS
'------------------------------------------------------------------------------
    'Require whole digits, at most MAX_DIGITS digits in all, and digits only
        If VBA.Len(WholeText) = 0 Then Exit Function
        If VBA.Len(WholeText) + VBA.Len(Fraction) > MAX_DIGITS Then Exit Function
        If Not AllDigits(WholeText & Fraction) Then Exit Function

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Compute ten to the number of fraction digits
        PowerOfTen = CDec(1)
        For DigitIndex = 1 To VBA.Len(Fraction)
            PowerOfTen = PowerOfTen * 10
        Next DigitIndex
    'Scale the whole number of digits down to the value
        Parsed = CDec(WholeText & Fraction) / PowerOfTen
        ParseDecimal = True

End Function

Private Function LineAt( _
    ByRef Lines() As String, _
    ByVal LineNumber As Long) As String
'
'------------------------------------------------------------------------------
'                        READ ONE LINE OF A SPLIT TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Reads one line of a Split result
'
' RETURNS
'   Line LineNumber (from 1) of a Split result, or empty past its end
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the line when it exists
        If LineNumber - 1 <= UBound(Lines) Then LineAt = Lines(LineNumber - 1)

End Function

Private Function SubtractsExactly( _
    ByVal First As String, _
    ByVal Second As String) As Boolean
'
'------------------------------------------------------------------------------
'                      TELL WHETHER A DIFFERENCE IS EXACT
'------------------------------------------------------------------------------
' PURPOSE
'   True when two parsed decimals, aligned on the point, need at most 28
'   digits, so that their difference is exact
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Add the longer whole part and the longer fraction
        SubtractsExactly = MaxOf(WholeDigits(First), WholeDigits(Second)) + _
            MaxOf(FractionDigits(First), FractionDigits(Second)) <= MAX_DIGITS

End Function

Private Function MultipliesExactly( _
    ByVal First As String, _
    ByVal Second As String) As Boolean
'
'------------------------------------------------------------------------------
'                       TELL WHETHER A PRODUCT IS EXACT
'------------------------------------------------------------------------------
' PURPOSE
'   True when the product of two parsed decimals has at most 28 significant
'   digits and 28 decimals, so that it is exact
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Bound both the significant digits and the decimals of the product
        MultipliesExactly = SignificantDigits(First) + SignificantDigits(Second) <= MAX_DIGITS _
            And FractionDigits(First) + FractionDigits(Second) <= MAX_DIGITS

End Function

Private Function SignificantDigits( _
    ByVal Source As String) As Long
'
'------------------------------------------------------------------------------
'                           COUNT SIGNIFICANT DIGITS
'------------------------------------------------------------------------------
' PURPOSE
'   Counts the significant digits of a parsed decimal
'
' RETURNS
'   The digits of a parsed decimal without its point and leading zeros; 1
'   for zero
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_Harness.SignificantDigits"

    Dim Digits                  As String       'Source without its point
    Dim FirstSignificant        As Long         'First significant position

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Drop the point
        Digits = Replace(Source, ".", vbNullString)
    'Skip the leading zeros, keeping at least the last digit
        FirstSignificant = 1
        Do While FirstSignificant < VBA.Len(Digits)
            If VBA.Mid$(Digits, FirstSignificant, 1) <> "0" Then Exit Do
            FirstSignificant = FirstSignificant + 1
        Loop
    'Count the digits from the first significant one
        SignificantDigits = VBA.Len(Digits) - FirstSignificant + 1

End Function

Private Function FractionDigits( _
    ByVal Source As String) As Long
'
'------------------------------------------------------------------------------
'                            COUNT FRACTION DIGITS
'------------------------------------------------------------------------------
' PURPOSE
'   Counts the digits after the point of a parsed decimal
'
' RETURNS
'   The digits after the point of a parsed decimal, 0 if none
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_Harness.FractionDigits"

    Dim PointAt                 As Long         'Position of the point, 0 if none

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Count the characters after the point
        PointAt = VBA.InStr(1, Source, ".", vbBinaryCompare)
        If PointAt > 0 Then FractionDigits = VBA.Len(Source) - PointAt

End Function

Private Function WholeDigits( _
    ByVal Source As String) As Long
'
'------------------------------------------------------------------------------
'                              COUNT WHOLE DIGITS
'------------------------------------------------------------------------------
' PURPOSE
'   Counts the significant digits before the point of a parsed decimal
'
' RETURNS
'   The significant digits before the point of a parsed decimal
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Subtract the fraction digits from the significant digits, at least 0
        WholeDigits = MaxOf(SignificantDigits(Source) - FractionDigits(Source), 0)

End Function

Private Function MaxOf( _
    ByVal First As Long, _
    ByVal Second As Long) As Long
'
'------------------------------------------------------------------------------
'                       RETURN THE LARGER OF TWO NUMBERS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the larger of two numbers
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the larger value
        If First > Second Then
            MaxOf = First
        Else
            MaxOf = Second
        End If

End Function

Private Function AllDigits( _
    ByVal Candidate As String) As Boolean
'
'------------------------------------------------------------------------------
'                      TELL WHETHER A TEXT IS ALL DIGITS
'------------------------------------------------------------------------------
' PURPOSE
'   True when every character of Candidate is a digit 0 to 9
'
' NOTES
'   True for an empty Candidate
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_Harness.AllDigits"

    Dim CharIndex               As Long         'Character position

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject any character that is not a digit
        For CharIndex = 1 To VBA.Len(Candidate)
            If Not VBA.Mid$(Candidate, CharIndex, 1) Like "[0-9]" Then Exit Function
        Next CharIndex
    'Accept the text
        AllDigits = True

End Function

Private Function Quoted( _
    ByVal Unquoted As String) As String
'
'------------------------------------------------------------------------------
'                                 QUOTE A TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Wraps a text in double quotes for a summary detail
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Add a double quote on each side
        Quoted = """" & Unquoted & """"

End Function
