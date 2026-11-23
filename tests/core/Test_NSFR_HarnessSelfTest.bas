Attribute VB_Name = "Test_NSFR_HarnessSelfTest"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: Test_NSFR_HarnessSelfTest
'==============================================================================
' PURPOSE
'   Self-test of Test_NSFR_Harness. Builds a suite that passes and suites
'   that fail on purpose in every reported way, and checks their summaries
'   against exact expected texts
'
' WHY THIS EXISTS
'   The harness must behave as documented before any suite relies on it;
'   the self-test is the evidence recorded for a host run
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - RunSelfTest: build the demonstration suites DEMO-PASS, DEMO-FAIL
'       (twice, in opposite orders) and DEMO-DECLARE, then check their
'       summaries in the suite HARNESS-SELF-TEST
'
' ERROR POLICY
'   None raised; every mismatch is a FAIL in the returned summary
'
' DEPENDENCIES
'   Test_NSFR_Harness; VBA only, no Excel object model
'
' NOTES
'   State: none at module level.
'   Compatibility: Windows 11, Microsoft 365 desktop Excel, VBA7, 64-bit
'   (#14).
'   Run: in the Immediate window
'
'       ?Test_NSFR_HarnessSelfTest.RunSelfTest()
'
'   The last line must read RESULT PASS (tests/README.md).
'   Author: danielep71
'
' PUBLIC SURFACE
'   RunSelfTest
'
' UPDATED
'   2026-11-05
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const FAIL_IDS          As String = _
        "F-001,F-002,F-003,F-004,F-005,F-006,F-007,F-008,F-009,F-010,F-011,F-012,F-013"         'Cases of DEMO-FAIL
    Private Const SELF_TEST_IDS     As String = "H-001,H-002,H-003,H-004,H-005,H-006,H-007"     'Cases of HARNESS-SELF-TEST


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Function RunSelfTest() As String

'
'------------------------------------------------------------------------------
'                          RUN THE HARNESS SELF-TEST
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the harness self-test and returns its summary
'
' WHY THIS EXISTS
'   One call, by name, runs the whole self-test in the Immediate window or
'   through Application.Run from a host script
'
' INPUTS
'   None
'
' RETURNS
'   The summary of suite HARNESS-SELF-TEST; its last line is "RESULT PASS"
'   when the harness behaves as documented
'
' BEHAVIOR
'   Builds the demonstration suites, then checks:
'     - H-001, H-002: DEMO-PASS has its expected summary and passes
'     - H-003, H-004: DEMO-FAIL has its expected summary and fails
'     - H-005: DEMO-FAIL recorded in reverse gives the same summary
'     - H-006, H-007: DEMO-DECLARE has its expected summary and fails
'
' ERROR POLICY
'   None raised; every mismatch is a FAIL in the returned summary
'
' DEPENDENCIES
'   Test_NSFR_Harness: BeginRun, CheckText, CheckExact, Summary, RunPassed
'
' NOTES
'   Called by name: keep its name. tests/README.md documents the exact
'   output
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_HarnessSelfTest.RunSelfTest"

    Dim SelfTest                As NsfrTestRun  'The outer suite
    Dim PassDemo                As NsfrTestRun  'Suite expected to pass
    Dim FailDemo                As NsfrTestRun  'Suite expected to fail
    Dim Reversed                As NsfrTestRun  'Same as FailDemo, recorded in reverse
    Dim DeclareDemo             As NsfrTestRun  'Suite with invalid declarations

'------------------------------------------------------------------------------
' BUILD THE DEMONSTRATION SUITES
'------------------------------------------------------------------------------
    'Build each suite, DEMO-FAIL in both orders
        BuildPassDemo PassDemo
        BuildFailDemo FailDemo, False
        BuildFailDemo Reversed, True
        BuildDeclareDemo DeclareDemo

'------------------------------------------------------------------------------
' CHECK THE SUMMARIES
'------------------------------------------------------------------------------
    'Check each summary and verdict against its expected value
        BeginRun SelfTest, "HARNESS-SELF-TEST", SELF_TEST_IDS
        CheckText SelfTest, "H-001", ExpectedPassSummary(), Summary(PassDemo)
        CheckExact SelfTest, "H-002", "yes", YesNo(RunPassed(PassDemo))
        CheckText SelfTest, "H-003", ExpectedFailSummary(), Summary(FailDemo)
        CheckExact SelfTest, "H-004", "no", YesNo(RunPassed(FailDemo))
        CheckText SelfTest, "H-005", Summary(FailDemo), Summary(Reversed)
        CheckText SelfTest, "H-006", ExpectedDeclareSummary(), Summary(DeclareDemo)
        CheckExact SelfTest, "H-007", "no", YesNo(RunPassed(DeclareDemo))

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the self-test summary
        RunSelfTest = Summary(SelfTest)

End Function

'
'------------------------------------------------------------------------------
'
'                             DEMONSTRATION SUITES
'
'------------------------------------------------------------------------------
'

Private Sub BuildPassDemo( _
    ByRef Demo As NsfrTestRun)
'
'------------------------------------------------------------------------------
'                           BUILD THE PASSING SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   One passing case per outcome kind; cleanup succeeds
'
' INPUTS
'   Demo
'     The run to build (output): suite DEMO-PASS
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' BUILD THE SUITE
'------------------------------------------------------------------------------
    'Declare the cases
        BeginRun Demo, "DEMO-PASS", "D-001,D-002,D-003,D-004"
    'Record an exact, an absolute-tolerance, a diagnostic and a relative-tolerance pass
        CheckExact Demo, "D-001", "1.000000", "1.000000"
        CheckTolerance Demo, "D-002", "6.766915", "6.766916", _
            MODE_ABSOLUTE, "0.000001", vbNullString
        CheckDiagnostic Demo, "D-003", "NSFR-W001", "NSFR-W001"
        CheckTolerance Demo, "D-004", "0", "0.000001", MODE_RELATIVE, "0.01", "0.000001"
    'Record a successful cleanup
        RecordCleanup Demo, "D-001", vbNullString

End Sub

Private Sub BuildFailDemo( _
    ByRef Demo As NsfrTestRun, _
    ByVal InReverse As Boolean)
'
'------------------------------------------------------------------------------
'                           BUILD THE FAILING SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Fails on purpose in every reported way: wrong values, a value outside
'   its tolerance, a missing diagnostic, a skipped and a missing case, a
'   cleanup failure, a repeated and an undeclared case, a non-numeric
'   value, invalid tolerances and a relative allowance that Decimal cannot
'   compute exactly
'
' INPUTS
'   Demo
'     The run to build (output): suite DEMO-FAIL
'
'   InReverse
'     Record the steps in reverse order; the summary must not change
'
' DEPENDENCIES
'   FailStep
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_HarnessSelfTest.BuildFailDemo"

    Dim StepNumber              As Long         'Step number, 1 to 14

'------------------------------------------------------------------------------
' BUILD THE SUITE
'------------------------------------------------------------------------------
    'Declare the cases
        BeginRun Demo, "DEMO-FAIL", FAIL_IDS
    'Run the steps in the requested order
        If InReverse Then
            For StepNumber = 14 To 1 Step -1
                FailStep Demo, StepNumber
            Next StepNumber
        Else
            For StepNumber = 1 To 14
                FailStep Demo, StepNumber
            Next StepNumber
        End If

End Sub

Private Sub FailStep( _
    ByRef Demo As NsfrTestRun, _
    ByVal StepNumber As Long)
'
'------------------------------------------------------------------------------
'                     RECORD ONE STEP OF THE FAILING SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Records one step of suite DEMO-FAIL
'
' WHY THIS EXISTS
'   BuildFailDemo runs the same steps in both orders to show that the
'   summary does not depend on the recording order
'
' INPUTS
'   Demo
'     The run that records the step
'
'   StepNumber
'     The step, 1 to 14
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Steps:
'     - 1 to 4: a wrong value, a value outside its tolerance, a missing
'       diagnostic and a skipped case
'     - 5: a passing case F-006 with a cleanup failure
'     - 6 and 7: F-007 recorded twice
'     - 8 to 11: a non-numeric value and three invalid tolerances
'     - 12: a different diagnostic
'     - 13: a relative allowance that Decimal cannot compute exactly
'     - 14: an undeclared case X-999
'   F-005 is never recorded, so it is missing
'
' ERROR POLICY
'   None; another StepNumber records nothing
'
' DEPENDENCIES
'   Test_NSFR_Harness: CheckExact, CheckTolerance, CheckDiagnostic,
'   SkipCase, RecordCleanup
'
' NOTES
'   ExpectedFailSummary holds the summary these steps must produce
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD THE STEP
'------------------------------------------------------------------------------
    'Record the outcome of the requested step
        Select Case StepNumber
            Case 1
                CheckExact Demo, "F-001", "1.234568", "1.234567"
            Case 2
                CheckTolerance Demo, "F-002", "6.766915", "6.766917", _
                    MODE_ABSOLUTE, "0.000001", vbNullString
            Case 3
                CheckDiagnostic Demo, "F-003", "NSFR-W001", vbNullString
            Case 4
                SkipCase Demo, "F-004", "host not available"
            Case 5
                CheckExact Demo, "F-006", "0", "0"
                RecordCleanup Demo, "F-006", "restore failed: synthetic"
            Case 6, 7
                CheckExact Demo, "F-007", "1", "1"
            Case 8
                CheckTolerance Demo, "F-008", "1", "NaN", MODE_ABSOLUTE, "0.1", vbNullString
            Case 9
                CheckTolerance Demo, "F-009", "1", "1", MODE_RELATIVE, "1", "0"
            Case 10
                CheckTolerance Demo, "F-010", "1", "1", MODE_RELATIVE, "0.1", vbNullString
            Case 11
                CheckTolerance Demo, "F-011", "1", "1", "percent", "0.1", vbNullString
            Case 12
                CheckDiagnostic Demo, "F-012", "NSFR-E014", "NSFR-E015"
            Case 13
                CheckTolerance Demo, "F-013", "0.14", "0.140000000000000000000000001", _
                    MODE_RELATIVE, "0.000000000000000000000000007", "0"
            Case 14
                CheckExact Demo, "X-999", "1", "1"
            Case Else
                'No action for other step numbers
        End Select

End Sub

Private Sub BuildDeclareDemo( _
    ByRef Demo As NsfrTestRun)
'
'------------------------------------------------------------------------------
'                     BUILD THE INVALID DECLARATION SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   A repeated, a lower-case and an empty declared ID
'
' INPUTS
'   Demo
'     The run to build (output): suite DEMO-DECLARE
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' BUILD THE SUITE
'------------------------------------------------------------------------------
    'Declare a repeated, a lower-case and an empty ID
        BeginRun Demo, "DEMO-DECLARE", "A-1,A-1,bad,"
    'Record the one valid case
        CheckExact Demo, "A-1", "x", "x"

End Sub

'
'------------------------------------------------------------------------------
'
'                              EXPECTED SUMMARIES
'
'------------------------------------------------------------------------------
'

Private Function ExpectedPassSummary() As String

'
'------------------------------------------------------------------------------
'                     WRITE THE EXPECTED DEMO-PASS SUMMARY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the exact summary BuildPassDemo must produce
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_HarnessSelfTest.ExpectedPassSummary"

    Dim Expected                As String       'The text being built

'------------------------------------------------------------------------------
' BUILD THE TEXT
'------------------------------------------------------------------------------
    'Write the suite and case lines
        Expected = "SUITE DEMO-PASS" & vbCrLf
        Expected = Expected & "D-001 PASS" & vbCrLf
        Expected = Expected & "D-002 PASS" & vbCrLf
        Expected = Expected & "D-003 PASS" & vbCrLf
        Expected = Expected & "D-004 PASS" & vbCrLf
    'Write the counts
        Expected = Expected & "COUNTS declared=4 executed=4 passed=4 failed=0 skipped=0" & _
            " missing=0 undeclared=0 cleanup_failed=0 problems=0" & vbCrLf

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Close with the result
        ExpectedPassSummary = Expected & "RESULT PASS" & vbCrLf

End Function

Private Function ExpectedFailSummary() As String

'
'------------------------------------------------------------------------------
'                     WRITE THE EXPECTED DEMO-FAIL SUMMARY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the exact summary BuildFailDemo must produce, in either order
'
' WHY THIS EXISTS
'   The self-test compares the harness output with a text written by hand
'
' INPUTS
'   None
'
' RETURNS
'   The summary text, lines ending with CR LF
'
' BEHAVIOR
'   Builds the case lines, the cleanup failure, the problems, the counts
'   and RESULT FAIL
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   F-005 is missing, F-006 and F-007 pass, X-999 is undeclared
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_HarnessSelfTest.ExpectedFailSummary"

    Dim Expected                As String       'The text being built

'------------------------------------------------------------------------------
' BUILD THE TEXT
'------------------------------------------------------------------------------
    'Write the suite and case lines
        Expected = "SUITE DEMO-FAIL" & vbCrLf
        Expected = Expected & "F-001 FAIL expected ""1.234568"", got ""1.234567""" & vbCrLf
        Expected = Expected & "F-002 FAIL got ""6.766917"", outside absolute 0.000001 of" & _
            " ""6.766915""" & vbCrLf
        Expected = Expected & "F-003 FAIL expected diagnostic ""NSFR-W001"", none reported" & vbCrLf
        Expected = Expected & "F-004 SKIPPED host not available" & vbCrLf
        Expected = Expected & "F-005 MISSING" & vbCrLf
        Expected = Expected & "F-006 PASS" & vbCrLf
        Expected = Expected & "F-007 PASS" & vbCrLf
        Expected = Expected & "F-008 FAIL not a decimal: ""NaN""" & vbCrLf
        Expected = Expected & "F-009 FAIL a relative tolerance must be below 1" & vbCrLf
        Expected = Expected & "F-010 FAIL a relative tolerance needs a floor" & vbCrLf
        Expected = Expected & "F-011 FAIL unknown tolerance mode ""percent""" & vbCrLf
        Expected = Expected & "F-012 FAIL expected diagnostic ""NSFR-E014"", got" & _
            " ""NSFR-E015""" & vbCrLf
        Expected = Expected & "F-013 FAIL tolerance x expected cannot be computed exactly" & _
            " in Decimal" & vbCrLf
    'Write the cleanup failure and the problems
        Expected = Expected & "CLEANUP FAILED F-006 restore failed: synthetic" & vbCrLf
        Expected = Expected & "PROBLEM case F-007 recorded twice" & vbCrLf
        Expected = Expected & "PROBLEM case X-999 is not declared" & vbCrLf
    'Write the counts
        Expected = Expected & "COUNTS declared=13 executed=11 passed=2 failed=9 skipped=1" & _
            " missing=1 undeclared=1 cleanup_failed=1 problems=2" & vbCrLf

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Close with the result
        ExpectedFailSummary = Expected & "RESULT FAIL" & vbCrLf

End Function

Private Function ExpectedDeclareSummary() As String

'
'------------------------------------------------------------------------------
'                   WRITE THE EXPECTED DEMO-DECLARE SUMMARY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the exact summary BuildDeclareDemo must produce
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "Test_NSFR_HarnessSelfTest.ExpectedDeclareSummary"

    Dim Expected                As String       'The text being built

'------------------------------------------------------------------------------
' BUILD THE TEXT
'------------------------------------------------------------------------------
    'Write the suite and case line
        Expected = "SUITE DEMO-DECLARE" & vbCrLf
        Expected = Expected & "A-1 PASS" & vbCrLf
    'Write the problems
        Expected = Expected & "PROBLEM case A-1 declared twice" & vbCrLf
        Expected = Expected & "PROBLEM invalid case ID """"" & vbCrLf
        Expected = Expected & "PROBLEM invalid case ID ""bad""" & vbCrLf
    'Write the counts
        Expected = Expected & "COUNTS declared=1 executed=1 passed=1 failed=0 skipped=0" & _
            " missing=0 undeclared=0 cleanup_failed=0 problems=3" & vbCrLf

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Close with the result
        ExpectedDeclareSummary = Expected & "RESULT FAIL" & vbCrLf

End Function

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function YesNo( _
    ByVal Flag As Boolean) As String
'
'------------------------------------------------------------------------------
'                         WRITE A BOOLEAN AS YES OR NO
'------------------------------------------------------------------------------
' PURPOSE
'   "yes" or "no", independent of the system language
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Map the flag to fixed text
        If Flag Then
            YesNo = "yes"
        Else
            YesNo = "no"
        End If

End Function
