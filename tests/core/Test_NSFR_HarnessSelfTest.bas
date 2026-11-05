Attribute VB_Name = "Test_NSFR_HarnessSelfTest"
'==============================================================================
' MODULE        Test_NSFR_HarnessSelfTest
' PURPOSE       Self-test of Test_NSFR_Harness. Builds a suite that passes
'               and a suite that fails on purpose in every reported way, and
'               checks their summaries against exact expected texts.
' PUBLIC        RunSelfTest
' DEPENDENCIES  Test_NSFR_Harness; VBA only, no Excel object model.
' STATE         None at module level.
' ERRORS        None raised; every mismatch is a FAIL in the returned summary.
' COMPATIBILITY Windows 11, Microsoft 365 desktop Excel, VBA7, 64-bit (#14).
' RUN           In the Immediate window: ?Test_NSFR_HarnessSelfTest.RunSelfTest()
'               The last line must read RESULT PASS (tests/README.md).
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
Private Const FAIL_IDS As String = _
    "F-001,F-002,F-003,F-004,F-005,F-006,F-007,F-008,F-009,F-010,F-011,F-012,F-013"
Private Const SELF_TEST_IDS As String = "H-001,H-002,H-003,H-004,H-005,H-006,H-007"


'------------------------------------------------------------------------------
'                                 ENTRY POINT
'------------------------------------------------------------------------------

Public Function RunSelfTest() As String
'PURPOSE        Run the harness self-test and return its summary.
'RETURNS        The summary of suite HARNESS-SELF-TEST; its last line is
'               "RESULT PASS" when the harness behaves as documented.
'UPDATED        2026-11-05
    Dim SelfTest    As NsfrTestRun              'The outer suite
    Dim PassDemo    As NsfrTestRun              'Suite expected to pass
    Dim FailDemo    As NsfrTestRun              'Suite expected to fail
    Dim Reversed    As NsfrTestRun              'Same as FailDemo, recorded in reverse
    Dim DeclareDemo As NsfrTestRun              'Suite with invalid declarations

        BuildPassDemo PassDemo
        BuildFailDemo FailDemo, False
        BuildFailDemo Reversed, True
        BuildDeclareDemo DeclareDemo

        BeginRun SelfTest, "HARNESS-SELF-TEST", SELF_TEST_IDS
        CheckText SelfTest, "H-001", ExpectedPassSummary(), Summary(PassDemo)
        CheckExact SelfTest, "H-002", "yes", YesNo(RunPassed(PassDemo))
        CheckText SelfTest, "H-003", ExpectedFailSummary(), Summary(FailDemo)
        CheckExact SelfTest, "H-004", "no", YesNo(RunPassed(FailDemo))
        CheckText SelfTest, "H-005", Summary(FailDemo), Summary(Reversed)
        CheckText SelfTest, "H-006", ExpectedDeclareSummary(), Summary(DeclareDemo)
        CheckExact SelfTest, "H-007", "no", YesNo(RunPassed(DeclareDemo))
        RunSelfTest = Summary(SelfTest)
End Function


'------------------------------------------------------------------------------
'                              DEMONSTRATION SUITES
'------------------------------------------------------------------------------

Private Sub BuildPassDemo( _
    ByRef Demo As NsfrTestRun)
'PURPOSE        One passing case per outcome kind; cleanup succeeds.

        BeginRun Demo, "DEMO-PASS", "D-001,D-002,D-003,D-004"
        CheckExact Demo, "D-001", "1.000000", "1.000000"
        CheckTolerance Demo, "D-002", "6.766915", "6.766916", _
                       MODE_ABSOLUTE, "0.000001", vbNullString
        CheckDiagnostic Demo, "D-003", "NSFR-W001", "NSFR-W001"
        CheckTolerance Demo, "D-004", "0", "0.000001", MODE_RELATIVE, "0.01", "0.000001"
        RecordCleanup Demo, "D-001", vbNullString
End Sub


Private Sub BuildFailDemo( _
    ByRef Demo As NsfrTestRun, _
    ByVal InReverse As Boolean)
'PURPOSE        Fail on purpose in every reported way: wrong values, a value
'               outside its tolerance, a missing diagnostic, a skipped and a
'               missing case, a cleanup failure, a repeated and an undeclared
'               case, a non-numeric value, invalid tolerances and a relative
'               allowance that Decimal cannot compute exactly.
'INPUTS         InReverse: record the steps in reverse order; the summary
'               must not change.
    Dim StepNumber  As Long                     'Step number, 1 to 14

        BeginRun Demo, "DEMO-FAIL", FAIL_IDS
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
        End Select
End Sub


Private Sub BuildDeclareDemo( _
    ByRef Demo As NsfrTestRun)
'PURPOSE        A repeated, a lower-case and an empty declared ID.

        BeginRun Demo, "DEMO-DECLARE", "A-1,A-1,bad,"
        CheckExact Demo, "A-1", "x", "x"
End Sub


'------------------------------------------------------------------------------
'                              EXPECTED SUMMARIES
'------------------------------------------------------------------------------

Private Function ExpectedPassSummary() As String
    Dim Expected    As String                   'The text being built

        Expected = "SUITE DEMO-PASS" & vbCrLf
        Expected = Expected & "D-001 PASS" & vbCrLf
        Expected = Expected & "D-002 PASS" & vbCrLf
        Expected = Expected & "D-003 PASS" & vbCrLf
        Expected = Expected & "D-004 PASS" & vbCrLf
        Expected = Expected & "COUNTS declared=4 executed=4 passed=4 failed=0 skipped=0" & _
                   " missing=0 undeclared=0 cleanup_failed=0 problems=0" & vbCrLf
        ExpectedPassSummary = Expected & "RESULT PASS" & vbCrLf
End Function


Private Function ExpectedFailSummary() As String
    Dim Expected    As String                   'The text being built

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
        Expected = Expected & "CLEANUP FAILED F-006 restore failed: synthetic" & vbCrLf
        Expected = Expected & "PROBLEM case F-007 recorded twice" & vbCrLf
        Expected = Expected & "PROBLEM case X-999 is not declared" & vbCrLf
        Expected = Expected & "COUNTS declared=13 executed=11 passed=2 failed=9 skipped=1" & _
                   " missing=1 undeclared=1 cleanup_failed=1 problems=2" & vbCrLf
        ExpectedFailSummary = Expected & "RESULT FAIL" & vbCrLf
End Function


Private Function ExpectedDeclareSummary() As String
    Dim Expected    As String                   'The text being built

        Expected = "SUITE DEMO-DECLARE" & vbCrLf
        Expected = Expected & "A-1 PASS" & vbCrLf
        Expected = Expected & "PROBLEM case A-1 declared twice" & vbCrLf
        Expected = Expected & "PROBLEM invalid case ID """"" & vbCrLf
        Expected = Expected & "PROBLEM invalid case ID ""bad""" & vbCrLf
        Expected = Expected & "COUNTS declared=1 executed=1 passed=1 failed=0 skipped=0" & _
                   " missing=0 undeclared=0 cleanup_failed=0 problems=3" & vbCrLf
        ExpectedDeclareSummary = Expected & "RESULT FAIL" & vbCrLf
End Function


'------------------------------------------------------------------------------
'                                  HELPERS
'------------------------------------------------------------------------------

Private Function YesNo( _
    ByVal Flag As Boolean) _
    As String
'PURPOSE        "yes" or "no", independent of the system language.

        If Flag Then
            YesNo = "yes"
        Else
            YesNo = "no"
        End If
End Function
