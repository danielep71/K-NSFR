Attribute VB_Name = "NSFR_Core_Util"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: NSFR_Core_Util
'------------------------------------------------------------------------------
' PURPOSE
'   Pure helpers of the prototype: value conversion of worksheet values,
'   residual-maturity buckets, table look-ups, growable lists and the
'   tolerance test of the reconciliation
'
' WHY THIS EXISTS
'   The engine, the import, the reconciliation and the API share these
'   conversions and list operations, so each rule is written once
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - value conversion: TextOf, IsBlankValue, NumberOf, DateOf,
'       IsAllDigits, IsCode4
'     - buckets: MaturityBucket, BucketCode, BucketFromCode, HasBucket,
'       IdDepth
'     - look-ups: FindRow2, FindRow1
'     - lists: ListAdd, ListToMatrix, NewRow, AddDiagnostic, CountSeverity
'     - comparison: ToleranceResult, FormatAmount
'
' ERROR POLICY
'   No On Error. A value that cannot be converted is reported through the
'   Valid output argument, never turned into zero
'
' DEPENDENCIES
'   NSFR_Core_Config for the bucket letters and the diagnostic columns
'
' NOTES
'   Excel VBA7, 64-bit; no references beyond the defaults
'
' PUBLIC SURFACE
'   None outside this VBA project (Option Private Module)
'
' STATE OWNERSHIP
'   None: every procedure depends only on its arguments
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------


'
'------------------------------------------------------------------------------
'
'                               VALUE CONVERSION
'
'------------------------------------------------------------------------------
'

Public Function TextOf( _
    ByVal Value As Variant) As String
'
'------------------------------------------------------------------------------
'                          TEXT OF A WORKSHEET VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the trimmed text of a worksheet value
'
' WHY THIS EXISTS
'   A cell value can be Empty, Null, an error value or an array; callers
'   need one text form that never fails on these
'
' INPUTS
'   Value
'     A cell value
'
' RETURNS
'   The trimmed text; vbNullString for Empty, Null, an error value or an
'   array
'
' BEHAVIOR
'   Returns vbNullString for Empty, Null, an error value or an array;
'   otherwise converts Value with CStr and trims it
'
' ERROR POLICY
'   No On Error; the cases above never raise
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Give no text for a value without one, else the trimmed text
        If IsError(Value) Or IsNull(Value) Or IsEmpty(Value) Then
            TextOf = vbNullString
        ElseIf IsArray(Value) Then
            TextOf = vbNullString
        Else
            TextOf = Trim$(CStr(Value))
        End If

End Function

Public Function IsBlankValue( _
    ByVal Value As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                         TELL A BLANK WORKSHEET VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   True for Empty, Null or text that is empty once trimmed
'
' WHY THIS EXISTS
'   A blank input and an unusable input are reported differently, so an
'   error value must not count as blank
'
' INPUTS
'   Value
'     A cell value
'
' RETURNS
'   True for Empty, Null or text that is empty once trimmed; False for an
'   error value and for any other value
'
' BEHAVIOR
'   Tests, in order: an error value, Empty or Null, then text
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   An error value is not blank: it is an unusable input
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Classify the value
        If IsError(Value) Then
            IsBlankValue = False
        ElseIf IsEmpty(Value) Or IsNull(Value) Then
            IsBlankValue = True
        ElseIf VarType(Value) = vbString Then
            IsBlankValue = (Len(Trim$(Value)) = 0)
        Else
            IsBlankValue = False
        End If

End Function

Public Function NumberOf( _
    ByVal Value As Variant, _
    ByRef Valid As Boolean) As Double
'
'------------------------------------------------------------------------------
'                         NUMBER OF A WORKSHEET VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the numeric value of a worksheet value
'
' WHY THIS EXISTS
'   Text, even if it looks numeric, is rejected so that the result never
'   depends on the locale
'
' INPUTS
'   Value
'     A cell value. Only numeric types are accepted
'
'   Valid
'     Output: True when Value is numeric
'
' RETURNS
'   The number; 0 with Valid = False when Value is not numeric
'
' BEHAVIOR
'   Resets the outputs, rejects an error value or an array, then converts
'   the numeric variant types (Double, Single, Integer, Long, Currency,
'   Decimal, Byte) with CDbl
'
' ERROR POLICY
'   No On Error; an unusable value is reported through Valid
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESET OUTPUTS
'------------------------------------------------------------------------------
    'Start from an invalid zero
        Valid = False
        NumberOf = 0

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an error value or an array
        If IsError(Value) Or IsArray(Value) Then Exit Function

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'Accept the numeric types only
        Select Case VarType(Value)
            Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte
                NumberOf = CDbl(Value)
                Valid = True
            Case Else
                'No action for other values
        End Select

End Function

Public Function DateOf( _
    ByVal Value As Variant, _
    ByRef Valid As Boolean) As Date
'
'------------------------------------------------------------------------------
'                          DATE OF A WORKSHEET VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the date of a worksheet value: an Excel date, a positive date
'   serial or text in the form yyyy-mm-dd
'
' WHY THIS EXISTS
'   Dates reach the prototype as Excel dates, serial numbers or ISO text;
'   other text is rejected rather than parsed in the locale
'
' INPUTS
'   Value
'     A cell value
'
'   Valid
'     Output: True when Value is one of the accepted forms
'
' RETURNS
'   The date, without a time part; Valid = False when Value is none of the
'   accepted forms
'
' BEHAVIOR
'   Rejects an error value or an array, then:
'     - a Date: its whole day
'     - a number from 1 to below 2958466: the whole day of that serial
'     - text of 10 characters yyyy-mm-dd with digits in each part, a month
'       of 1 to 12 and a day of 1 to 31: the date, valid only when the
'       day exists in that month
'
' ERROR POLICY
'   No On Error; an unusable value is reported through Valid
'
' DEPENDENCIES
'   IsAllDigits
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
    Const PROC_NAME             As String = "NSFR_Core_Util.DateOf"

    Dim Text                    As String       'Value as text
    Dim YearPart                As Long         'yyyy
    Dim MonthPart               As Long         'mm
    Dim DayPart                 As Long         'dd

'------------------------------------------------------------------------------
' RESET OUTPUTS
'------------------------------------------------------------------------------
    'Start from an invalid result
        Valid = False

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an error value or an array
        If IsError(Value) Or IsArray(Value) Then Exit Function

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'Convert by variant type
        Select Case VarType(Value)
            Case vbDate
                DateOf = Int(CDbl(Value))
                Valid = True
            Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency
                If CDbl(Value) >= 1 And CDbl(Value) < 2958466 Then
                    DateOf = CDate(Int(CDbl(Value)))
                    Valid = True
                End If
            Case vbString
                Text = Trim$(Value)
                If Len(Text) = 10 And Mid$(Text, 5, 1) = "-" _
                    And Mid$(Text, 8, 1) = "-" Then
                    If IsAllDigits(Left$(Text, 4)) And IsAllDigits(Mid$(Text, 6, 2)) _
                        And IsAllDigits(Right$(Text, 2)) Then
                        YearPart = CLng(Left$(Text, 4))
                        MonthPart = CLng(Mid$(Text, 6, 2))
                        DayPart = CLng(Right$(Text, 2))
                        'Keep a day that exists in the month: DateSerial rolls an overflow over
                        If MonthPart >= 1 And MonthPart <= 12 And DayPart >= 1 And DayPart <= 31 Then
                            DateOf = DateSerial(YearPart, MonthPart, DayPart)
                            Valid = (Day(DateOf) = DayPart)
                        End If
                    End If
                End If
            Case Else
                'No action for other values
        End Select

End Function

Public Function IsAllDigits( _
    ByVal Text As String) As Boolean
'
'------------------------------------------------------------------------------
'                          TELL A TEXT OF DIGITS ONLY
'------------------------------------------------------------------------------
' PURPOSE
'   True when Text is not empty and holds the digits 0-9 only
'
' WHY THIS EXISTS
'   Codes and date parts are checked character by character, without a
'   locale-dependent numeric conversion
'
' INPUTS
'   Text
'     The text to test
'
' RETURNS
'   True when Text is not empty and every character is 0 to 9; False
'   otherwise
'
' BEHAVIOR
'   Returns False at the first character outside 0 to 9
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   VBA library only
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
    Const PROC_NAME             As String = "NSFR_Core_Util.IsAllDigits"

    Dim CharIndex               As Long         'Position in Text

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject empty text
        If Len(Text) = 0 Then Exit Function

'------------------------------------------------------------------------------
' TEST EACH CHARACTER
'------------------------------------------------------------------------------
    'Stop at the first character that is not a digit
        For CharIndex = 1 To Len(Text)
            If Mid$(Text, CharIndex, 1) < "0" Or _
                Mid$(Text, CharIndex, 1) > "9" Then Exit Function
        Next CharIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Every character is a digit
        IsAllDigits = True

End Function

Public Function IsCode4( _
    ByVal Value As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                            TELL A FOUR-DIGIT CODE
'------------------------------------------------------------------------------
' PURPOSE
'   True for a four-digit row or column code held as text, such as "0010"
'
' WHY THIS EXISTS
'   The reporting templates hold their row and column codes in this form
'
' INPUTS
'   Value
'     A cell value
'
' RETURNS
'   True when Value is text of four digits once trimmed; False otherwise,
'   also for a number
'
' BEHAVIOR
'   Rejects any value that is not text, then tests the trimmed text
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   IsAllDigits
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Accept text only
        If VarType(Value) <> vbString Then Exit Function
    'Require four digits once trimmed
        IsCode4 = (Len(Trim$(Value)) = 4 And IsAllDigits(Trim$(Value)))

End Function

'
'------------------------------------------------------------------------------
'
'                                   BUCKETS
'
'------------------------------------------------------------------------------
'

Public Function MaturityBucket( _
    ByVal ReportingDate As Date, _
    ByVal EndDate As Date) As Long
'
'------------------------------------------------------------------------------
'                      RESIDUAL-MATURITY BUCKET OF A DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the residual-maturity bucket of a date
'
' WHY THIS EXISTS
'   Positions and encumbrances are assigned to the maturity buckets of the
'   templates by one rule
'
' INPUTS
'   ReportingDate
'     The as-of date
'
'   EndDate
'     Maturity date, or the date until which an asset is encumbered
'
' RETURNS
'   BUCKET_LT6M when EndDate is before the reporting date plus 6 calendar
'   months, BUCKET_6M1Y when before plus 12 months, otherwise BUCKET_GE1Y
'
' BEHAVIOR
'   Compares EndDate with ReportingDate plus 6, then plus 12 months
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   NSFR_Core_Config: BUCKET_LT6M, BUCKET_6M1Y, BUCKET_GE1Y
'
' NOTES
'   Months are calendar months with the day clamped to the month's end
'   (DateAdd).
'   Reference: PROTOTYPE.md, Maturity and encumbrance
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Pick the first bucket whose limit EndDate falls before
        If EndDate < DateAdd("m", 6, ReportingDate) Then
            MaturityBucket = BUCKET_LT6M
        ElseIf EndDate < DateAdd("m", 12, ReportingDate) Then
            MaturityBucket = BUCKET_6M1Y
        Else
            MaturityBucket = BUCKET_GE1Y
        End If

End Function

Public Function BucketCode( _
    ByVal Bucket As Long) As String
'
'------------------------------------------------------------------------------
'                           DISPLAY CODE OF A BUCKET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the display code of a bucket: <6M, 6M-1Y, >=1Y or HQLA
'
' WHY THIS EXISTS
'   Outputs and messages show buckets by code rather than by number
'
' INPUTS
'   Bucket
'     A bucket number of NSFR_Core_Config
'
' RETURNS
'   The code; vbNullString for any other number
'
' BEHAVIOR
'   Maps each bucket constant to its code
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   NSFR_Core_Config: BUCKET_LT6M, BUCKET_6M1Y, BUCKET_GE1Y, BUCKET_HQLA
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Map the bucket to its code
        Select Case Bucket
            Case BUCKET_LT6M
                BucketCode = "<6M"
            Case BUCKET_6M1Y
                BucketCode = "6M-1Y"
            Case BUCKET_GE1Y
                BucketCode = ">=1Y"
            Case BUCKET_HQLA
                BucketCode = "HQLA"
            Case Else
                'No action for other values
        End Select

End Function

Public Function BucketFromCode( _
    ByVal Code As String) As Long
'
'------------------------------------------------------------------------------
'                          BUCKET OF AN OVERRIDE CODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the maturity bucket of an override code
'
' WHY THIS EXISTS
'   A position can name its bucket by code instead of by maturity date
'
' INPUTS
'   Code
'     <6M, 6M-1Y or >=1Y; case and spaces are ignored
'
' RETURNS
'   The bucket; 0 when the code is not <6M, 6M-1Y or >=1Y
'
' BEHAVIOR
'   Trims the code, removes its spaces, upper-cases it and maps it
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   NSFR_Core_Config: BUCKET_LT6M, BUCKET_6M1Y, BUCKET_GE1Y
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Map the normalized code to its bucket, any other code to 0
        Select Case UCase$(Replace(Trim$(Code), " ", vbNullString))
            Case "<6M"
                BucketFromCode = BUCKET_LT6M
            Case "6M-1Y"
                BucketFromCode = BUCKET_6M1Y
            Case ">=1Y"
                BucketFromCode = BUCKET_GE1Y
            Case Else
                BucketFromCode = 0
        End Select

End Function

Public Function HasBucket( _
    ByVal Cells As String, _
    ByVal Bucket As Long) As Boolean
'
'------------------------------------------------------------------------------
'                       TELL A BUCKET OF A TEMPLATE ROW
'------------------------------------------------------------------------------
' PURPOSE
'   True when the cell list of a template row (e.g. "123H") holds the
'   bucket
'
' WHY THIS EXISTS
'   Not every template row has a cell in every bucket
'
' INPUTS
'   Cells
'     The cell letters of the row, from BUCKET_LETTERS
'
'   Bucket
'     A bucket number from 1 to BUCKET_COUNT
'
' RETURNS
'   True when Cells holds the bucket's letter; False otherwise, also for a
'   bucket out of range
'
' BEHAVIOR
'   Looks up the bucket's letter in BUCKET_LETTERS and searches Cells for
'   it, case-sensitively
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   NSFR_Core_Config: BUCKET_COUNT, BUCKET_LETTERS
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a bucket out of range
        If Bucket < 1 Or Bucket > BUCKET_COUNT Then Exit Function

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Search the row's cells for the bucket's letter
        HasBucket = (InStr(1, Cells, Mid$(BUCKET_LETTERS, Bucket, 1), vbBinaryCompare) > 0)

End Function

Public Function IdDepth( _
    ByVal RowId As String) As Long
'
'------------------------------------------------------------------------------
'                              DEPTH OF A ROW ID
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of dots in a hierarchical row ID: 0 for "1", 3 for
'   "1.1.1.1"
'
' WHY THIS EXISTS
'   The depth of a row in the template hierarchy follows from its ID
'
' INPUTS
'   RowId
'     A hierarchical row ID
'
' RETURNS
'   The number of dots in RowId
'
' BEHAVIOR
'   Compares the length of RowId with and without its dots
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Count the dots
        IdDepth = Len(RowId) - Len(Replace(RowId, ".", vbNullString))

End Function

'
'------------------------------------------------------------------------------
'
'                                   LOOK-UPS
'
'------------------------------------------------------------------------------
'

Public Function FindRow2( _
    ByRef Table() As Variant, _
    ByVal Rows As Long, _
    ByVal Column1 As Long, _
    ByVal Value1 As String, _
    ByVal Column2 As Long, _
    ByVal Value2 As String) As Long
'
'------------------------------------------------------------------------------
'                            FIND A ROW BY TWO KEYS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first row of a 2-D table whose two key columns equal the
'   given texts
'
' WHY THIS EXISTS
'   Parameter tables are keyed by two columns, such as template and row ID
'
' INPUTS
'   Table
'     1-based 2-D array, read only
'
'   Rows
'     Rows to search, from row 1
'
'   Column1
'     Column of the first key
'
'   Value1
'     Text of the first key
'
'   Column2
'     Column of the second key
'
'   Value2
'     Text of the second key
'
' RETURNS
'   The row index, or 0
'
' BEHAVIOR
'   Compares the trimmed text of each key cell (TextOf) with the given
'   text, case-sensitively, and stops at the first match
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   TextOf
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
    Const PROC_NAME             As String = "NSFR_Core_Util.FindRow2"

    Dim RowIndex                As Long         'Row of Table

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Return the first row whose two keys match
        For RowIndex = 1 To Rows
            If TextOf(Table(RowIndex, Column1)) = Value1 Then
                If TextOf(Table(RowIndex, Column2)) = Value2 Then
                    FindRow2 = RowIndex
                    Exit Function
                End If
            End If
        Next RowIndex

End Function

Public Function FindRow1( _
    ByRef Table() As Variant, _
    ByVal Rows As Long, _
    ByVal Column1 As Long, _
    ByVal Value1 As String) As Long
'
'------------------------------------------------------------------------------
'                            FIND A ROW BY ONE KEY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first row of a 2-D table whose key column equals the text
'
' WHY THIS EXISTS
'   Look-ups and duplicate checks on tables keyed by one column
'
' INPUTS
'   Table
'     1-based 2-D array, read only
'
'   Rows
'     Rows to search, from row 1
'
'   Column1
'     Column of the key
'
'   Value1
'     Text of the key
'
' RETURNS
'   The row index, or 0
'
' BEHAVIOR
'   Compares the trimmed text of each key cell (TextOf) with the given
'   text, case-sensitively, and stops at the first match
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   TextOf
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
    Const PROC_NAME             As String = "NSFR_Core_Util.FindRow1"

    Dim RowIndex                As Long         'Row of Table

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Return the first row whose key matches
        For RowIndex = 1 To Rows
            If TextOf(Table(RowIndex, Column1)) = Value1 Then
                FindRow1 = RowIndex
                Exit Function
            End If
        Next RowIndex

End Function

'
'------------------------------------------------------------------------------
'
'                                    LISTS
'
'------------------------------------------------------------------------------
'

Public Sub ListAdd( _
    ByRef List() As Variant, _
    ByRef Count As Long, _
    ByVal Item As Variant)
'
'------------------------------------------------------------------------------
'                           APPEND AN ITEM TO A LIST
'------------------------------------------------------------------------------
' PURPOSE
'   Appends an item, usually a 1-based 1-D array holding one output row,
'   to a growable list
'
' WHY THIS EXISTS
'   Output rows are collected before their number is known
'
' INPUTS
'   List
'     1-based 1-D array, grown here by doubling
'
'   Count
'     Number of items held, incremented here
'
'   Item
'     The item to append
'
' RETURNS
'   Nothing; updates List and Count
'
' BEHAVIOR
'   Sizes an empty list to 16 items and doubles a full one, then stores
'   Item after the last one
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' GROW THE LIST
'------------------------------------------------------------------------------
    'Size an empty list, or double a full one
        If Count = 0 Then
            ReDim List(1 To 16)
        ElseIf Count >= UBound(List) Then
            ReDim Preserve List(1 To 2 * UBound(List))
        End If

'------------------------------------------------------------------------------
' APPEND THE ITEM
'------------------------------------------------------------------------------
    'Store the item after the last one
        Count = Count + 1
        List(Count) = Item

End Sub

Public Function ListToMatrix( _
    ByRef List() As Variant, _
    ByVal Count As Long, _
    ByVal Columns As Long) As Variant
'
'------------------------------------------------------------------------------
'                          COPY A LIST INTO A MATRIX
'------------------------------------------------------------------------------
' PURPOSE
'   Copies a list of row arrays into a 1-based 2-D array, ready for one
'   write to a range
'
' WHY THIS EXISTS
'   A range is written in one operation from a 2-D array
'
' INPUTS
'   List
'     The list of 1-based row arrays, read only
'
'   Count
'     Number of items held
'
'   Columns
'     Columns of the result; each row array has at least this many items
'
' RETURNS
'   Array(1 To Max(Count,1), 1 To Columns); an empty list gives one empty
'   row
'
' BEHAVIOR
'   Sizes the result, then copies each row array into its row
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   VBA library only
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
    Const PROC_NAME             As String = "NSFR_Core_Util.ListToMatrix"

    Dim Matrix()                As Variant      'Result
    Dim RowItem                 As Variant      'One list item
    Dim RowIndex                As Long         'Row of Matrix, item of List
    Dim ColumnIndex             As Long         'Column of Matrix

'------------------------------------------------------------------------------
' FILL THE MATRIX
'------------------------------------------------------------------------------
    'Give one empty row for an empty list, else copy each item
        If Count < 1 Then
            ReDim Matrix(1 To 1, 1 To Columns)
        Else
            ReDim Matrix(1 To Count, 1 To Columns)
            For RowIndex = 1 To Count
                RowItem = List(RowIndex)
                For ColumnIndex = 1 To Columns
                    Matrix(RowIndex, ColumnIndex) = RowItem(ColumnIndex)
                Next ColumnIndex
            Next RowIndex
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the matrix
        ListToMatrix = Matrix

End Function

Public Function NewRow( _
    ByVal Columns As Long) As Variant
'
'------------------------------------------------------------------------------
'                             CREATE AN EMPTY ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a 1-based 1-D array of Columns empty values
'
' WHY THIS EXISTS
'   Output rows are built field by field before they join a list
'
' INPUTS
'   Columns
'     Number of values
'
' RETURNS
'   Array(1 To Columns) of Empty values
'
' BEHAVIOR
'   Sizes a new array and returns it
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   VBA library only
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
    Const PROC_NAME             As String = "NSFR_Core_Util.NewRow"

    Dim RowItem()               As Variant      'Result

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Size the row and return it
        ReDim RowItem(1 To Columns)
        NewRow = RowItem

End Function

Public Sub AddDiagnostic( _
    ByRef List() As Variant, _
    ByRef Count As Long, _
    ByVal Severity As String, _
    ByVal SheetName As String, _
    ByVal TableRow As Long, _
    ByVal Reference As String, _
    ByVal Message As String)
'
'------------------------------------------------------------------------------
'                             APPEND A DIAGNOSTIC
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one diagnostic (severity, sheet, row, reference, message) to a
'   list
'
' WHY THIS EXISTS
'   Every module reports its findings as rows of one diagnostic table
'
' INPUTS
'   List
'     The diagnostic list, grown by ListAdd
'
'   Count
'     Number of diagnostics held, incremented here
'
'   Severity
'     Severity text of NSFR_Core_Config
'
'   SheetName
'     Sheet the message concerns
'
'   TableRow
'     Row of the input table (1 = first data row), or 0 when the message
'     concerns no single row
'
'   Reference
'     Field, code or key the message concerns
'
'   Message
'     The message text
'
' RETURNS
'   Nothing; updates List and Count
'
' BEHAVIOR
'   Builds a row of DG_COLUMNS values, leaves the row column empty when
'   TableRow is 0, and appends it
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   NewRow, ListAdd; NSFR_Core_Config: DG_COLUMNS and the DG_ columns
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
    Const PROC_NAME             As String = "NSFR_Core_Util.AddDiagnostic"

    Dim RowItem                 As Variant      'Diagnostic row

'------------------------------------------------------------------------------
' BUILD THE ROW
'------------------------------------------------------------------------------
    'Fill the diagnostic fields, leaving the row empty when there is none
        RowItem = NewRow(DG_COLUMNS)
        RowItem(DG_SEVERITY) = Severity
        RowItem(DG_SHEET) = SheetName
        If TableRow > 0 Then RowItem(DG_ROW) = TableRow
        RowItem(DG_REFERENCE) = Reference
        RowItem(DG_MESSAGE) = Message

'------------------------------------------------------------------------------
' APPEND THE ROW
'------------------------------------------------------------------------------
    'Add the row to the list
        ListAdd List, Count, RowItem

End Sub

Public Function CountSeverity( _
    ByRef List() As Variant, _
    ByVal Count As Long, _
    ByVal Severity As String) As Long
'
'------------------------------------------------------------------------------
'                      COUNT DIAGNOSTICS OF ONE SEVERITY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of diagnostics of one severity
'
' WHY THIS EXISTS
'   Run outcomes depend on how many errors and warnings were reported
'
' INPUTS
'   List
'     The diagnostic list, read only
'
'   Count
'     Number of diagnostics held
'
'   Severity
'     Severity text to count
'
' RETURNS
'   The number of diagnostics whose severity equals Severity
'
' BEHAVIOR
'   Compares the severity field of each diagnostic, case-sensitively
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   NSFR_Core_Config: DG_SEVERITY
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
    Const PROC_NAME             As String = "NSFR_Core_Util.CountSeverity"

    Dim RowItem                 As Variant      'One diagnostic
    Dim ItemIndex               As Long         'Item of List

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Count the diagnostics of the severity
        For ItemIndex = 1 To Count
            RowItem = List(ItemIndex)
            If RowItem(DG_SEVERITY) = Severity Then CountSeverity = CountSeverity + 1
        Next ItemIndex

End Function

'
'------------------------------------------------------------------------------
'
'                                  COMPARISON
'
'------------------------------------------------------------------------------
'

Public Function ToleranceResult( _
    ByVal Difference As Double, _
    ByVal BaseValue As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double) As String
'
'------------------------------------------------------------------------------
'                   TEST A DIFFERENCE AGAINST THE TOLERANCES
'------------------------------------------------------------------------------
' PURPOSE
'   Returns MATCH when a difference is within the absolute or the relative
'   tolerance, else DIFF
'
' WHY THIS EXISTS
'   The reconciliation compares the bank's figures with the benchmark by
'   one tolerance test
'
' INPUTS
'   Difference
'     The difference between the two figures
'
'   BaseValue
'     The figure the relative tolerance applies to
'
'   AbsTolerance
'     The absolute tolerance
'
'   RelTolerance
'     The relative tolerance
'
' RETURNS
'   RESULT_MATCH when |Difference| <= AbsTolerance, or when BaseValue is
'   not 0 and |Difference| / |BaseValue| <= RelTolerance; else RESULT_DIFF
'
' BEHAVIOR
'   Tests the absolute tolerance first, then the relative one when
'   BaseValue is not 0
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   NSFR_Core_Config: RESULT_MATCH, RESULT_DIFF
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Match within the absolute tolerance, else within the relative one
        If Abs(Difference) <= AbsTolerance Then
            ToleranceResult = RESULT_MATCH
        ElseIf BaseValue <> 0 Then
            If Abs(Difference) / Abs(BaseValue) <= RelTolerance Then
                ToleranceResult = RESULT_MATCH
            Else
                ToleranceResult = RESULT_DIFF
            End If
        Else
            ToleranceResult = RESULT_DIFF
        End If

End Function

Public Function FormatAmount( _
    ByVal Amount As Double) As String
'
'------------------------------------------------------------------------------
'                        FORMAT AN AMOUNT FOR MESSAGES
'------------------------------------------------------------------------------
' PURPOSE
'   Returns an amount as text with thousands separators and no decimals,
'   for messages
'
' WHY THIS EXISTS
'   Messages show amounts in one readable form
'
' INPUTS
'   Amount
'     The amount
'
' RETURNS
'   The amount formatted with "#,##0"
'
' BEHAVIOR
'   Formats the amount with Format$
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   None
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Format the amount
        FormatAmount = Format$(Amount, "#,##0")

End Function
