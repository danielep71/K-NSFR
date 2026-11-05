Attribute VB_Name = "NSFR_Core_Util"
'==============================================================================
' MODULE: NSFR_Core_Util
'------------------------------------------------------------------------------
' PURPOSE
'   Pure helpers of the prototype: value conversion of worksheet values,
'   residual-maturity buckets, table look-ups, growable lists and the
'   tolerance test of the reconciliation.
'
' PUBLIC SURFACE
'   None outside this VBA project (Option Private Module).
'
' DEPENDENCIES
'   NSFR_Core_Config for the bucket letters and the diagnostic columns.
'
' STATE OWNERSHIP
'   None: every procedure depends only on its arguments.
'
' ERROR POLICY
'   No On Error. A value that cannot be converted is reported through the
'   Valid output argument, never turned into zero.
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
'                              VALUE CONVERSION
'------------------------------------------------------------------------------
Public Function TextOf( _
    ByVal Value As Variant) _
    As String
'
'   PURPOSE    Trimmed text of a worksheet value; "" for Empty, Null or an
'              error value.
'   UPDATED    2026-11-05
'
        If IsError(Value) Or IsNull(Value) Or IsEmpty(Value) Then
            TextOf = vbNullString
        ElseIf IsArray(Value) Then
            TextOf = vbNullString
        Else
            TextOf = Trim$(CStr(Value))
        End If

End Function


Public Function IsBlankValue( _
    ByVal Value As Variant) _
    As Boolean
'
'   PURPOSE    True for Empty, Null or text that is empty once trimmed. An
'              error value is not blank: it is an unusable input.
'   UPDATED    2026-11-05
'
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
    ByRef Valid As Boolean) _
    As Double
'
'   PURPOSE    Numeric value of a worksheet value.
'   INPUTS     Value: a cell value. Only numeric types are accepted; text,
'              even if it looks numeric, is rejected so that the result
'              never depends on the locale.
'   RETURNS    The number; Valid = False (and 0) when Value is not numeric.
'   UPDATED    2026-11-05
'
        Valid = False
        NumberOf = 0
        If IsError(Value) Or IsArray(Value) Then Exit Function
        Select Case VarType(Value)
            Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte
                NumberOf = CDbl(Value)
                Valid = True
        End Select

End Function


Public Function DateOf( _
    ByVal Value As Variant, _
    ByRef Valid As Boolean) _
    As Date
'
'   PURPOSE    Date of a worksheet value: an Excel date, a positive date
'              serial or text in the form yyyy-mm-dd.
'   RETURNS    The date; Valid = False when Value is none of these.
'   UPDATED    2026-11-05
'
    Dim Text            As String      'Value as text
    Dim YearPart        As Long        'yyyy
    Dim MonthPart       As Long        'mm
    Dim DayPart         As Long        'dd

        Valid = False
        If IsError(Value) Or IsArray(Value) Then Exit Function
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
                If Len(Text) = 10 And Mid$(Text, 5, 1) = "-" And Mid$(Text, 8, 1) = "-" Then
                    If IsAllDigits(Left$(Text, 4)) And IsAllDigits(Mid$(Text, 6, 2)) _
                       And IsAllDigits(Right$(Text, 2)) Then
                        YearPart = CLng(Left$(Text, 4))
                        MonthPart = CLng(Mid$(Text, 6, 2))
                        DayPart = CLng(Right$(Text, 2))
                        If MonthPart >= 1 And MonthPart <= 12 And DayPart >= 1 And DayPart <= 31 Then
                            DateOf = DateSerial(YearPart, MonthPart, DayPart)
                            Valid = (Day(DateOf) = DayPart)
                        End If
                    End If
                End If
        End Select

End Function


Public Function IsAllDigits( _
    ByVal Text As String) _
    As Boolean
'
'   PURPOSE    True when Text is not empty and holds the digits 0-9 only.
'   UPDATED    2026-11-05
'
    Dim i               As Long

        If Len(Text) = 0 Then Exit Function
        For i = 1 To Len(Text)
            If Mid$(Text, i, 1) < "0" Or Mid$(Text, i, 1) > "9" Then Exit Function
        Next i
        IsAllDigits = True

End Function


Public Function IsCode4( _
    ByVal Value As Variant) _
    As Boolean
'
'   PURPOSE    True for a four-digit row or column code held as text, such
'              as "0010", the form the reporting templates use.
'   UPDATED    2026-11-05
'
        If VarType(Value) <> vbString Then Exit Function
        IsCode4 = (Len(Trim$(Value)) = 4 And IsAllDigits(Trim$(Value)))

End Function

'------------------------------------------------------------------------------
'                                   BUCKETS
'------------------------------------------------------------------------------
Public Function MaturityBucket( _
    ByVal ReportingDate As Date, _
    ByVal EndDate As Date) _
    As Long
'
'   PURPOSE    Residual-maturity bucket of a date.
'   INPUTS     ReportingDate: the as-of date. EndDate: maturity date, or
'              the date until which an asset is encumbered.
'   RETURNS    BUCKET_LT6M when EndDate is before the reporting date plus 6
'              calendar months, BUCKET_6M1Y when before plus 12 months,
'              otherwise BUCKET_GE1Y. Months are calendar months with the
'              day clamped to the month's end (DateAdd).
'   REFERENCE  PROTOTYPE.md, Maturity and encumbrance.
'   UPDATED    2026-11-05
'
        If EndDate < DateAdd("m", 6, ReportingDate) Then
            MaturityBucket = BUCKET_LT6M
        ElseIf EndDate < DateAdd("m", 12, ReportingDate) Then
            MaturityBucket = BUCKET_6M1Y
        Else
            MaturityBucket = BUCKET_GE1Y
        End If

End Function


Public Function BucketCode( _
    ByVal Bucket As Long) _
    As String
'
'   PURPOSE    Display code of a bucket: <6M, 6M-1Y, >=1Y or HQLA.
'   UPDATED    2026-11-05
'
        Select Case Bucket
            Case BUCKET_LT6M: BucketCode = "<6M"
            Case BUCKET_6M1Y: BucketCode = "6M-1Y"
            Case BUCKET_GE1Y: BucketCode = ">=1Y"
            Case BUCKET_HQLA: BucketCode = "HQLA"
        End Select

End Function


Public Function BucketFromCode( _
    ByVal Code As String) _
    As Long
'
'   PURPOSE    Maturity bucket of an override code; 0 when the code is not
'              <6M, 6M-1Y or >=1Y.
'   UPDATED    2026-11-05
'
        Select Case UCase$(Replace(Trim$(Code), " ", vbNullString))
            Case "<6M": BucketFromCode = BUCKET_LT6M
            Case "6M-1Y": BucketFromCode = BUCKET_6M1Y
            Case ">=1Y": BucketFromCode = BUCKET_GE1Y
            Case Else: BucketFromCode = 0
        End Select

End Function


Public Function HasBucket( _
    ByVal Cells As String, _
    ByVal Bucket As Long) _
    As Boolean
'
'   PURPOSE    True when the cell list of a template row (e.g. "123H")
'              holds the bucket.
'   UPDATED    2026-11-05
'
        If Bucket < 1 Or Bucket > BUCKET_COUNT Then Exit Function
        HasBucket = (InStr(1, Cells, Mid$(BUCKET_LETTERS, Bucket, 1), vbBinaryCompare) > 0)

End Function


Public Function IdDepth( _
    ByVal RowId As String) _
    As Long
'
'   PURPOSE    Number of dots in a hierarchical row ID: 0 for "1", 3 for
'              "1.1.1.1".
'   UPDATED    2026-11-05
'
        IdDepth = Len(RowId) - Len(Replace(RowId, ".", vbNullString))

End Function

'------------------------------------------------------------------------------
'                                  LOOK-UPS
'------------------------------------------------------------------------------
Public Function FindRow2( _
    ByRef Table() As Variant, _
    ByVal Rows As Long, _
    ByVal Column1 As Long, _
    ByVal Value1 As String, _
    ByVal Column2 As Long, _
    ByVal Value2 As String) _
    As Long
'
'   PURPOSE    First row of a 2-D table whose two key columns equal the
'              given texts (trimmed, case-sensitive).
'   INPUTS     Table: 1-based 2-D array, read only. Rows: rows to search.
'   RETURNS    The row index, or 0.
'   UPDATED    2026-11-05
'
    Dim r               As Long

        For r = 1 To Rows
            If TextOf(Table(r, Column1)) = Value1 Then
                If TextOf(Table(r, Column2)) = Value2 Then
                    FindRow2 = r
                    Exit Function
                End If
            End If
        Next r

End Function


Public Function FindRow1( _
    ByRef Table() As Variant, _
    ByVal Rows As Long, _
    ByVal Column1 As Long, _
    ByVal Value1 As String) _
    As Long
'
'   PURPOSE    First row of a 2-D table whose key column equals the text.
'   RETURNS    The row index, or 0.
'   UPDATED    2026-11-05
'
    Dim r               As Long

        For r = 1 To Rows
            If TextOf(Table(r, Column1)) = Value1 Then
                FindRow1 = r
                Exit Function
            End If
        Next r

End Function

'------------------------------------------------------------------------------
'                                    LISTS
'------------------------------------------------------------------------------
Public Sub ListAdd( _
    ByRef List() As Variant, _
    ByRef Count As Long, _
    ByVal Item As Variant)
'
'   PURPOSE    Append an item (usually a 1-based 1-D array holding one
'              output row) to a growable list.
'   INPUTS     List: 1-based 1-D array, grown here by doubling. Count:
'              number of items held, incremented here.
'   UPDATED    2026-11-05
'
        If Count = 0 Then
            ReDim List(1 To 16)
        ElseIf Count >= UBound(List) Then
            ReDim Preserve List(1 To 2 * UBound(List))
        End If
        Count = Count + 1
        List(Count) = Item

End Sub


Public Function ListToMatrix( _
    ByRef List() As Variant, _
    ByVal Count As Long, _
    ByVal Columns As Long) _
    As Variant
'
'   PURPOSE    Copy a list of row arrays into a 1-based 2-D array, ready for
'              one write to a range.
'   RETURNS    Array(1 To Max(Count,1), 1 To Columns); an empty list gives
'              one empty row.
'   UPDATED    2026-11-05
'
    Dim Matrix()        As Variant     'Result
    Dim RowItem         As Variant     'One list item
    Dim r               As Long
    Dim c               As Long

        If Count < 1 Then
            ReDim Matrix(1 To 1, 1 To Columns)
        Else
            ReDim Matrix(1 To Count, 1 To Columns)
            For r = 1 To Count
                RowItem = List(r)
                For c = 1 To Columns
                    Matrix(r, c) = RowItem(c)
                Next c
            Next r
        End If
        ListToMatrix = Matrix

End Function


Public Function NewRow( _
    ByVal Columns As Long) _
    As Variant
'
'   PURPOSE    A 1-based 1-D array of Columns empty values.
'   UPDATED    2026-11-05
'
    Dim RowItem()       As Variant     'Result

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
'   PURPOSE    Append one diagnostic (severity, sheet, row, reference,
'              message) to a list.
'   INPUTS     TableRow: row of the input table (1 = first data row), or 0
'              when the message concerns no single row.
'   UPDATED    2026-11-05
'
    Dim RowItem         As Variant     'Diagnostic row

        RowItem = NewRow(DG_COLUMNS)
        RowItem(DG_SEVERITY) = Severity
        RowItem(DG_SHEET) = SheetName
        If TableRow > 0 Then RowItem(DG_ROW) = TableRow
        RowItem(DG_REFERENCE) = Reference
        RowItem(DG_MESSAGE) = Message
        ListAdd List, Count, RowItem

End Sub


Public Function CountSeverity( _
    ByRef List() As Variant, _
    ByVal Count As Long, _
    ByVal Severity As String) _
    As Long
'
'   PURPOSE    Number of diagnostics of one severity.
'   UPDATED    2026-11-05
'
    Dim RowItem         As Variant     'One diagnostic
    Dim i               As Long

        For i = 1 To Count
            RowItem = List(i)
            If RowItem(DG_SEVERITY) = Severity Then CountSeverity = CountSeverity + 1
        Next i

End Function

'------------------------------------------------------------------------------
'                                 COMPARISON
'------------------------------------------------------------------------------
Public Function ToleranceResult( _
    ByVal Difference As Double, _
    ByVal BaseValue As Double, _
    ByVal AbsTolerance As Double, _
    ByVal RelTolerance As Double) _
    As String
'
'   PURPOSE    MATCH when |Difference| <= AbsTolerance, or when BaseValue is
'              not 0 and |Difference| / |BaseValue| <= RelTolerance; else
'              DIFF.
'   UPDATED    2026-11-05
'
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
    ByVal Amount As Double) _
    As String
'
'   PURPOSE    An amount as text with thousands separators and no decimals,
'              for messages.
'   UPDATED    2026-11-05
'
        FormatAmount = Format$(Amount, "#,##0")

End Function
