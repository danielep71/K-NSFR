Attribute VB_Name = "NSFR_Core_Import"
'==============================================================================
' MODULE: NSFR_Core_Import
'------------------------------------------------------------------------------
' PURPOSE
'   Turn a bank's NSFR figures into BankCells rows (template, row, column,
'   value), from two sources:
'
'   - a worksheet laid out as the reporting templates of Annex XII: a
'     title cell starting with "C 80.00", "C 81.00" or "C 84.00", one
'     header row of four-digit column codes ("0010") and one column of
'     four-digit row codes, as text;
'   - a CSV line "C 80.00 c0010 r0040,a,123.45" (RowColumn_Index, Suffix,
'     Amount), the export format of the bank's reporting workbook.
'
' PUBLIC SURFACE
'   None outside this VBA project. TemplateOfSheet, ScanTemplateSheet and
'   ParseCsvLine are Public for the host.
'
' DEPENDENCIES
'   NSFR_Core_Config, NSFR_Core_Util.
'
' STATE OWNERSHIP
'   None: the host reads the sheet or file and passes its values.
'
' ERROR POLICY
'   No On Error. A worksheet error value is kept as the text "ERROR" so
'   that the reconciliation reports it; a malformed CSV line is returned
'   with its reason.
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
        Private Const TITLE_SCAN        As Long = 8      'Rows and columns searched for the title
        Private Const HEADER_SCAN       As Long = 30     'Rows searched for the column codes

'------------------------------------------------------------------------------
'                                  WORKSHEETS
'------------------------------------------------------------------------------
Public Function TemplateOfSheet( _
    ByRef Values() As Variant) _
    As String
'
'   PURPOSE    Template code named in the title of a sheet.
'   INPUTS     Values: the sheet's cells from A1, 1-based 2-D, read only.
'   RETURNS    "C 80.00" to "C 84.00", or "" when no title cell (longer
'              than the code itself, e.g. "C 80.00 - NSFR ...") starts with
'              one of them.
'   UPDATED    2026-11-05
'
    Dim Text            As String      'Cell text
    Dim r               As Long
    Dim c               As Long

        For r = 1 To Application_Min(TITLE_SCAN, UBound(Values, 1))
            For c = 1 To Application_Min(TITLE_SCAN, UBound(Values, 2))
                Text = TextOf(Values(r, c))
                If Len(Text) > 7 Then          'a title, not a bare code in an index listing
                    If Left$(Text, 3) = "C 8" And Mid$(Text, 4, 1) >= "0" And Mid$(Text, 4, 1) <= "4" _
                       And Mid$(Text, 5, 3) = ".00" Then
                        TemplateOfSheet = Left$(Text, 7)
                        Exit Function
                    End If
                End If
            Next c
        Next r

End Function


Public Function ScanTemplateSheet( _
    ByRef Values() As Variant, _
    ByVal TemplateCode As String, _
    ByVal Source As String, _
    ByRef Cells() As Variant, _
    ByRef CellCount As Long, _
    ByRef Problem As String) _
    As Long
'
'   PURPOSE    Read every numeric (or error) cell at the crossing of a row
'              code and a column code of a template sheet.
'   INPUTS     Values: the sheet's cells from A1, read only. TemplateCode:
'              from TemplateOfSheet. Source: text kept with each row.
'   RETURNS    Number of cells appended to Cells, a list of BankCells rows
'              (BC_COLUMNS wide) holding CellCount items. Problem: why the
'              sheet could not be read; "" otherwise.
'   UPDATED    2026-11-05
'
    Dim Rows            As Long        'Rows of Values
    Dim Columns         As Long        'Columns of Values
    Dim HeaderRow       As Long        'Row of the column codes
    Dim FirstCodeColumn As Long        'First column holding a column code
    Dim CodeCount       As Long        'Column codes in a row
    Dim RowCodeColumn   As Long        'Column of the row codes
    Dim BestCount       As Long        'Row codes found in the best column
    Dim Found           As Long        'Row codes in a candidate column
    Dim RowCode         As String      'Row code of a data row
    Dim Value           As Variant     'Cell value
    Dim RowItem         As Variant     'BankCells row
    Dim r               As Long
    Dim c               As Long

        Problem = vbNullString
        Rows = UBound(Values, 1)
        Columns = UBound(Values, 2)
        For r = 1 To Application_Min(HEADER_SCAN, Rows)
            CodeCount = 0
            FirstCodeColumn = 0
            For c = 2 To Columns
                If IsCode4(Values(r, c)) Then
                    CodeCount = CodeCount + 1
                    If FirstCodeColumn = 0 Then FirstCodeColumn = c
                End If
            Next c
            If CodeCount >= 2 Then
                HeaderRow = r
                Exit For
            End If
        Next r
        If HeaderRow = 0 Then
            Problem = "no header row of column codes (0010, 0020, ...) in the first " & HEADER_SCAN & " rows"
            Exit Function
        End If
        For c = 1 To FirstCodeColumn - 1
            Found = 0
            For r = HeaderRow + 1 To Rows
                If IsCode4(Values(r, c)) Then Found = Found + 1
            Next r
            If Found > BestCount Then
                BestCount = Found
                RowCodeColumn = c
            End If
        Next c
        If RowCodeColumn = 0 Then
            Problem = "no column of row codes left of the first column code"
            Exit Function
        End If
        For r = HeaderRow + 1 To Rows
            If IsCode4(Values(r, RowCodeColumn)) Then
                RowCode = Trim$(Values(r, RowCodeColumn))
                For c = FirstCodeColumn To Columns
                    If IsCode4(Values(HeaderRow, c)) Then
                        Value = Values(r, c)
                        If IsError(Value) Then
                            Value = "ERROR"
                        ElseIf VarType(Value) = vbString Or IsEmpty(Value) Then
                            Value = Empty
                        End If
                        If Not IsEmpty(Value) Then
                            RowItem = NewRow(BC_COLUMNS)
                            RowItem(BC_TEMPLATE) = TemplateCode
                            RowItem(BC_ROW) = RowCode
                            RowItem(BC_COLUMN) = Trim$(Values(HeaderRow, c))
                            RowItem(BC_VALUE) = Value
                            RowItem(BC_SOURCE) = Source
                            ListAdd Cells, CellCount, RowItem
                            ScanTemplateSheet = ScanTemplateSheet + 1
                        End If
                    End If
                Next c
            End If
        Next r

End Function


Private Function Application_Min( _
    ByVal First As Long, _
    ByVal Second As Long) _
    As Long
'
'   PURPOSE    Smaller of two numbers.
'   UPDATED    2026-11-05
'
        If First < Second Then Application_Min = First Else Application_Min = Second

End Function

'------------------------------------------------------------------------------
'                                     CSV
'------------------------------------------------------------------------------
Public Function ParseCsvLine( _
    ByVal LineText As String, _
    ByRef TemplateCode As String, _
    ByRef RowCode As String, _
    ByRef ColumnCode As String, _
    ByRef Value As Double, _
    ByRef Problem As String) _
    As Boolean
'
'   PURPOSE    Parse one line "C 80.00 c0010 r0040,a,123.45" of the bank's
'              CSV export.
'   RETURNS    True with the four parts when the line is a data line.
'              False with Problem = "" for a blank or header line, and with
'              the reason otherwise. The amount uses a decimal point.
'   UPDATED    2026-11-05
'
    Dim Fields          As Variant     'Comma-separated fields
    Dim Tokens          As Variant     'Space-separated parts of the index
    Dim Token           As String      'One part
    Dim AmountText      As String      'Third field
    Dim Index           As String      'First field, spaces collapsed
    Dim i               As Long

        Problem = vbNullString
        TemplateCode = vbNullString
        RowCode = vbNullString
        ColumnCode = vbNullString
        LineText = Trim$(LineText)
        If Len(LineText) = 0 Then Exit Function
        If LCase$(Left$(LineText, 15)) = "rowcolumn_index" Then Exit Function
        Fields = Split(LineText, ",")
        If UBound(Fields) < 2 Then
            Problem = "fewer than three comma-separated fields"
            Exit Function
        End If
        Index = Trim$(Fields(0))
        Do While InStr(Index, "  ") > 0
            Index = Replace(Index, "  ", " ")
        Loop
        Tokens = Split(Index, " ")
        For i = LBound(Tokens) To UBound(Tokens)
            Token = Tokens(i)
            If Len(Token) = 5 And LCase$(Left$(Token, 1)) = "c" And IsAllDigits(Mid$(Token, 2)) Then
                ColumnCode = Mid$(Token, 2)
            ElseIf Len(Token) = 5 And LCase$(Left$(Token, 1)) = "r" And IsAllDigits(Mid$(Token, 2)) Then
                RowCode = Mid$(Token, 2)
            ElseIf Len(Token) = 5 And Mid$(Token, 3, 1) = "." Then
                TemplateCode = "C " & Token
            ElseIf Len(Token) = 6 And UCase$(Left$(Token, 1)) = "C" And Mid$(Token, 4, 1) = "." Then
                TemplateCode = "C " & Mid$(Token, 2)
            End If
        Next i
        If Len(TemplateCode) = 0 Or Len(RowCode) = 0 Or Len(ColumnCode) = 0 Then
            Problem = "index '" & Fields(0) & "' is not 'C nn.nn cNNNN rNNNN'"
            Exit Function
        End If
        AmountText = Trim$(Fields(UBound(Fields)))
        If Len(AmountText) = 0 Then
            Problem = "empty amount"
            Exit Function
        End If
        For i = 1 To Len(AmountText)
            If InStr("0123456789.-+Ee", Mid$(AmountText, i, 1)) = 0 Then
                Problem = "amount '" & AmountText & "' is not a number with a decimal point"
                Exit Function
            End If
        Next i
        Value = Val(AmountText)
        ParseCsvLine = True

End Function
