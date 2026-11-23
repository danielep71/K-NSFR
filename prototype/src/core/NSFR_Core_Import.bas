Attribute VB_Name = "NSFR_Core_Import"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: NSFR_Core_Import
'------------------------------------------------------------------------------
' PURPOSE
'   Turns a bank's NSFR figures into BankCells rows (template, row, column,
'   value), from two sources:
'     - a worksheet laid out as the reporting templates of Annex XII: a
'       title cell starting with "C 80.00", "C 81.00" or "C 84.00", one
'       header row of four-digit column codes ("0010") and one column of
'       four-digit row codes, as text
'     - a CSV line "C 80.00 c0010 r0040,a,123.45" (RowColumn_Index, Suffix,
'       Amount), the export format of the bank's reporting workbook
'
' WHY THIS EXISTS
'   The reconciliation reads the bank's figures as BankCells rows, whichever
'   of the two sources they come from
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - TemplateOfSheet: the template code named in a sheet's title
'     - ScanTemplateSheet: the BankCells rows of a template sheet
'     - ParseCsvLine: the parts of one line of the CSV export
'
' ERROR POLICY
'   No On Error. A worksheet error value is kept as the text "ERROR" so
'   that the reconciliation reports it; a malformed CSV line is returned
'   with its reason
'
' DEPENDENCIES
'   NSFR_Core_Config, NSFR_Core_Util
'
' NOTES
'   Excel VBA7, 64-bit; no references beyond the defaults
'
' PUBLIC SURFACE
'   None outside this VBA project. TemplateOfSheet, ScanTemplateSheet and
'   ParseCsvLine are Public for the host
'
' STATE OWNERSHIP
'   None: the host reads the sheet or file and passes its values
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const TITLE_SCAN            As Long = 8         'Rows and columns searched for the title
    Private Const HEADER_SCAN           As Long = 30        'Rows searched for the column codes


'
'------------------------------------------------------------------------------
'
'                                  WORKSHEETS
'
'------------------------------------------------------------------------------
'

Public Function TemplateOfSheet( _
    ByRef Values() As Variant) As String
'
'------------------------------------------------------------------------------
'                           TEMPLATE CODE OF A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the template code named in the title of a sheet
'
' WHY THIS EXISTS
'   A bank workbook holds several sheets; only those titled with a template
'   code are read
'
' INPUTS
'   Values
'     The sheet's cells from A1, 1-based 2-D, read only
'
' RETURNS
'   "C 80.00" to "C 84.00", or "" when no title cell (longer than the code
'   itself, e.g. "C 80.00 - NSFR ...") starts with one of them
'
' BEHAVIOR
'   Searches the first TITLE_SCAN rows and columns, row by row, and returns
'   the first seven characters of the first cell of more than seven
'   characters that starts with "C 80.00" to "C 84.00"
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   TextOf, Application_Min
'
' NOTES
'   A bare code, as in an index listing, is not a title
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Core_Import.TemplateOfSheet"

    Dim Text                    As String       'Cell text
    Dim RowIndex                As Long         'Row of Values
    Dim ColumnIndex             As Long         'Column of Values

'------------------------------------------------------------------------------
' SEARCH THE TITLE
'------------------------------------------------------------------------------
    'Return the code of the first title cell in the scanned corner
        For RowIndex = 1 To Application_Min(TITLE_SCAN, UBound(Values, 1))
            For ColumnIndex = 1 To Application_Min(TITLE_SCAN, UBound(Values, 2))
                Text = TextOf(Values(RowIndex, ColumnIndex))
                'Require a title, not a bare code in an index listing
                If Len(Text) > 7 Then
                    If Left$(Text, 3) = "C 8" And Mid$(Text, 4, 1) >= "0" _
                        And Mid$(Text, 4, 1) <= "4" And Mid$(Text, 5, 3) = ".00" Then
                        TemplateOfSheet = Left$(Text, 7)
                        Exit Function
                    End If
                End If
            Next ColumnIndex
        Next RowIndex

End Function

Public Function ScanTemplateSheet( _
    ByRef Values() As Variant, _
    ByVal TemplateCode As String, _
    ByVal Source As String, _
    ByRef Cells() As Variant, _
    ByRef CellCount As Long, _
    ByRef Problem As String) As Long
'
'------------------------------------------------------------------------------
'                      READ THE CELLS OF A TEMPLATE SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Reads every numeric (or error) cell at the crossing of a row code and a
'   column code of a template sheet
'
' WHY THIS EXISTS
'   The bank's figures are read by their template codes, not by their
'   position on the sheet
'
' INPUTS
'   Values
'     The sheet's cells from A1, 1-based 2-D, read only
'
'   TemplateCode
'     The sheet's template, from TemplateOfSheet
'
'   Source
'     Text kept with each row
'
'   Cells
'     The BankCells list (BC_COLUMNS wide), appended to here
'
'   CellCount
'     Number of items in Cells, incremented here
'
'   Problem
'     Output: why the sheet could not be read; "" otherwise
'
' RETURNS
'   Number of cells appended to Cells
'
' BEHAVIOR
'   Steps:
'     - the header row is the first of the first HEADER_SCAN rows with at
'       least two column codes from column 2 on
'     - the row-code column is the column left of the first column code
'       holding the most row codes below the header
'     - each cell at a row code and a column code is kept: an error value
'       as the text "ERROR", any other value that is neither text nor
'       empty as it is; text and empty cells are skipped
'
' ERROR POLICY
'   No On Error; a sheet without a header row or a row-code column returns
'   0 with Problem set
'
' DEPENDENCIES
'   IsCode4, NewRow, ListAdd, Application_Min; NSFR_Core_Config: BC_COLUMNS
'   and the BC_ columns
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
    Const PROC_NAME             As String = "NSFR_Core_Import.ScanTemplateSheet"

    Dim Rows                    As Long         'Rows of Values
    Dim Columns                 As Long         'Columns of Values
    Dim HeaderRow               As Long         'Row of the column codes
    Dim FirstCodeColumn         As Long         'First column holding a column code
    Dim CodeCount               As Long         'Column codes in a row
    Dim RowCodeColumn           As Long         'Column of the row codes
    Dim BestCount               As Long         'Row codes found in the best column
    Dim Found                   As Long         'Row codes in a candidate column
    Dim RowCode                 As String       'Row code of a data row
    Dim Value                   As Variant      'Cell value
    Dim RowItem                 As Variant      'BankCells row
    Dim RowIndex                As Long         'Row of Values
    Dim ColumnIndex             As Long         'Column of Values

'------------------------------------------------------------------------------
' MEASURE THE SHEET
'------------------------------------------------------------------------------
    'Clear the problem and take the sheet's size
        Problem = vbNullString
        Rows = UBound(Values, 1)
        Columns = UBound(Values, 2)

'------------------------------------------------------------------------------
' FIND THE HEADER ROW
'------------------------------------------------------------------------------
    'Take the first row with at least two column codes
        For RowIndex = 1 To Application_Min(HEADER_SCAN, Rows)
            CodeCount = 0
            FirstCodeColumn = 0
            For ColumnIndex = 2 To Columns
                If IsCode4(Values(RowIndex, ColumnIndex)) Then
                    CodeCount = CodeCount + 1
                    If FirstCodeColumn = 0 Then FirstCodeColumn = ColumnIndex
                End If
            Next ColumnIndex
            If CodeCount >= 2 Then
                HeaderRow = RowIndex
                Exit For
            End If
        Next RowIndex
    'Give up without a header row
        If HeaderRow = 0 Then
            Problem = "no header row of column codes (0010, 0020, ...) in the first " & HEADER_SCAN & " rows"
            Exit Function
        End If

'------------------------------------------------------------------------------
' FIND THE ROW-CODE COLUMN
'------------------------------------------------------------------------------
    'Take the column left of the column codes with the most row codes
        For ColumnIndex = 1 To FirstCodeColumn - 1
            Found = 0
            For RowIndex = HeaderRow + 1 To Rows
                If IsCode4(Values(RowIndex, ColumnIndex)) Then Found = Found + 1
            Next RowIndex
            If Found > BestCount Then
                BestCount = Found
                RowCodeColumn = ColumnIndex
            End If
        Next ColumnIndex
    'Give up without a row-code column
        If RowCodeColumn = 0 Then
            Problem = "no column of row codes left of the first column code"
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ THE CELLS
'------------------------------------------------------------------------------
    'Append a BankCells row for each number or error at a row and a column code
        For RowIndex = HeaderRow + 1 To Rows
            If IsCode4(Values(RowIndex, RowCodeColumn)) Then
                RowCode = Trim$(Values(RowIndex, RowCodeColumn))
                For ColumnIndex = FirstCodeColumn To Columns
                    If IsCode4(Values(HeaderRow, ColumnIndex)) Then
                        'Keep an error as "ERROR"; skip text and empty cells
                        Value = Values(RowIndex, ColumnIndex)
                        If IsError(Value) Then
                            Value = "ERROR"
                        ElseIf VarType(Value) = vbString Or IsEmpty(Value) Then
                            Value = Empty
                        End If
                        If Not IsEmpty(Value) Then
                            RowItem = NewRow(BC_COLUMNS)
                            RowItem(BC_TEMPLATE) = TemplateCode
                            RowItem(BC_ROW) = RowCode
                            RowItem(BC_COLUMN) = Trim$(Values(HeaderRow, ColumnIndex))
                            RowItem(BC_VALUE) = Value
                            RowItem(BC_SOURCE) = Source
                            ListAdd Cells, CellCount, RowItem
                            ScanTemplateSheet = ScanTemplateSheet + 1
                        End If
                    End If
                Next ColumnIndex
            End If
        Next RowIndex

End Function

Private Function Application_Min( _
    ByVal First As Long, _
    ByVal Second As Long) As Long
'
'------------------------------------------------------------------------------
'                            SMALLER OF TWO NUMBERS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the smaller of two numbers
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the smaller number
        If First < Second Then Application_Min = First Else Application_Min = Second

End Function

'
'------------------------------------------------------------------------------
'
'                                     CSV
'
'------------------------------------------------------------------------------
'

Public Function ParseCsvLine( _
    ByVal LineText As String, _
    ByRef TemplateCode As String, _
    ByRef RowCode As String, _
    ByRef ColumnCode As String, _
    ByRef Value As Double, _
    ByRef Problem As String) As Boolean
'
'------------------------------------------------------------------------------
'                       PARSE ONE LINE OF THE CSV EXPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Parses one line "C 80.00 c0010 r0040,a,123.45" of the bank's CSV export
'
' WHY THIS EXISTS
'   The bank's reporting workbook exports its figures in this format
'
' INPUTS
'   LineText
'     One line of the file
'
'   TemplateCode
'     Output: the template, for example "C 80.00"
'
'   RowCode
'     Output: the four-digit row code
'
'   ColumnCode
'     Output: the four-digit column code
'
'   Value
'     Output: the amount
'
'   Problem
'     Output: why a non-blank line is not a data line; "" otherwise
'
' RETURNS
'   True with the four parts when the line is a data line. False with
'   Problem = "" for a blank or header line, and with the reason otherwise
'
' BEHAVIOR
'   Skips a blank line and the header line (starting with RowColumn_Index).
'   Splits the line at commas: the first field is the index, the last the
'   amount. In the index, with spaces collapsed:
'     - cNNNN is the column code and rNNNN the row code
'     - nn.nn or Cnn.nn is the template
'   The amount may hold only digits and . - + E e, and is read with Val
'
' ERROR POLICY
'   No On Error; a malformed line returns False with its reason
'
' DEPENDENCIES
'   IsAllDigits
'
' NOTES
'   The amount uses a decimal point
'
' UPDATED
'   2026-11-05
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Core_Import.ParseCsvLine"

    Dim Fields                  As Variant      'Comma-separated fields
    Dim Tokens                  As Variant      'Space-separated parts of the index
    Dim Token                   As String       'One part
    Dim AmountText              As String       'Third field
    Dim Index                   As String       'First field, spaces collapsed
    Dim LoopIndex               As Long         'Index into Tokens, then position in AmountText

'------------------------------------------------------------------------------
' RESET OUTPUTS
'------------------------------------------------------------------------------
    'Clear the problem and the codes
        Problem = vbNullString
        TemplateCode = vbNullString
        RowCode = vbNullString
        ColumnCode = vbNullString

'------------------------------------------------------------------------------
' SKIP BLANK AND HEADER LINES
'------------------------------------------------------------------------------
    'Return False without a problem for a blank line or the header line
        LineText = Trim$(LineText)
        If Len(LineText) = 0 Then Exit Function
        If LCase$(Left$(LineText, 15)) = "rowcolumn_index" Then Exit Function

'------------------------------------------------------------------------------
' SPLIT THE FIELDS
'------------------------------------------------------------------------------
    'Require at least three fields
        Fields = Split(LineText, ",")
        If UBound(Fields) < 2 Then
            Problem = "fewer than three comma-separated fields"
            Exit Function
        End If

'------------------------------------------------------------------------------
' PARSE THE INDEX
'------------------------------------------------------------------------------
    'Collapse repeated spaces in the index
        Index = Trim$(Fields(0))
        Do While InStr(1, Index, "  ", vbBinaryCompare) > 0
            Index = Replace(Index, "  ", " ")
        Loop
    'Recognize the column code, the row code and the template
        Tokens = Split(Index, " ")
        For LoopIndex = LBound(Tokens) To UBound(Tokens)
            Token = Tokens(LoopIndex)
            If Len(Token) = 5 And LCase$(Left$(Token, 1)) = "c" _
                And IsAllDigits(Mid$(Token, 2)) Then
                ColumnCode = Mid$(Token, 2)
            ElseIf Len(Token) = 5 And LCase$(Left$(Token, 1)) = "r" _
                And IsAllDigits(Mid$(Token, 2)) Then
                RowCode = Mid$(Token, 2)
            ElseIf Len(Token) = 5 And Mid$(Token, 3, 1) = "." Then
                TemplateCode = "C " & Token
            ElseIf Len(Token) = 6 And UCase$(Left$(Token, 1)) = "C" _
                And Mid$(Token, 4, 1) = "." Then
                TemplateCode = "C " & Mid$(Token, 2)
            End If
        Next LoopIndex
    'Require all three parts
        If Len(TemplateCode) = 0 Or Len(RowCode) = 0 Or Len(ColumnCode) = 0 Then
            Problem = "index '" & Fields(0) & "' is not 'C nn.nn cNNNN rNNNN'"
            Exit Function
        End If

'------------------------------------------------------------------------------
' PARSE THE AMOUNT
'------------------------------------------------------------------------------
    'Require an amount in the last field
        AmountText = Trim$(Fields(UBound(Fields)))
        If Len(AmountText) = 0 Then
            Problem = "empty amount"
            Exit Function
        End If
    'Accept only the characters of a number with a decimal point
        For LoopIndex = 1 To Len(AmountText)
            If InStr(1, "0123456789.-+Ee", Mid$(AmountText, LoopIndex, 1), _
                vbBinaryCompare) = 0 Then
                Problem = "amount '" & AmountText & "' is not a number with a decimal point"
                Exit Function
            End If
        Next LoopIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Read the amount and report a data line
        Value = Val(AmountText)
        ParseCsvLine = True

End Function
