Attribute VB_Name = "NSFR_Core_Config"
'==============================================================================
' MODULE: NSFR_Core_Config
'------------------------------------------------------------------------------
' PURPOSE
'   Layout constants shared by the prototype's core, API and host: the
'   column order of every input and output table, the bucket codes, the
'   profile identifier and the status and severity texts. The tables are
'   1-based 2-D arrays whose columns follow these indexes; the template
'   builder (tools/build_template.py) writes the sheets in the same order.
'
' PUBLIC SURFACE
'   Constants only. Option Private Module keeps them inside the project.
'
' DEPENDENCIES
'   None.
'
' STATE OWNERSHIP
'   None.
'
' ERROR POLICY
'   Not applicable.
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
' PROFILE
'------------------------------------------------------------------------------
        Public Const PROFILE_ID          As String = "eu-crr-standard"  'Only profile the prototype runs
        Public Const TEMPLATE_RSF        As String = "C 80.00"          'Required stable funding
        Public Const TEMPLATE_ASF        As String = "C 81.00"          'Available stable funding
        Public Const TEMPLATE_SUMMARY    As String = "C 84.00"          'NSFR summary

'------------------------------------------------------------------------------
' BUCKETS
'------------------------------------------------------------------------------
    'Bucket 1 to 3 are residual-maturity buckets; bucket 4 is the HQLA
    'column of C 80.00. Cell letters "1", "2", "3" and "H" name them in the
    'Amount cells and Factor cells columns of the factor table.
        Public Const BUCKET_COUNT        As Long = 4
        Public Const BUCKET_LT6M         As Long = 1                    'Residual maturity < 6 months
        Public Const BUCKET_6M1Y         As Long = 2                    '6 months to < 1 year
        Public Const BUCKET_GE1Y         As Long = 3                    '1 year or more
        Public Const BUCKET_HQLA         As Long = 4                    'HQLA column (C 80.00 only)
        Public Const BUCKET_LETTERS      As String = "123H"             'Cell letter of each bucket

'------------------------------------------------------------------------------
' FACTOR TABLE (Params)
'------------------------------------------------------------------------------
        Public Const FC_TEMPLATE         As Long = 1      'Template code, e.g. C 80.00
        Public Const FC_ROW              As Long = 2      'Row code, e.g. 0040
        Public Const FC_ID               As Long = 3      'Hierarchical ID, e.g. 1.1.1.1
        Public Const FC_ITEM             As Long = 4      'Row label
        Public Const FC_KIND             As Long = 5      'AGG, LEAF, APPL or MEMO
        Public Const FC_AMOUNT_CELLS     As Long = 6      'Buckets with an amount cell, e.g. 123H
        Public Const FC_FACTOR_CELLS     As Long = 7      'Buckets with an applicable-factor cell
        Public Const FC_REG              As Long = 8      'Regulatory factor, bucket 1 (to 11)
        Public Const FC_APPLIED          As Long = 12     'Applied factor, bucket 1 (to 15)
        Public Const FC_REFERENCE        As Long = 16     'Source locator
        Public Const FC_CHILDREN         As Long = 17     'Row codes summed into an AGG row
        Public Const FC_COLUMNS          As Long = 17

        Public Const KIND_AGG            As String = "AGG"   'Sum of its children
        Public Const KIND_LEAF           As String = "LEAF"  'Regulatory factors
        Public Const KIND_APPL           As String = "APPL"  'Factors set by the competent authority
        Public Const KIND_MEMO           As String = "MEMO"  '"of which" row, outside the totals

'------------------------------------------------------------------------------
' CATEGORY TABLE (Params)
'------------------------------------------------------------------------------
        Public Const CC_CODE             As Long = 1      'Category code
        Public Const CC_DESCRIPTION      As Long = 2
        Public Const CC_TEMPLATE         As Long = 3      'C 80.00, C 81.00, or blank for DERIVATIVE
        Public Const CC_RULE             As Long = 4      'MATURITY, HQLA, FIXED1, FIXED3, DERIVATIVE
        Public Const CC_UNDATED          As Long = 5      'Bucket 1-3 of an undated position; blank = required
        Public Const CC_ROW_UNENCUMBERED As Long = 6      'Row ID when unencumbered or encumbered < 6 months
        Public Const CC_ROW_ENC_6M1Y     As Long = 7      'Row ID when encumbered 6 months to < 1 year
        Public Const CC_ROW_ENC_GE1Y     As Long = 8      'Row ID when encumbered 1 year or more
        Public Const CC_MEMO             As Long = 9      '"of which" row ID, or blank
        Public Const CC_REFERENCE        As Long = 10
        Public Const CC_COLUMNS          As Long = 10

        Public Const RULE_MATURITY       As String = "MATURITY"
        Public Const RULE_HQLA           As String = "HQLA"
        Public Const RULE_FIXED1         As String = "FIXED1"
        Public Const RULE_FIXED3         As String = "FIXED3"
        Public Const RULE_DERIVATIVE     As String = "DERIVATIVE"

        Public Const DRV_FV_ASSET        As String = "DRV_FV_ASSET"
        Public Const DRV_FV_LIAB         As String = "DRV_FV_LIAB"
        Public Const DRV_VM_RECEIVED     As String = "DRV_VM_RECEIVED"
        Public Const DRV_VM_POSTED       As String = "DRV_VM_POSTED"
        Public Const DRV_ROW_LIABILITIES As String = "1.7.1"   'C 80.00: 5% of derivative liabilities
        Public Const DRV_ROW_NET_ASSETS  As String = "1.7.2"   'C 80.00: NSFR derivative assets
        Public Const DRV_ROW_NET_LIABS   As String = "2.7"     'C 81.00: net derivative liabilities

'------------------------------------------------------------------------------
' SUMMARY MAP (Params)
'------------------------------------------------------------------------------
        Public Const SC_ROW              As Long = 1      'C 84.00 row code
        Public Const SC_ID               As Long = 2
        Public Const SC_ITEM             As Long = 3
        Public Const SC_MEASURE          As Long = 4      'RSF, ASF or NSFR
        Public Const SC_PLUS             As Long = 5      'Template|ID list, separated by ";"
        Public Const SC_MINUS            As Long = 6
        Public Const SC_COLUMNS          As Long = 6

'------------------------------------------------------------------------------
' POSITIONS (input)
'------------------------------------------------------------------------------
        Public Const PC_ID               As Long = 1
        Public Const PC_DESCRIPTION      As Long = 2
        Public Const PC_CATEGORY         As Long = 3
        Public Const PC_AMOUNT           As Long = 4      'Carrying amount, units of the reporting currency
        Public Const PC_CURRENCY         As Long = 5      'Informative
        Public Const PC_MATURITY         As Long = 6      'Date; blank = the category's undated bucket
        Public Const PC_ENCUMBERED       As Long = 7      'Date until which the asset is encumbered
        Public Const PC_BUCKET_OVERRIDE  As Long = 8      '<6M, 6M-1Y or >=1Y
        Public Const PC_FACTOR_OVERRIDE  As Long = 9      'Decimal 0-1; makes the run WHAT-IF
        Public Const PC_COMMENT          As Long = 10
        Public Const PC_COLUMNS          As Long = 10

    'PositionCalc output columns.
        Public Const PO_ID               As Long = 1
        Public Const PO_CATEGORY         As Long = 2
        Public Const PO_TEMPLATE         As Long = 3
        Public Const PO_ROW              As Long = 4
        Public Const PO_ROW_ID           As Long = 5
        Public Const PO_BUCKET           As Long = 6
        Public Const PO_ENCUMBRANCE      As Long = 7
        Public Const PO_AMOUNT           As Long = 8
        Public Const PO_FACTOR           As Long = 9
        Public Const PO_STABLE_FUNDING   As Long = 10
        Public Const PO_MEMO             As Long = 11
        Public Const PO_STATUS           As Long = 12
        Public Const PO_COLUMNS          As Long = 12

'------------------------------------------------------------------------------
' BANK CELLS (input) AND RECONCILIATION (output)
'------------------------------------------------------------------------------
        Public Const BC_TEMPLATE         As Long = 1
        Public Const BC_ROW              As Long = 2
        Public Const BC_COLUMN           As Long = 3
        Public Const BC_VALUE            As Long = 4
        Public Const BC_SOURCE           As Long = 5
        Public Const BC_COLUMNS          As Long = 5

        Public Const RL_CHECK            As Long = 1      'ARITHMETIC, FACTOR, AMOUNT, STABLE FUNDING
        Public Const RL_TEMPLATE         As Long = 2
        Public Const RL_ROW              As Long = 3
        Public Const RL_ID               As Long = 4
        Public Const RL_ITEM             As Long = 5
        Public Const RL_COLUMN           As Long = 6
        Public Const RL_BANK             As Long = 7      'Bank figure, scaled to units
        Public Const RL_OTHER            As Long = 8      'Expected or benchmark figure
        Public Const RL_DIFFERENCE       As Long = 9      'Other - bank
        Public Const RL_RELATIVE         As Long = 10     'Difference / |bank|
        Public Const RL_RESULT           As Long = 11     'MATCH, DIFF, MISSING, INFO
        Public Const RL_IMPACT           As Long = 12     'Stable-funding impact of a factor difference
        Public Const RL_COLUMNS          As Long = 12

        Public Const CHECK_ARITHMETIC    As String = "ARITHMETIC"
        Public Const CHECK_FACTOR        As String = "FACTOR"
        Public Const CHECK_AMOUNT        As String = "AMOUNT"
        Public Const CHECK_STABLE        As String = "STABLE FUNDING"

        Public Const RESULT_MATCH        As String = "MATCH"
        Public Const RESULT_DIFF         As String = "DIFF"
        Public Const RESULT_MISSING      As String = "MISSING"
        Public Const RESULT_INFO         As String = "INFO"

    'Headline rows (1-4) and columns (1-3) of the attribution block.
        Public Const HL_REPORTED         As Long = 1
        Public Const HL_REPERFORMED      As Long = 2
        Public Const HL_RESTATED         As Long = 3
        Public Const HL_BENCHMARK        As Long = 4
        Public Const HL_ASF              As Long = 1
        Public Const HL_RSF              As Long = 2
        Public Const HL_NSFR             As Long = 3

'------------------------------------------------------------------------------
' DIAGNOSTICS
'------------------------------------------------------------------------------
        Public Const DG_SEVERITY         As Long = 1
        Public Const DG_SHEET            As Long = 2
        Public Const DG_ROW              As Long = 3      'Row of the input table (1 = first data row)
        Public Const DG_REFERENCE        As Long = 4
        Public Const DG_MESSAGE          As Long = 5
        Public Const DG_COLUMNS          As Long = 5

        Public Const SEV_ERROR           As String = "ERROR"
        Public Const SEV_WARNING         As String = "WARNING"
        Public Const SEV_INFO            As String = "INFO"

        Public Const STATUS_VALID        As String = "VALID"
        Public Const STATUS_INVALID      As String = "INVALID"
        Public Const STATUS_STOPPED      As String = "STOPPED"
        Public Const CLASS_REFERENCE     As String = "REFERENCE"
        Public Const CLASS_WHATIF        As String = "WHAT-IF"
