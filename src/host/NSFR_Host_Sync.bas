Attribute VB_Name = "NSFR_Host_Sync"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: NSFR_Host_Sync
'------------------------------------------------------------------------------
' PURPOSE
'   Synchronizes the workbook's VBA components with the repository: imports
'   or exports exactly the components that src/manifest.txt lists
'
' WHY THIS EXISTS
'   The exported source files, not the workbook, are the reviewed source.
'   An import takes a backup copy of the workbook first, and neither
'   direction ever changes sheets, cells, formats, workbook names or a
'   component the manifest omits
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - SyncImport: backup, then replace each listed component
'     - SyncExport: write each listed component to its source file
'
' ERROR POLICY
'   Each public function returns a report whose last line is RESULT PASS or
'   RESULT FAIL; it stops at the first failure, names the step and Excel's
'   message, and keeps the backup
'
' DEPENDENCIES
'   Excel object model: ThisWorkbook.VBProject and SaveCopyAs; VBA file
'   statements. No reference and no CreateObject. Excel's "Trust access to
'   the VBA project object model" must be on while a synchronization runs
'
' NOTES
'   Windows 11, Microsoft 365 desktop Excel, VBA7, 64-bit.
'   Rules: docs/INSTALLATION.md (template and synchronization)
'
' UPDATED
'   2026-11-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const MANIFEST_FILE         As String = "src\manifest.txt"              'Manifest path in the clone
    Private Const SELF_NAME             As String = "NSFR_Host_Sync"                'This module, never replaced
    Private Const TYPE_STD_MODULE       As Long = 1                                 'vbext_ct_StdModule
    Private Const TYPE_CLASS_MODULE     As Long = 2                                 'vbext_ct_ClassModule
    Private Const TYPE_FORM             As Long = 3                                 'vbext_ct_MSForm
    Private Const SPARE_PREFIX          As String = "NsfrSyncOld"                   'Temporary name of a replaced component
    Private Const RESULT_PASS           As String = "RESULT PASS"                   'Last report line on success
    Private Const RESULT_FAIL           As String = "RESULT FAIL"                   'Last report line on failure

    'Error numbers - every synchronization failure
    Private Const ERR_SYNC              As Long = vbObjectError + 513               'A synchronization step failed


'
'------------------------------------------------------------------------------
'
'                             PUBLIC ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Function SyncImport( _
    ByVal RepoRoot As String) As String
'
'------------------------------------------------------------------------------
'                         IMPORT THE LISTED COMPONENTS
'------------------------------------------------------------------------------
' PURPOSE
'   Replaces each component the manifest lists with its source file, after a
'   backup copy of the workbook
'
' WHY THIS EXISTS
'   A workbook built from the template receives the reviewed source in one
'   step, and a failure leaves a backup to return to
'
' INPUTS
'   RepoRoot
'     The folder of a clean clone, for example "C:\path\to\K-NSFR"
'
' RETURNS
'   The report: BACKUP, one IMPORTED or KEPT line per component, then
'   RESULT PASS, or RESULT FAIL with the step and reason
'
' BEHAVIOR
'   Reads the manifest, checks access to the VBA project, saves a backup,
'   then imports each component in manifest order. This module itself is
'   kept, because it is running: import it by hand
'
' ERROR POLICY
'   Contained: a failure ends the report; the backup stays
'
' DEPENDENCIES
'   ReadManifest, RequireProjectAccess, BackupCopy, ImportComponent
'
' NOTES
'   Changes the VBA project of ThisWorkbook only
'
' UPDATED
'   2026-11-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Sync.SyncImport"

    Dim Names()                 As String       'Component names of the manifest
    Dim Paths()                 As String       'Their source files
    Dim Report                  As String       'Report lines so far
    Dim Stage                   As String       'Step under way, for a failure
    Dim ComponentIndex          As Long         'Index into Names
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain every failure in the report
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
    'Read the manifest
        Stage = "manifest"
        ReadManifest RepoRoot, Names, Paths
    'Fail early when Excel denies access to the VBA project
        Stage = "trust access"
        RequireProjectAccess
    'Save a backup copy before changing anything
        Stage = "backup"
        Report = "BACKUP " & BackupCopy() & vbCrLf

'------------------------------------------------------------------------------
' IMPORT COMPONENTS
'------------------------------------------------------------------------------
    'Replace each listed component, keeping the running module
        For ComponentIndex = LBound(Names) To UBound(Names)
            If VBA.StrComp(Names(ComponentIndex), SELF_NAME, vbTextCompare) = 0 Then
                Report = Report & "KEPT " & Names(ComponentIndex) & " (running; import it by hand)" & vbCrLf
            Else
                Stage = "import " & Names(ComponentIndex)
                ImportComponent Names(ComponentIndex), Paths(ComponentIndex)
                Report = Report & "IMPORTED " & Names(ComponentIndex) & vbCrLf
            End If
        Next ComponentIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Close the report with the pass line
        SyncImport = Report & RESULT_PASS

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error description
        ErrorDescription = Err.Description
    'Close the report with the failed step and its reason
        SyncImport = Report & RESULT_FAIL & " at " & Stage & ": " & ErrorDescription

End Function

Public Function SyncExport( _
    ByVal RepoRoot As String) As String
'
'------------------------------------------------------------------------------
'                         EXPORT THE LISTED COMPONENTS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes each component the manifest lists to its source file
'
' WHY THIS EXISTS
'   Git then shows what changed in the workbook
'
' INPUTS
'   RepoRoot
'     The folder of the clone to write to
'
' RETURNS
'   The report: one EXPORTED line per component, then RESULT PASS, or
'   RESULT FAIL with the step and reason
'
' BEHAVIOR
'   Reads the manifest, checks access to the VBA project, then exports each
'   component in manifest order
'
' ERROR POLICY
'   Contained: a failure ends the report. Files already written stay
'   written; Git restores them
'
' DEPENDENCIES
'   ReadManifest, RequireProjectAccess, FindComponent, ExportComponent
'
' NOTES
'   Reads the VBA project; changes only the listed files
'
' UPDATED
'   2026-11-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Sync.SyncExport"

    Dim Names()                 As String       'Component names of the manifest
    Dim Paths()                 As String       'Their source files
    Dim Component               As Object       'The component to export
    Dim Report                  As String       'Report lines so far
    Dim Stage                   As String       'Step under way, for a failure
    Dim ComponentIndex          As Long         'Index into Names
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain every failure in the report
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
    'Read the manifest
        Stage = "manifest"
        ReadManifest RepoRoot, Names, Paths
    'Fail early when Excel denies access to the VBA project
        Stage = "trust access"
        RequireProjectAccess

'------------------------------------------------------------------------------
' EXPORT COMPONENTS
'------------------------------------------------------------------------------
    'Export each listed component over its source file
        For ComponentIndex = LBound(Names) To UBound(Names)
            Stage = "export " & Names(ComponentIndex)
            Set Component = FindComponent(Names(ComponentIndex))
            If Component Is Nothing Then
                Err.Raise ERR_SYNC, , "the workbook has no component " & Names(ComponentIndex)
            End If
            ExportComponent Component, Paths(ComponentIndex)
            Report = Report & "EXPORTED " & Names(ComponentIndex) & vbCrLf
        Next ComponentIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Close the report with the pass line
        SyncExport = Report & RESULT_PASS

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error description
        ErrorDescription = Err.Description
    'Close the report with the failed step and its reason
        SyncExport = Report & RESULT_FAIL & " at " & Stage & ": " & ErrorDescription

End Function

'
'------------------------------------------------------------------------------
'
'                                   MANIFEST
'
'------------------------------------------------------------------------------
'

Private Sub ReadManifest( _
    ByVal RepoRoot As String, _
    ByRef Names() As String, _
    ByRef Paths() As String)
'
'------------------------------------------------------------------------------
'                           READ THE SOURCE MANIFEST
'------------------------------------------------------------------------------
' PURPOSE
'   Reads src/manifest.txt: one "name path" pair per line
'
' WHY THIS EXISTS
'   The manifest is the only list of components the synchronization touches
'
' INPUTS
'   RepoRoot
'     The clone's folder, with or without a final "\"
'
'   Names
'     Output: the component names
'
'   Paths
'     Output: the full Windows paths of their source files
'
' RETURNS
'   Nothing; fills Names and Paths
'
' BEHAVIOR
'   Skips empty lines and lines that start with #. Checks that each listed
'   file exists
'
' ERROR POLICY
'   Raises ERR_SYNC for a missing manifest or file, a line that is not a
'   pair, or an empty list
'
' DEPENDENCIES
'   VBA file statements
'
' NOTES
'   None
'
' UPDATED
'   2026-11-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Sync.ReadManifest"

    Dim Root                    As String       'RepoRoot without a final "\"
    Dim FilePath                As String       'Full path of the manifest
    Dim FileNo                  As Integer      'File handle
    Dim Text                    As String       'The manifest's bytes as text
    Dim Lines()                 As String       'The manifest's lines
    Dim Fields()                As String       'Name and path of one line
    Dim Entry                   As String       'One trimmed line
    Dim Count                   As Long         'Pairs read
    Dim LineIndex               As Long         'Index into Lines

'------------------------------------------------------------------------------
' LOCATE THE MANIFEST
'------------------------------------------------------------------------------
    'Drop a final backslash from the clone's folder
        Root = RepoRoot
        If VBA.Right$(Root, 1) = "\" Then Root = VBA.Left$(Root, VBA.Len(Root) - 1)
    'Require the manifest
        FilePath = Root & "\" & MANIFEST_FILE
        If VBA.Len(Dir$(FilePath, vbNormal)) = 0 Then
            Err.Raise ERR_SYNC, , "no " & MANIFEST_FILE & " in " & Root
        End If

'------------------------------------------------------------------------------
' READ THE FILE
'------------------------------------------------------------------------------
    'Read the whole file as text
        FileNo = FreeFile
        Open FilePath For Binary Access Read As #FileNo
        Text = Space$(LOF(FileNo))
        Get #FileNo, , Text
        Close #FileNo
    'Split it into lines and size the outputs for the worst case
        Lines = Split(Replace(Text, vbCr, vbNullString), vbLf)
        ReDim Names(0 To UBound(Lines))
        ReDim Paths(0 To UBound(Lines))

'------------------------------------------------------------------------------
' PARSE THE ENTRIES
'------------------------------------------------------------------------------
    'Keep each "name path" pair whose file exists
        For LineIndex = LBound(Lines) To UBound(Lines)
            Entry = VBA.Trim$(Lines(LineIndex))
            If VBA.Len(Entry) > 0 And VBA.Left$(Entry, 1) <> "#" Then
                Fields = Split(Entry, " ")
                If UBound(Fields) <> 1 Then
                    Err.Raise ERR_SYNC, , "manifest line " & (LineIndex + 1) & " is not 'name path'"
                End If
                Names(Count) = Fields(0)
                Paths(Count) = Root & "\" & Replace(Fields(1), "/", "\")
                If VBA.Len(Dir$(Paths(Count), vbNormal)) = 0 Then
                    Err.Raise ERR_SYNC, , "no file " & Fields(1)
                End If
                Count = Count + 1
            End If
        Next LineIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Reject an empty list
        If Count = 0 Then Err.Raise ERR_SYNC, , "the manifest lists no component"
    'Trim the outputs to the pairs read
        ReDim Preserve Names(0 To Count - 1)
        ReDim Preserve Paths(0 To Count - 1)

End Sub

'
'------------------------------------------------------------------------------
'
'                                 VBA PROJECT
'
'------------------------------------------------------------------------------
'

Private Sub RequireProjectAccess()

'
'------------------------------------------------------------------------------
'                      REQUIRE ACCESS TO THE VBA PROJECT
'------------------------------------------------------------------------------
' PURPOSE
'   Fails early, with the remedy, when Excel denies access to the VBA
'   project
'
' WHY THIS EXISTS
'   Without the trust setting every later step fails with a less helpful
'   message
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads the number of components, which needs access to the project
'
' ERROR POLICY
'   Raises ERR_SYNC when the trust setting is off
'
' DEPENDENCIES
'   ThisWorkbook.VBProject
'
' NOTES
'   None
'
' UPDATED
'   2026-11-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Sync.RequireProjectAccess"

    Dim Count                   As Long         'Components, read to probe access

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn a denied access into the remedy
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' PROBE ACCESS
'------------------------------------------------------------------------------
    'Read the project, which fails when access is denied
        Count = ThisWorkbook.VBProject.VBComponents.Count

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise the remedy
        Err.Raise ERR_SYNC, , "Excel does not allow access to the VBA project: turn on " & _
            "'Trust access to the VBA project object model' while the sync runs"

End Sub

Private Function BackupCopy() As String

'
'------------------------------------------------------------------------------
'                              SAVE A BACKUP COPY
'------------------------------------------------------------------------------
' PURPOSE
'   Saves a copy of the workbook beside it before an import
'
' WHY THIS EXISTS
'   An import replaces components; the copy is the way back
'
' INPUTS
'   None
'
' RETURNS
'   The full path of the copy, stamped with the local time
'
' BEHAVIOR
'   Names the copy <workbook>.backup-<yyyymmdd-hhnnss>, adding -2, -3 and so
'   on when that name exists, then saves it
'
' ERROR POLICY
'   Raises ERR_SYNC for a workbook never saved; SaveCopyAs errors propagate
'
' DEPENDENCIES
'   ThisWorkbook.SaveCopyAs
'
' NOTES
'   Writes one file; the open workbook does not change
'
' UPDATED
'   2026-11-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Sync.BackupCopy"

    Dim FullName                As String       'Name of the saved workbook
    Dim Dot                     As Long         'Position of its extension
    Dim Stem                    As String       'Path of the copy without its extension
    Dim Target                  As String       'Path of the copy
    Dim Suffix                  As Long         'Makes the path unique within one second

'------------------------------------------------------------------------------
' VALIDATE THE WORKBOOK
'------------------------------------------------------------------------------
    'Require a saved workbook
        If VBA.Len(ThisWorkbook.Path) = 0 Then
            Err.Raise ERR_SYNC, , "save the workbook as .xlsm before a sync"
        End If

'------------------------------------------------------------------------------
' NAME THE COPY
'------------------------------------------------------------------------------
    'Split the workbook name at its extension
        FullName = ThisWorkbook.Name
        Dot = VBA.InStrRev(FullName, ".")
        If Dot = 0 Then Dot = VBA.Len(FullName) + 1
    'Stamp the copy with the local time
        Stem = ThisWorkbook.Path & "\" & VBA.Left$(FullName, Dot - 1) & ".backup-" & Format$(Now, "yyyymmdd-hhnnss")
        Target = Stem & VBA.Mid$(FullName, Dot)
    'Make the name unique within the same second
        Suffix = 1
        Do While VBA.Len(Dir$(Target, vbNormal)) > 0
            Suffix = Suffix + 1
            Target = Stem & "-" & Suffix & VBA.Mid$(FullName, Dot)
        Loop

'------------------------------------------------------------------------------
' SAVE THE COPY
'------------------------------------------------------------------------------
    'Save the copy and return its path
        ThisWorkbook.SaveCopyAs Target
        BackupCopy = Target

End Function

Private Sub ImportComponent( _
    ByVal ComponentName As String, _
    ByVal SourcePath As String)
'
'------------------------------------------------------------------------------
'                            REPLACE ONE COMPONENT
'------------------------------------------------------------------------------
' PURPOSE
'   Replaces one component with its source file
'
' WHY THIS EXISTS
'   The VBE cannot import over an existing component of the same name
'
' INPUTS
'   ComponentName
'     The manifest's name
'
'   SourcePath
'     Its source file
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Renames and removes the existing component, imports the file, and checks
'   the imported name
'
' ERROR POLICY
'   Raises ERR_SYNC for a sheet or workbook module of that name, or a file
'   that declares another name; VBE errors propagate
'
' DEPENDENCIES
'   FindComponent, IsOwnedType, VBComponents.Import
'
' NOTES
'   Removes and adds that one component
'
' UPDATED
'   2026-11-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Sync.ImportComponent"

    Dim Components              As Object       'The project's components
    Dim Existing                As Object       'The component to replace
    Dim Imported                As Object       'The component just imported
    Dim Spare                   As Long         'Number of a free temporary name

'------------------------------------------------------------------------------
' REMOVE THE EXISTING COMPONENT
'------------------------------------------------------------------------------
    'Find the component to replace
        Set Components = ThisWorkbook.VBProject.VBComponents
        Set Existing = FindComponent(ComponentName)
        If Not Existing Is Nothing Then
            If Not IsOwnedType(Existing.Type) Then
                Err.Raise ERR_SYNC, , ComponentName & " is a sheet or workbook module, which the sync never replaces"
            End If
            'Free the name first: a removed component can linger until the code ends.
            'The temporary name is short, because a module name has at most 31 characters.
            Spare = 1
            Do While Not FindComponent(SPARE_PREFIX & Spare) Is Nothing
                Spare = Spare + 1
            Loop
            Existing.Name = SPARE_PREFIX & Spare
            Components.Remove Existing
        End If

'------------------------------------------------------------------------------
' IMPORT THE SOURCE FILE
'------------------------------------------------------------------------------
    'Import the file and check the name it declares
        Set Imported = Components.Import(SourcePath)
        If VBA.StrComp(Imported.Name, ComponentName, vbBinaryCompare) <> 0 Then
            Err.Raise ERR_SYNC, , SourcePath & " declares " & Imported.Name & ", not " & ComponentName
        End If

End Sub

Private Sub ExportComponent( _
    ByVal Component As Object, _
    ByVal TargetPath As String)
'
'------------------------------------------------------------------------------
'                             EXPORT ONE COMPONENT
'------------------------------------------------------------------------------
' PURPOSE
'   Exports one component over its source file
'
' WHY THIS EXISTS
'   The VBE refuses an existing destination, so the export goes to a new
'   file first, which replaces the source only once it is written
'
' INPUTS
'   Component
'     The component
'
'   TargetPath
'     Its source file
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exports to TargetPath & ".export", deletes the source, then renames the
'   new file
'
' ERROR POLICY
'   VBE and file errors propagate; the source stays untouched until the
'   export has succeeded
'
' DEPENDENCIES
'   VBComponent.Export, Kill, Name
'
' NOTES
'   Writes TargetPath and, briefly, TargetPath & ".export"
'
' UPDATED
'   2026-11-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Sync.ExportComponent"

    Dim Fresh                   As String       'The new file

'------------------------------------------------------------------------------
' EXPORT
'------------------------------------------------------------------------------
    'Clear a leftover new file
        Fresh = TargetPath & ".export"
        If VBA.Len(Dir$(Fresh, vbNormal)) > 0 Then Kill Fresh
    'Export, then replace the source
        Component.Export Fresh
        Kill TargetPath
        Name Fresh As TargetPath

End Sub

Private Function FindComponent( _
    ByVal ComponentName As String) As Object
'
'------------------------------------------------------------------------------
'                           FIND A COMPONENT BY NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Finds a component of the project by name
'
' RETURNS
'   The component, or Nothing
'
' UPDATED
'   2026-11-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "NSFR_Host_Sync.FindComponent"

    Dim Component               As Object       'One component of the project

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Return the first component with that name, ignoring case
        For Each Component In ThisWorkbook.VBProject.VBComponents
            If VBA.StrComp(Component.Name, ComponentName, vbTextCompare) = 0 Then
                Set FindComponent = Component
                Exit Function
            End If
        Next Component

End Function

Private Function IsOwnedType( _
    ByVal ComponentType As Long) As Boolean
'
'------------------------------------------------------------------------------
'                      TELL A REPLACEABLE COMPONENT TYPE
'------------------------------------------------------------------------------
' PURPOSE
'   True for the component types the sync may replace: standard modules,
'   classes and forms, never sheet or workbook modules
'
' UPDATED
'   2026-11-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Accept standard modules, classes and forms
        IsOwnedType = (ComponentType = TYPE_STD_MODULE Or ComponentType = TYPE_CLASS_MODULE _
                       Or ComponentType = TYPE_FORM)

End Function
