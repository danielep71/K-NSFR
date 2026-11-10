Attribute VB_Name = "NSFR_Host_Sync"
'==============================================================================
' MODULE        NSFR_Host_Sync
' PURPOSE       Synchronize the workbook's VBA components with the repository:
'               import or export exactly the components that the manifest
'               src/manifest.txt lists, after a backup copy of the workbook
'               for an import. It never changes sheets, cells, formats or
'               workbook names, and never a component the manifest omits.
' PUBLIC        SyncImport, SyncExport
' DEPENDENCIES  Excel object model: ThisWorkbook.VBProject and SaveCopyAs;
'               VBA file statements. No reference and no CreateObject.
'               Excel's "Trust access to the VBA project object model" must
'               be on while a synchronization runs.
' STATE         None at module level.
' ERRORS        Each public function returns a report whose last line is
'               RESULT PASS or RESULT FAIL; it stops at the first failure,
'               names the step and Excel's message, and keeps the backup.
' COMPATIBILITY Windows 11, Microsoft 365 desktop Excel, VBA7, 64-bit.
' RULES         docs/INSTALLATION.md (template and synchronization), ADR-008.
' UPDATED       2026-11-10
' AUTHOR        danielep71
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
Private Const MANIFEST_FILE As String = "src\manifest.txt"
Private Const SELF_NAME As String = "NSFR_Host_Sync"
Private Const TYPE_STD_MODULE As Long = 1       'vbext_ct_StdModule
Private Const TYPE_CLASS_MODULE As Long = 2     'vbext_ct_ClassModule
Private Const TYPE_FORM As Long = 3             'vbext_ct_MSForm
Private Const ERR_SYNC As Long = vbObjectError + 513
Private Const SPARE_PREFIX As String = "NsfrSyncOld"  'Temporary name of a replaced component
Private Const RESULT_PASS As String = "RESULT PASS"
Private Const RESULT_FAIL As String = "RESULT FAIL"


'------------------------------------------------------------------------------
'                              PUBLIC ENTRY POINTS
'------------------------------------------------------------------------------

Public Function SyncImport( _
    ByVal RepoRoot As String) _
    As String
'PURPOSE        Replace each component the manifest lists with its source
'               file, after a backup copy of the workbook. This module itself
'               is kept, because it is running: import it by hand.
'INPUTS         RepoRoot: the folder of a clean clone, for example
'               "C:\path\to\K-NSFR".
'RETURNS        The report: BACKUP, one IMPORTED or KEPT line per component,
'               then RESULT PASS, or RESULT FAIL with the step and reason.
'ERROR POLICY   Contained: a failure ends the report; the backup stays.
'STATE OWNERSHIP Changes the VBA project of ThisWorkbook only.
'UPDATED        2026-11-10
    Dim Names()     As String                   'Component names of the manifest
    Dim Paths()     As String                   'Their source files
    Dim Report      As String                   'Report lines so far
    Dim Stage       As String                   'Step under way, for a failure
    Dim i           As Long                     'Index into Names

    On Error GoTo Failed
        Stage = "manifest"
        ReadManifest RepoRoot, Names, Paths
        Stage = "trust access"
        RequireProjectAccess
        Stage = "backup"
        Report = "BACKUP " & BackupCopy() & vbCrLf
        For i = LBound(Names) To UBound(Names)
            If StrComp(Names(i), SELF_NAME, vbTextCompare) = 0 Then
                Report = Report & "KEPT " & Names(i) & " (running; import it by hand)" & vbCrLf
            Else
                Stage = "import " & Names(i)
                ImportComponent Names(i), Paths(i)
                Report = Report & "IMPORTED " & Names(i) & vbCrLf
            End If
        Next i
        SyncImport = Report & RESULT_PASS
        Exit Function
Failed:
    SyncImport = Report & RESULT_FAIL & " at " & Stage & ": " & Err.Description
End Function


Public Function SyncExport( _
    ByVal RepoRoot As String) _
    As String
'PURPOSE        Write each component the manifest lists to its source file,
'               so that Git shows what changed in the workbook.
'INPUTS         RepoRoot: the folder of the clone to write to.
'RETURNS        The report: one EXPORTED line per component, then RESULT
'               PASS, or RESULT FAIL with the step and reason.
'ERROR POLICY   Contained: a failure ends the report. Files already written
'               stay written; Git restores them.
'STATE OWNERSHIP Reads the VBA project; changes only the listed files.
'UPDATED        2026-11-10
    Dim Names()     As String                   'Component names of the manifest
    Dim Paths()     As String                   'Their source files
    Dim Component   As Object                   'The component to export
    Dim Report      As String                   'Report lines so far
    Dim Stage       As String                   'Step under way, for a failure
    Dim i           As Long                     'Index into Names

    On Error GoTo Failed
        Stage = "manifest"
        ReadManifest RepoRoot, Names, Paths
        Stage = "trust access"
        RequireProjectAccess
        For i = LBound(Names) To UBound(Names)
            Stage = "export " & Names(i)
            Set Component = FindComponent(Names(i))
            If Component Is Nothing Then
                Err.Raise ERR_SYNC, , "the workbook has no component " & Names(i)
            End If
            ExportComponent Component, Paths(i)
            Report = Report & "EXPORTED " & Names(i) & vbCrLf
        Next i
        SyncExport = Report & RESULT_PASS
        Exit Function
Failed:
    SyncExport = Report & RESULT_FAIL & " at " & Stage & ": " & Err.Description
End Function


'------------------------------------------------------------------------------
'                              MANIFEST
'------------------------------------------------------------------------------

Private Sub ReadManifest( _
    ByVal RepoRoot As String, _
    ByRef Names() As String, _
    ByRef Paths() As String)
'PURPOSE        Read src/manifest.txt: one "name path" pair per line; lines
'               that are empty or start with # are skipped.
'INPUTS         RepoRoot: the clone's folder, with or without a final "\".
'RETURNS        Names and the full Windows paths of their source files.
'ERROR POLICY   Raises ERR_SYNC for a missing manifest or file, a line that
'               is not a pair, or an empty list.
'UPDATED        2026-11-10
    Dim Root        As String                   'RepoRoot without a final "\"
    Dim FilePath    As String                   'Full path of the manifest
    Dim FileNo      As Integer                  'File handle
    Dim Text        As String                   'The manifest's bytes as text
    Dim Lines()     As String                   'The manifest's lines
    Dim Fields()    As String                   'Name and path of one line
    Dim Entry       As String                   'One trimmed line
    Dim Count       As Long                     'Pairs read
    Dim i           As Long                     'Index into Lines

        Root = RepoRoot
        If Right$(Root, 1) = "\" Then Root = Left$(Root, Len(Root) - 1)
        FilePath = Root & "\" & MANIFEST_FILE
        If Len(Dir$(FilePath, vbNormal)) = 0 Then
            Err.Raise ERR_SYNC, , "no " & MANIFEST_FILE & " in " & Root
        End If
        FileNo = FreeFile
        Open FilePath For Binary Access Read As #FileNo
        Text = Space$(LOF(FileNo))
        Get #FileNo, , Text
        Close #FileNo
        Lines = Split(Replace(Text, vbCr, vbNullString), vbLf)
        ReDim Names(0 To UBound(Lines))
        ReDim Paths(0 To UBound(Lines))
        For i = LBound(Lines) To UBound(Lines)
            Entry = Trim$(Lines(i))
            If Len(Entry) > 0 And Left$(Entry, 1) <> "#" Then
                Fields = Split(Entry, " ")
                If UBound(Fields) <> 1 Then
                    Err.Raise ERR_SYNC, , "manifest line " & (i + 1) & " is not 'name path'"
                End If
                Names(Count) = Fields(0)
                Paths(Count) = Root & "\" & Replace(Fields(1), "/", "\")
                If Len(Dir$(Paths(Count), vbNormal)) = 0 Then
                    Err.Raise ERR_SYNC, , "no file " & Fields(1)
                End If
                Count = Count + 1
            End If
        Next i
        If Count = 0 Then Err.Raise ERR_SYNC, , "the manifest lists no component"
        ReDim Preserve Names(0 To Count - 1)
        ReDim Preserve Paths(0 To Count - 1)
End Sub


'------------------------------------------------------------------------------
'                              VBA PROJECT
'------------------------------------------------------------------------------

Private Sub RequireProjectAccess()
'PURPOSE        Fail early, with the remedy, when Excel denies access to the
'               VBA project.
'ERROR POLICY   Raises ERR_SYNC when the trust setting is off.
'UPDATED        2026-11-10
    Dim Count       As Long                     'Components, read to probe access

    On Error GoTo Denied
        Count = ThisWorkbook.VBProject.VBComponents.Count
        Exit Sub
Denied:
    Err.Raise ERR_SYNC, , "Excel does not allow access to the VBA project: turn on " & _
        "'Trust access to the VBA project object model' while the sync runs"
End Sub


Private Function BackupCopy() _
    As String
'PURPOSE        Save a copy of the workbook beside it before an import.
'RETURNS        The full path of the copy, stamped with the local time.
'ERROR POLICY   Raises ERR_SYNC for a workbook never saved; SaveCopyAs
'               errors propagate.
'STATE OWNERSHIP Writes one file; the open workbook does not change.
'UPDATED        2026-11-10
    Dim FullName    As String                   'Name of the saved workbook
    Dim Dot         As Long                     'Position of its extension
    Dim Stem        As String                   'Path of the copy without its extension
    Dim Target      As String                   'Path of the copy
    Dim Suffix      As Long                     'Makes the path unique within one second

        If Len(ThisWorkbook.Path) = 0 Then
            Err.Raise ERR_SYNC, , "save the workbook as .xlsm before a sync"
        End If
        FullName = ThisWorkbook.Name
        Dot = InStrRev(FullName, ".")
        If Dot = 0 Then Dot = Len(FullName) + 1
        Stem = ThisWorkbook.Path & "\" & Left$(FullName, Dot - 1) & ".backup-" & Format$(Now, "yyyymmdd-hhnnss")
        Target = Stem & Mid$(FullName, Dot)
        Suffix = 1
        Do While Len(Dir$(Target, vbNormal)) > 0
            Suffix = Suffix + 1
            Target = Stem & "-" & Suffix & Mid$(FullName, Dot)
        Loop
        ThisWorkbook.SaveCopyAs Target
        BackupCopy = Target
End Function


Private Sub ImportComponent( _
    ByVal ComponentName As String, _
    ByVal SourcePath As String)
'PURPOSE        Replace one component with its source file.
'INPUTS         ComponentName: the manifest's name. SourcePath: its file.
'ERROR POLICY   Raises ERR_SYNC for a sheet or workbook module of that name,
'               or a file that declares another name; VBE errors propagate.
'STATE OWNERSHIP Removes and adds that one component.
'UPDATED        2026-11-10
    Dim Components  As Object                   'The project's components
    Dim Existing    As Object                   'The component to replace
    Dim Imported    As Object                   'The component just imported
    Dim Spare       As Long                     'Number of a free temporary name

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
        Set Imported = Components.Import(SourcePath)
        If StrComp(Imported.Name, ComponentName, vbBinaryCompare) <> 0 Then
            Err.Raise ERR_SYNC, , SourcePath & " declares " & Imported.Name & ", not " & ComponentName
        End If
End Sub


Private Sub ExportComponent( _
    ByVal Component As Object, _
    ByVal TargetPath As String)
'PURPOSE        Export one component over its source file. The VBE refuses an
'               existing destination, so the export goes to a new file first,
'               which replaces the source only once it is written.
'INPUTS         Component: the component. TargetPath: its source file.
'ERROR POLICY   VBE and file errors propagate; the source stays untouched
'               until the export has succeeded.
'STATE OWNERSHIP Writes TargetPath and, briefly, TargetPath & ".export".
'UPDATED        2026-11-10
    Dim Fresh       As String                   'The new file

        Fresh = TargetPath & ".export"
        If Len(Dir$(Fresh, vbNormal)) > 0 Then Kill Fresh
        Component.Export Fresh
        Kill TargetPath
        Name Fresh As TargetPath
End Sub


Private Function FindComponent( _
    ByVal ComponentName As String) _
    As Object
'PURPOSE        Find a component of the project by name.
'RETURNS        The component, or Nothing.
'UPDATED        2026-11-10
    Dim Component   As Object                   'One component of the project

        For Each Component In ThisWorkbook.VBProject.VBComponents
            If StrComp(Component.Name, ComponentName, vbTextCompare) = 0 Then
                Set FindComponent = Component
                Exit Function
            End If
        Next Component
End Function


Private Function IsOwnedType( _
    ByVal ComponentType As Long) _
    As Boolean
'PURPOSE        True for the component types the sync may replace: standard
'               modules, classes and forms, never sheet or workbook modules.
'UPDATED        2026-11-10
        IsOwnedType = (ComponentType = TYPE_STD_MODULE Or ComponentType = TYPE_CLASS_MODULE _
                       Or ComponentType = TYPE_FORM)
End Function
