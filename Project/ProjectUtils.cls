VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Common.Project Utils"
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed
Option Explicit

Private Const PROJECT_FOLDER As String = "Project"
'Private Const COMMON_FOLDER As String = "Common"
Private Const REFERENCES_FILE As String = "References.csv"

Private Type TProjectUtils
    Project As VBIDE.VBProject
    ProjectPath As String
    Paths As Scripting.Dictionary
    Files As Scripting.Dictionary
    fso As Scripting.FileSystemObject
    wsh As IWshRuntimeLibrary.WshShell
    EnvVarNames As Variant
End Type
Private this As TProjectUtils


Private Sub Class_Initialize()
    With this
        Select Case Application.Name
            Case "Microsoft Excel"
                Set .Project = Application.ActiveWorkbook.VBProject
            Case "Microsoft Access", "Microsoft Word"
                Set .Project = Application.VBE.ActiveVBProject
            Case "Microsoft PowerPoint"
                '@Ignore IndexedDefaultMemberAccess
                Set .Project = Application.VBE.VBProjects(1)
        End Select
        
        Set .Paths = New Scripting.Dictionary
        .Paths.CompareMode = TextCompare
        Set .Files = New Scripting.Dictionary
        .Files.CompareMode = TextCompare
        Set .fso = New Scripting.FileSystemObject
        Set .wsh = New IWshRuntimeLibrary.WshShell
        .ProjectPath = .fso.GetParentFolderName(.Project.Filename) & Application.PathSeparator & PROJECT_FOLDER
        If Not .fso.FolderExists(.ProjectPath) Then .fso.CreateFolder .ProjectPath
        .EnvVarNames = Array("CommonProgramFiles(x86)", "CommonProgramFiles", "ProgramFiles (x86)", "ProgramFiles", "SystemRoot")
    End With
End Sub


Private Sub Class_Terminate()
    With this
        Set .Project = Nothing
        Set .Paths = Nothing
        Set .Files = Nothing
        Set .fso = Nothing
        Set .wsh = Nothing
    End With
End Sub


'@Description "Save references info for the Project of the ActiveWorkbook into tsv/csv file."
Public Sub ReferencesSaveToFile()
Attribute ReferencesSaveToFile.VB_Description = "Save references info for the Project of the ActiveWorkbook into tsv/csv file."
    Dim References As VBIDE.References
    Set References = this.Project.References
    Dim ReferenceString As String
    Dim ReferencesArray() As Variant
    ReDim ReferencesArray(0 To References.Count)
    Dim Reference As VBIDE.Reference
    Dim ReferenceIndex As Long: ReferenceIndex = 0
    
    ReferenceString = "Name" & _
              vbTab & "GUID" & _
              vbTab & "Major" & _
              vbTab & "Minor" & _
              vbTab & "FullPath"
    ReferencesArray(ReferenceIndex) = ReferenceString
    ReferenceIndex = ReferenceIndex + 1
    
    Dim EnvVarIndex As Long
    Dim FullPath As String
    Dim FullPathLen As Long
    Dim EnvVarCount As Long: EnvVarCount = UBound(this.EnvVarNames, 1) + 1
    For Each Reference In References
        '''' Replace path prefix with corresponding common environment variable if exists
        FullPath = Reference.FullPath
        FullPathLen = Len(FullPath)
        EnvVarIndex = 0
        Do While (Len(FullPath) = FullPathLen) And (EnvVarIndex < EnvVarCount)
            FullPath = Replace(FullPath, Environ(this.EnvVarNames(EnvVarIndex)), "%" & this.EnvVarNames(EnvVarIndex) & "%", 1, 1, vbTextCompare)
            EnvVarIndex = EnvVarIndex + 1
        Loop
        If Len(FullPath) = FullPathLen Then
            FullPath = Replace(FullPath, "C:\ProgramFiles", "%ProgramFiles%", 1, 1, vbTextCompare)
        End If
        
        ReferenceString = Reference.Name & _
                  vbTab & Reference.GUID & _
                  vbTab & CStr(Reference.Major) & _
                  vbTab & CStr(Reference.Minor) & _
                  vbTab & FullPath
        ReferencesArray(ReferenceIndex) = ReferenceString
        ReferenceIndex = ReferenceIndex + 1
    Next Reference
    
    '''' Save references to a tab separated file in the project folder
    Dim PathName As String
    PathName = this.ProjectPath & Application.PathSeparator & REFERENCES_FILE
    Dim FileHandle As Long: FileHandle = FreeFile
    Open PathName For Output As #FileHandle
    Print #FileHandle, Join(ReferencesArray, vbNewLine)
    Close #FileHandle
End Sub


'''' Errors during reference addition are ignored (expect the major source due to already
'''' activated references). Alternatively, read activated refernces first and skip addition
'''' of activated references.
'@Description "Add references from the tsv/csv file to the Project of the ActiveWorkbook ."
Public Sub ReferencesAddFromFile(Optional ByVal UseGUID As Boolean = True)
Attribute ReferencesAddFromFile.VB_Description = "Add references from the tsv/csv file to the Project of the ActiveWorkbook ."
    Dim PathName As String
    PathName = this.ProjectPath & Application.PathSeparator & REFERENCES_FILE
    Dim ReadBuffer As String
    
    Dim FileHandle As Long: FileHandle = FreeFile
    Open PathName For Input As #FileHandle
    ReadBuffer = Input$(LOF(FileHandle), #FileHandle)
    Close #FileHandle
    
    Dim refj As Long
    If Len(ReadBuffer) > 0 Then
        '''' Split buffer into record lines
        Dim ReadLines() As String
        ReadLines = Split(ReadBuffer, vbNewLine)
        Dim ReferencesArray() As Variant
        ReDim ReferencesArray(0 To UBound(ReadLines, 1) - 2)
        For refj = 0 To UBound(ReadLines, 1) - 2
            '''' Split records into fields
            ReferencesArray(refj) = Split(ReadLines(refj + 1), vbTab)
            ReferencesArray(refj)(4) = this.wsh.ExpandEnvironmentStrings(ReferencesArray(refj)(4))
        Next refj
    End If
    
    '''' Add all references ignoring errors
    On Error Resume Next
    If UseGUID Then
        For refj = 0 To UBound(ReferencesArray, 1)
            this.Project.References.AddFromGuid ReferencesArray(refj)(1), CLng(ReferencesArray(refj)(2)), CLng(ReferencesArray(refj)(3))
            If VBA.Information.Err.Number > 0 Then Debug.Print VBA.Information.Err.Description
        Next refj
    Else
        For refj = 0 To UBound(ReferencesArray, 1)
            this.Project.References.AddFromFile ReferencesArray(refj)(4)
            If VBA.Information.Err.Number > 0 Then Debug.Print VBA.Information.Err.Description
        Next refj
    End If
    On Error GoTo 0
End Sub


Public Sub ProjectStructureParse()
    this.Paths.RemoveAll
    this.Files.RemoveAll
    
    Dim Component As VBComponent
    Dim AnnotateStart As Long
    Dim AnnotateStop As Long
    Dim Path As String
    Dim Ext As String
    Dim ComponentType As String
    Dim ComponentDeclareLines As String

End Sub
