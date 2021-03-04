VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Common.Project Utils"
Option Explicit

Private Const PROJECT_FOLDER As String = "Project"
Private Const COMMON_FOLDER As String = "Common"
Private Const REFERENCES_FILE As String = "References.csv"

Private Type TProjectUtils
    Project As VBIDE.VBProject
    fso As Scripting.FileSystemObject
    ProjectPath As String
End Type
Private this As TProjectUtils


Private Sub Class_Initialize()
    Select Case Application.Name
        Case "Microsoft Excel"
            Set this.Project = Application.ActiveWorkbook.VBProject
        Case "Microsoft Access", "Microsoft Word"
            Set this.Project = Application.VBE.ActiveVBProject
        Case "Microsoft PowerPoint"
            Set this.Project = Application.VBE.VBProjects(1)
    End Select
    
    Set this.fso = New Scripting.FileSystemObject
    this.ProjectPath = this.fso.GetParentFolderName(this.Project.Filename) & Application.PathSeparator & PROJECT_FOLDER
    If Not this.fso.FolderExists(this.ProjectPath) Then this.fso.CreateFolder this.ProjectPath
End Sub


Private Sub Class_Terminate()
    Set this.Project = Nothing
    Set this.fso = Nothing
End Sub


Public Sub ReferencesSaveToFile()
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
    
    For Each Reference In References
        ReferenceString = Reference.Name & _
                  vbTab & Reference.GUID & _
                  vbTab & CStr(Reference.Major) & _
                  vbTab & CStr(Reference.Minor) & _
                  vbTab & Reference.FullPath
        ReferencesArray(ReferenceIndex) = ReferenceString
        ReferenceIndex = ReferenceIndex + 1
    Next Reference
    
    Dim PathName As String
    PathName = this.ProjectPath & Application.PathSeparator & REFERENCES_FILE
    Dim FileHandle As Long: FileHandle = FreeFile
    Open PathName For Output As #FileHandle
    Print #FileHandle, Join(ReferencesArray, vbNewLine)
    Close #FileHandle
End Sub


Public Sub ReferencesAddFromFile(Optional ByVal UseGUID = True)
    Dim PathName As String
    PathName = this.ProjectPath & Application.PathSeparator & REFERENCES_FILE
    Dim ReadBuffer As String
    
    Dim FileHandle As Long: FileHandle = FreeFile
    Open PathName For Binary As #FileHandle
    ReadBuffer = Space$(LOF(FileHandle))
    Get #FileHandle, , ReadBuffer
    Close #FileHandle
    
    Dim refj As Long
    If Len(ReadBuffer) > 0 Then
        Dim ReadLines() As String
        ReadLines = Split(ReadBuffer, vbNewLine)
        Dim ReferencesArray() As Variant
        ReDim ReferencesArray(0 To UBound(ReadLines, 1) - 2)
        For refj = 0 To UBound(ReadLines, 1) - 2
            ReferencesArray(refj) = Split(ReadLines(refj + 1), vbTab)
        Next refj
    End If
    
    On Error Resume Next
    If UseGUID Then
        For refj = 0 To UBound(ReferencesArray, 1)
            this.Project.References.AddFromGuid ReferencesArray(refj)(1), CLng(ReferencesArray(refj)(2)), CLng(ReferencesArray(refj)(3))
        Next refj
    Else
        For refj = 0 To UBound(ReferencesArray, 1)
            this.Project.References.AddFromFile ReferencesArray(refj)(4)
        Next refj
    End If
    On Error GoTo 0
End Sub
