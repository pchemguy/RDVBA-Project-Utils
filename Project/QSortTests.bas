Attribute VB_Name = "QSortTests"
'@Folder "Common.QuickSort"
'@TestModule
'@IgnoreModule AssignmentNotUsed, VariableNotUsed, LineLabelNotUsed, UnhandledOnErrorResumeNext, IndexedDefaultMemberAccess
Option Explicit
Option Private Module


#Const LateBind = LateBindTests
#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("QSort")
Private Sub NumericArrayFullSort()
    On Error GoTo TestFail

Arrange:
    Dim Sample() As Variant
    Sample = Array(45, 30, 25, 15, 10, 5, 40, 20, 35, 50, 75, 85, 60, 80, 55, 65, 70, 75)
Act:
    QuickSortArray Sample
Assert:
    Assert.AreEqual "51015202530354045505560657075758085", Join(Sample, vbNullString)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QSort")
Private Sub NumericArrayPartialSort()
    On Error GoTo TestFail

Arrange:
    Dim Sample() As Variant
    Sample = Array(45, 30, 25, 15, 10, 5, 40, 20, 35, 50, 75, 85, 60, 80, 55, 65, 70, 75)
Act:
    QuickSortArray Sample, 5, 12
Assert:
    Assert.AreEqual "45302515105203540506075858055657075", Join(Sample, vbNullString)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QSort")
Private Sub NumericArray1FullSort()
    On Error GoTo TestFail

Arrange:
    Dim Sample() As Variant
    Sample = Array(45)
Act:
    QuickSortArray Sample
Assert:
    Assert.AreEqual 45, Sample(0)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QSort")
Private Sub NumericArray2FullSort()
    On Error GoTo TestFail

Arrange:
    Dim Sample() As Variant
    Sample = Array(45, 15)
Act:
    QuickSortArray Sample
Assert:
    Assert.AreEqual 45, Sample(1)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QSort")
Private Sub TextArrayFullSort()
    On Error GoTo TestFail

Arrange:
    Dim Sample() As Variant
    Sample = Array("Kas", "Qman", "Cs", "Ib", "Zd", "Csg", "bs", "afeee", "i", "Oddd")
Act:
    QuickSortArray Sample
Assert:
    Assert.AreEqual "afeeebsCsCsgiIbKasOdddQmanZd", Join(Sample, vbNullString)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QSort")
Private Sub TextArrayPartialSort()
    On Error GoTo TestFail

Arrange:
    Dim Sample() As Variant
    Sample = Array("Kas", "Qman", "Cs", "Ib", "Zd", "Csg", "bs", "afeee", "i", "Oddd")
Act:
    QuickSortArray Sample, 2, 7
Assert:
    Assert.AreEqual "KasQmanafeeebsCsCsgIbZdiOddd", Join(Sample, vbNullString)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
