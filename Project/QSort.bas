Attribute VB_Name = "QSort"
'@Folder "Common.QuickSort"
Option Explicit
Option Compare Text

'
' Sorts an array of numbers or strings using quicksort algorithm.
' It first creates an array of indecies and orders it.
' Once sorting is complete, the ordered array of indecies
' is used to order the original array. (This indirect
' approach is used to reduce the number of swaps of the
' elements in the original array, which can be strings
' and, in general, objects.)
' Ascii strings are sorted in case insensitive fashion (Option Compare Text).
'
Public Sub QuickSortArray(ByRef ValueArray As Variant, _
                 Optional ByVal Start As Long = -1, _
                 Optional ByVal Finish As Long = -1)
                 
    If IsEmpty(ValueArray) Then Exit Sub
    
    If Start = -1 Then Start = LBound(ValueArray, 1)
    If Finish = -1 Then Finish = UBound(ValueArray, 1)
    
    '''' Prepare an array of indecies to be sorted.
    '''' Each element in ValueIndex says that in the position "i" of the sorted array
    '''' should be an element with index ValueIndex(i) from the original unordered array.
    Dim ValueIndex() As Long
    ReDim ValueIndex(Start To Finish)
    Dim i As Long
    For i = Start To Finish
        ValueIndex(i) = i
    Next i
    
    Randomize

    '''' Sort array of indecies
    QuickSortArrayCore ValueArray, ValueIndex, Start, Finish
    
    '''' Use sorted ValueIndex array to order original array in place
    Dim Buffer As Variant
    Dim ValuePointer As Long
    Dim ValuePointerNext As Long
    Dim ValuePointerBuffer As Long
    ValuePointerNext = Finish
    
    '''' The original unordered array may be represented by one or several rings, such that
    '''' within each ring elements change positions between themselves. Starting from a given
    '''' end/direction locate the first element that is out of order and order the elements
    '''' in the ValueArray and ValueIndex, tracing the corresponding ring. If that ring does
    '''' not include all elements, continue scanning the ValueIndex array and find the first
    '''' remaining element that is out of order and repeat cycle, until both arrays are oredered.
    Do While ValuePointerNext > Start
        '''' Start from the last element and go backwards. Find the first element that is out of order
        ValuePointer = ValuePointerNext
        Do While (ValueIndex(ValuePointer) = ValuePointer) And (ValuePointer > Start)
            ValuePointer = ValuePointer - 1
        Loop
        If ValuePointer = Start Then
            Exit Sub
        Else
            ValuePointerNext = ValuePointer - 1
        End If
        
        '''' Order the current ring
        Buffer = ValueArray(ValuePointer)
        Do While ValueIndex(ValuePointer) < ValuePointerNext + 1
            ValueArray(ValuePointer) = ValueArray(ValueIndex(ValuePointer))
            ValuePointerBuffer = ValueIndex(ValuePointer)
            ValueIndex(ValuePointer) = ValuePointer
            ValuePointer = ValuePointerBuffer
        Loop
        ValueArray(ValuePointer) = Buffer
        ValueIndex(ValuePointer) = ValuePointer
    Loop
End Sub


Private Sub QuickSortArrayCore(ByRef ValueArray As Variant, _
                               ByRef ValueIndex As Variant, _
                      Optional ByVal Start As Long = -1, _
                      Optional ByVal Finish As Long = -1)
    Dim PivotValue As Variant
    Dim PivotIndex As Long
    Dim LeftIndex As Long
    Dim RightIndex As Long
    Dim Buffer As Long

    If Start >= Finish Then Exit Sub
    
    PivotIndex = Start + CLng(Round(Rnd * (Finish - Start)))
    PivotValue = ValueArray(ValueIndex(PivotIndex))
    LeftIndex = Start
    RightIndex = Finish
                        
    Do While LeftIndex < RightIndex
        Do While (ValueArray(ValueIndex(RightIndex)) > PivotValue) And (LeftIndex < RightIndex)
            RightIndex = RightIndex - 1
        Loop
        Do While (ValueArray(ValueIndex(LeftIndex)) < PivotValue) And (LeftIndex < RightIndex)
             LeftIndex = LeftIndex + 1
        Loop
        If LeftIndex < RightIndex Then
            Buffer = ValueIndex(RightIndex)
            ValueIndex(RightIndex) = ValueIndex(LeftIndex)
            ValueIndex(LeftIndex) = Buffer
            RightIndex = RightIndex - 1
            LeftIndex = LeftIndex + 1
        End If
    Loop
    
    '''' Handle edge cases
    If LeftIndex > RightIndex Then
        RightIndex = RightIndex + 1
        LeftIndex = LeftIndex - 1
    ElseIf LeftIndex = RightIndex Then
        If RightIndex = Finish Then
            LeftIndex = LeftIndex - 1
        ElseIf LeftIndex = Start Then
            RightIndex = RightIndex + 1
        ElseIf ValueArray(ValueIndex(RightIndex)) >= PivotValue Then
            LeftIndex = LeftIndex - 1
        Else
            RightIndex = RightIndex + 1
        End If
    End If
          
    If Start < LeftIndex Then
        QuickSortArrayCore ValueArray, ValueIndex, Start, LeftIndex
    End If
    If RightIndex < Finish Then
        QuickSortArrayCore ValueArray, ValueIndex, RightIndex, Finish
    End If
End Sub


