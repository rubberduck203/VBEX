Attribute VB_Name = "seq"
Option Explicit

' seq
' ===
' Just Great Helper Functions
' ---------------------------
'
' Probably better to put it in another module but I think
' this will be the root of everything.
'
Public Sub Assign(ByRef x As Variant, ByVal y As Variant)
    
    If IsObject(y) Then
        Set x = y
    Else
        x = y
    End If
    
End Sub
''
' Can easily be extracted to use Floor(Average(a,b)) but I would like one 
' way dependence.
Public Function MiddleInt(ByVal a As Variant, ByVal b As Variant) As Variant
    MiddleInt = (a + b) \ 2
End Function

' Swapping
' --------

''
' Swap should work on an array or any two variables. It
' will not work on elements of sequence objects as the
' accessors of those return a value not a reference. For
' those use `SwapIndexes`.
'
' x = "a": y = "b"
' Swap x, y ' x="b", y="a"
'
' a = Array("a", "b")
' Swap a(0), a(1) ' a = [b, a]
'
Public Sub Swap(ByRef x As Variant, ByRef y As Variant)
    
    Dim t As Variant
    
    Assign t, x
    Assign x, y
    Assign y, t
    
End Sub
''
' `SwapIndexes` is to be used on sequence objects instead of `Swap`.  It uses the 
' default property of the object to access the elements.  If the default property
' is read-only an error is raised.  This cannot be used with a collection.
Public Sub SwapIndexes(ByRef sequence As Variant, ByVal a As Long, ByVal b As Long)
    ' We cannot implement `Assign` in case sequence is an object as `Collection.Item`
    ' returns a value and not a reference. Therefore we must reuse the pattern
    
    On Error GoTo IsCollection
    
    Dim t As Variant
    Assign t, sequence(a)
        
    If IsObject(sequence(b)) Then
        Set sequence(a) = sequence(b)
    Else
        sequence(a) = sequence(b)
    End If
    
    If IsObject(t) Then
        Set sequence(b) = t
    Else
        sequence(b) = t
    End If
    
CleanExit:
    Exit Sub
IsCollection:
    
    Err.Raise 13, "seq.SwapIndexes", "Sequence's default property is not read-write."
    
End Sub
''
' TODO: Raise Errors
Public Function Enumeration(ByVal a As Long, ByVal b As Long, _
        Optional ByVal s As Long = 1) As Variant()
        
    Debug.Assert s > 0
    
    Dim domain As Long
    domain = Abs(a - b)
    
    Debug.Assert domain Mod s = 0
    
    Dim size As Long
    size = domain \ s
    
    Dim result() As Variant
    ReDim result(size) As Variant
    
    s = IIF((a < b), s, -s)
    
    Dim i As Long, n As Long
    n = a
    For i = 0 To size
        result(i) = n
        n = n + s
    Next i
    
    Enumeration = result
    
End Function

Public Sub Fill(ByRef sequence As Variant, ByVal filler As Variant, _
        ByVal lower As Long, Byval upper As Long)
    
    Dim i As Long
    For i = lower To upper
        sequence(i) = filler
    Next i
    
End Sub
'
' Comparison
' ----------
'
Public Function Compare(ByRef seqA As Variant, ByRef seqB As Variant, _
        ByVal lowA As Long, ByVal highA As Long, _
        ByVal lowB As Long, ByVal highB As Long) As Boolean
    
    Compare = False

    If Not ((upA - lowA) = (highB - lowB)) The Exit Function
    
    Dim offset As Long
    offset = lowB - lowA
    
    Dim i As Long
    For i = lowA To highA
        If Not (seqA(i) = seqB(i + offset)) Then Exit Function
    Next i
    
    Compare = True
    
End Function
'
' In-Place Operations
' -------------------
' All In-Place operations must have bounds passed to support multiple data-types
'
' Reverse should accept an array or collection
Public Sub Reverse(ByRef sequence As Variant, _
        ByVal lower As Long, Byval upper As Long)
    
    While lower < upper
        
        SwapIndexes sequence, lower, upper
        
        lower = lower + 1
        upper = upper - 1
        
    Wend
    
End Sub
'
' Search
' ------
'
' ### Maximums
'
''
' MaxIndex: Returns the index of `sequence` that has the maximum value
Public Function MaxIndex(ByRef sequence As Variant, _
        ByVal lower As Long, ByVal upper As Long) As Long
    
    MaxIndex = lower
    Dim i As Long
    For i = lower To upper
        If sequence(MaxIndex) < sequence(i) Then MaxIndex = i
    Next i
    
End Function
''
' MaxValue: Returns the value of `sequence` that is the Maximum
' Uses `MaxIndex`
Public Function MaxValue(ByRef sequence As Variant, _
        ByVal lower As Long, ByVal upper As Long) As Variant
    
    Assign MaxValue, sequence(MaxIndex(sequence, lower, upper))
    
End Function
''
' LinearSearch:
Public Function LinearSearch(ByVal value As Variant, sequence As Variant, _
        ByVal lower As Long, ByVal upper As Long) As Long
    
    Dim i As Long
    For i = lower To upper
        
        If sequence(i) = value Then
            
            LinearSearch = i
            Exit Function
        
        End If
        
    Next i
    
    LinearSearch = -1
    
End Function
''
' Binary Search: Sequence must be sorted.  Has the option of returning where the
' value should be instead of not found.
Public Function BinarySearch(ByVal value As Variant, ByRef sorted_seq As Variant, _
        ByVal lower As Long, ByVal upper As Long, _
        Optional ByVal nearest As Boolean = False) As Long
    
    While lower < upper
        
        Dim middle As Long
        middle = MiddleInt(lower, upper)
        
        If sorted_seq(middle) >= value Then
            upper = middle
        Else
            lower = middle + 1
        End If
        
    Wend
    
    BinarySearch = IIF(sorted_seq(upper) = value Or nearest, upper, -1)
    
End Function
