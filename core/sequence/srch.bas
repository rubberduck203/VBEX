Attribute VB_Name = "srch"
Option Explicit

' srch
' ====
'
' ### Max|Min
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
    
    seq.Assign MaxValue, sequence(MaxIndex(sequence, lower, upper))
    
End Function
''
' MinIndex
Public Function MinIndex(ByRef sequence As Variant, _
        ByVal lower As Long, ByVal upper As Long) As Long
    
    MinIndex = lower
    Dim i As Long
    For i = lower To upper
        If sequence(MinIndex) > sequence(i) Then MinIndex = i
    Next i
    
End Function
''
' MinValue
Public Function MinValue(ByRef sequence As Variant, _
        ByVal lower As Long, ByVal upper As Long) As Variant
    
    seq.Assign MinValue, sequence(MinIndex(sequence, lower, upper))
    
End Function
'
' ### Value Specific
'
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
Public Function BinarySearch(ByVal value As Variant, ByRef sortedSequence As Variant, _
        ByVal lower As Long, ByVal upper As Long, _
        Optional ByVal nearest As Boolean = False) As Long
    
    While lower < upper
        
        Dim middle As Long
        middle = seq.MiddleInt(lower, upper)
        
        If sortedSequence(middle) >= value Then
            upper = middle
        Else
            lower = middle + 1
        End If
        
    Wend
    
    BinarySearch = IIF(sortedSequence(upper) = value Or nearest, upper, -1)
    
End Function
