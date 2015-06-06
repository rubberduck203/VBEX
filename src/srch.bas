Attribute VB_Name = "srch"
Option Explicit

' srch
' ====
'
' ### Max|Min
'
Private Function GenericExtremum(ByVal lg As CompareResult, _
        ByVal sequence As IIterable) As Long

    Dim result As Long
    result = sequence.LowerBound
    
    Dim curVal
    Assign curVal, sequence.Item(result)
    
    Dim i As Long
    For i = sequence.LowerBound To sequence.UpperBound
    
        If Not (Compare(curVal, sequence.Item(i)) = lg) Then
        
            result = i
            Assign curVal, sequence.Item(result)
            
        End If
        
    Next
    
    GenericExtremum = result

End Function
''
' MaxIndex: Returns the index of `sequence` that has the maximum value
Public Function MaxIndex(ByVal sequence As IIterable) As Long
    
    MaxIndex = GenericExtremum(gt, sequence)
    
End Function
''
' MaxValue: Returns the value of `sequence` that is the Maximum
' Uses `MaxIndex`
Public Function MaxValue(ByVal sequence As IIterable) As Variant
    
    Assign MaxValue, sequence.Item(MaxIndex(sequence))
    
End Function
Public Function Max(ParamArray vals() As Variant) As Variant

    Assign Max, MaxValue(List.Copy(vals))
    
End Function
''
' MinIndex
Public Function MinIndex(ByVal sequence As IIterable) As Long
    
    MinIndex = GenericExtremum(lt, sequence)
    
End Function
''
' MinValue
Public Function MinValue(ByVal sequence As IIterable) As Variant
    
    Assign MinValue, sequence.Item(MinIndex(sequence))
    
End Function
Public Function Min(ParamArray vals() As Variant) As Variant

    Assign Min, MinValue(List.Copy(vals))
    
End Function
'
' ### Value Specific
'
''
' LinearSearch:
Public Function LinearSearch(ByVal sought, ByVal sequence As IIterable) As Maybe
    
    Dim i As Long
    For i = sequence.LowerBound To sequence.UpperBound
        
        If Equals(sequence.Item(i), sought) Then
            Set LinearSearch = Maybe.Some(i)
            Exit Function
        End If
        
    Next i
    
    Set LinearSearch = Maybe.None
    
End Function
''
' Binary Search: Sequence must be sorted.  Has the option of returning where the
' value should be instead of not found.
Public Function BinarySearch(ByVal sought, ByVal sortedSequence As IIterable, _
        Optional ByVal nearest As Boolean = False) As Maybe
    
    Dim lower As Long
    lower = sortedSequence.LowerBound
    
    Dim upper As Long
    upper = sortedSequence.UpperBound
    
    Do While lower < upper
        
        Dim middle As Long
        middle = (lower + upper) \ 2
        
        Dim curVal
        Assign curVal, sortedSequence.Item(middle)
        
        If GreaterThanOrEqualTo(curVal, sought) Then
            upper = middle
        Else
            lower = middle + 1
        End If
        
    Loop
    
    Dim found As Boolean
    found = Equals(sortedSequence.Item(upper), sought)
    
    Set BinarySearch = Maybe.MakeIf(found Or nearest, upper)
    
End Function
