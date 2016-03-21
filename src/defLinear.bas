Attribute VB_Name = "defLinear"
'
' # Default Implementations for Linear collections
'
' We can only assume that the collection

''
'
Public Function Size(ByVal sequence As Linear) As Long

    Size = sequence.UpperBound - sequence.LowerBound + 1

End Function
''
' Converts an sequence to an array.  The LBound and UBound of the array will be
' the same as the LowerBound and UpperBound of the sequence. Unless the sequence
' is empty then an empty array will be returned, whose bounds are always (0,-1).
Public Function ToArray(ByVal sequence As Linear) As Variant()

    Dim lower As Long
    lower = sequence.LowerBound

    Dim upper As Long
    upper = sequence.UpperBound
    
    Dim result()
    If lower <= upper Then
        
        ReDim result(lower To upper)
        
        Dim i As Long
        For i = lower To upper
            Assign result(i), sequence.Item(i)
        Next
    
    Else
        result = Array()
    End If
    
    ToArray = result
    
End Function
''
' Converts an sequence to a collection
Public Function ToCollection(ByVal sequence As Linear) As Collection

    Dim lower As Long
    lower = sequence.LowerBound

    Dim upper As Long
    upper = sequence.UpperBound

    Dim result As New Collection

    Dim i As Long
    For i = lower To upper
        result.Add sequence.Item(i)
    Next
    
    Set ToCollection = result

End Function
''
' Converts an sequence to any Buildable
Public Function ToBuildable(ByVal seed As Buildable, ByVal sequence As Linear) _
        As Buildable

    Dim result As Buildable
    Set result = seed.MakeEmpty
    
    Dim index As Long
    For index = sequence.LowerBound To sequence.UpperBound
        result.AddItem sequence(index)
    Next

    Set ToBuildable = result

End Function
Public Function IndexWhere(ByVal sequence As Linear, ByVal pred As Applicable) _
        As Maybe

    Dim result As Maybe
    Set result = Maybe.None
    
    Dim i As Long
    i = sequence.LowerBound
    
    Dim upper As Long
    upper = sequence.UpperBound
    
    Do While result.IsNone And i <= upper
        If pred.Apply(sequence.Item(i)) Then
            Set result = Maybe.Some(i)
        End If
    Loop
    
    Set IndexWhere = result

End Function
Public Function IndexOf(ByVal sequence As Linear, ByVal val As Variant) As Maybe

    Dim pred As Applicable
    Set pred = InternalDelegate.Make("Equals").Partial(val, Empty)
    
    Set IndexOf = IndexWhere(sequence, pred)
    
End Function
Public Function LastIndexWhere(ByVal sequence As Linear, _
        ByVal pred As Applicable) As Maybe

    Dim result As Maybe
    Set result = Maybe.None
    
    Dim i As Long
    i = sequence.UpperBound
    
    Dim lower As Long
    lower = sequence.LowerBound
    
    Do While result.IsNone And i >= lower
        If pred.Apply(sequence.Item(i)) Then
            Set result = Maybe.Some(i)
        End If
    Loop
    
    Set LastIndexWhere = result
    
End Function
Public Function LastIndexOf(ByVal sequence As Linear, ByVal val As Variant) _
        As Maybe

    Dim pred As Applicable
    Set pred = InternalDelegate.Make("Equals").Partial(val, Empty)
    
    Set LastIndexOf = LastIndexWhere(sequence, pred)
    
End Function
Public Function Find(ByVal sequence As Linear, ByVal pred As Applicable) As Maybe

    Set Find = IndexWhere(sequence, pred).Map(OnArgs.Make("Item", VbGet, sequence))
    
End Function
Public Function FindLast(ByVal sequence As Linear, ByVal pred As Applicable) As Maybe

    Set FindLast = LastIndexWhere(sequence, pred).Map(OnArgs.Make("Item", VbGet, sequence))

End Function
Public Function CountWhere(ByVal sequence As Linear, ByVal pred As Applicable) As Long

    Dim result As Long
    result = 0
    
    Dim i As Long
    For i = sequence.LowerBound To sequence.UpperBound
        If pred.Apply(sequence.Item(i)) Then
            result = result + 1
        End If
    Next

    CountWhere = result

End Function
