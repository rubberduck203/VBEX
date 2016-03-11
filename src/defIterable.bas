Attribute VB_Name = "defIterable"
''
' Converts an iterable to an array.  The LBound and UBound of the array will be
' the same as the LowerBound and UpperBound of the iterable. Unless the iterable
' is empty then an empty array will be returned, whose bounds are always (0,-1).
Public Function ToArray(ByVal iterable As IIterable) As Variant()

    Dim lower As Long
    lower = iterable.LowerBound

    Dim upper As Long
    upper = iterable.UpperBound
    
    Dim result()
    If lower <= upper Then
        
        ReDim result(lower To upper)
        
        Dim i As Long
        For i = lower To upper
            Assign result(i), iterable.Item(i)
        Next
    
    Else
        result = Array()
    End If
    
    ToArray = result
    
End Function
''
' Converts an iterable to a collection
Public Function ToCollection(ByVal iterable As IIterable) As Collection

    Dim lower As Long
    lower = iterable.LowerBound

    Dim upper As Long
    upper = iterable.UpperBound

    Dim result As New Collection

    Dim i As Long
    For i = lower To upper
        result.Add iterable.Item(i)
    Next
    
    Set ToCollection = result

End Function
''
' Converts an iterable to any IBuildable
Public Function ToBuildable(ByVal seed As IBuildable, ByVal iterable As IIterable) _
        As IBuildable

    Dim result As IBuildable
    Set result = seed.MakeEmpty
    
    Dim index As Long
    For index = iterable.LowerBound To iterable.UpperBound
        result.AddItem iterable(index)
    Next

    Set ToBuildable = result

End Function
