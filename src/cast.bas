Attribute VB_Name = "cast"
Public Sub Assign(ByRef x As Variant, ByVal y As Variant)
    
    If IsObject(y) Then
        Set x = y
    Else
        x = y
    End If
    
End Sub
Public Function CArray(ByVal xs As Variant) As Variant()

    CArray = xs
    
End Function
Public Function ToString(ByVal x As Variant) As String

    If IsObject(x) Then
        If TypeOf x Is IPrintable Then
            ToString = x.ToString
        End If
    Else
        ToString = CString(x)
    End If

End Function
Public Function SequenceToString(Byval xs As ISequence, Optional ByVal delim As String, _
        Optional ByVal lcap As String, Optional ByVal rcap As String) As String

    Dim ss() As Variant
    ss = xs.ToArray
    
    Dim i As Long
    For i = LBound(ss) To Ubound(ss)
        ss(i) = ToString(ss(i))
    Next i
    
    SequenceToString = lcap & Join(ss, delim) & rcap
    
End Function
