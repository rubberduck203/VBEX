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
