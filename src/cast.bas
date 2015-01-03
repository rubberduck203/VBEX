Attribute VB_Name = "cast"
Public Sub Assign(ByRef x As Variant, ByVal y As Variant)
    
    If IsObject(y) Then
        Set x = y
    Else
        x = y
    End If
    
End Sub
Public Function CArray(ByVal a As Variant) As Variant()

    CArray = a
    
End Function

