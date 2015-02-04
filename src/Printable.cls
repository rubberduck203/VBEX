VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Printable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Printable
' ---------
'
' Default Implementation of IPrintable Interface
'
Public Function ToString(ByVal x As Variant) As String

    Dim result As String
    
    If TypeOf x Is IPrintable Then
        result = x.ToString
    ElseIf IsObject(x) Then
        result = DefaultObjectToString(x)
    Else
        result = CStr(x)
    End If
    
    ToString = result
    
End Function
Private Function DefaultObjectToString(ByVal x As Object) As String

    Dim repr As String
    repr = "&" & ObjPtr(x)

    DefaultObjectToString = ObjectToString(x, cast.CArray(Array(repr)))

End Function
Public Function ParamObjectToString(ByVal o As Object, _
        ParamArray members() As Variant) As String
        
    ParamObjectToString = ObjectToString(o, cast.CArray(members))
        
End Function
Public Function ObjectToString(ByVal o As Object, ByRef members() As Variant, _
        Optional ByVal delim As String = ", ") As String

    Dim stringMembers() As String
    If LBound(members) <= UBound(members) Then
        ReDim stringMembers(LBound(members) To UBound(members))
    End If
    
    Dim i As Long
    For i = LBound(members) To UBound(members)
        stringMembers(i) = ToString(members(i))
    Next i
    
    ObjectToString = TypeName(o) & "(" & Join(stringMembers, delim) & ")"

End Function

