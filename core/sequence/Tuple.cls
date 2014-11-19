VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private pArray As Variant

Friend Sub Contain(ByVal Values As Variant)
    pArray = Values
End Sub
Public Function Pack(ParamArray Values() As Variant) As Tuple
   Set Pack = Implode(Values)
End Function
''
' offsetting required for indexes
Public Function Implode(ByVal arr As Variant) As Tuple
    
    Dim arrLen As Long
    arrLen = UBound(arr)
    
    Dim dest() As Variant
    ReDim dest(arrLen) As Variant
    CopyArray dest, arr, arrLen
    
    Dim result As New Tuple
    result.Contain dest
    
    Set Implode = result
    
End Function
Public Sub Unpack(ParamArray elements() As Variant)

    If UBound(pArray) = UBound(elements) Then
        
        Dim i As Long
        For i = 0 To UBound(pArray)
            elements(i) = pArray(i)
        Next i
        
    Else
        RaiseIndexError "Explode", elements
    End If
    
End Sub
Public Sub Explode(ByRef arr As Variant)
    Debug.Assert LBound(arr) = 0
    
    If UBound(pArray) = UBound(arr) Then
        CopyArray arr, pArray, UBound(pArray)
    Else
        RaiseIndexError "Explode", arr
    End If
    
End Sub
Private Sub CopyArray(ByRef dest As Variant, ByRef src As Variant, ByVal size As Long)
    Dim i As Long
    For i = 0 To size
        dest(i) = src(i)
    Next i
End Sub
Private Sub RaiseIndexError(ByVal method As String, ByVal offending As Variant)
    Err.Raise 9, TypeName(Me) & "." & method, _
        "Offending sequence is of size: " & UBound(offending) + 1 & " not " & UBound(pArray) + 1
End Sub
Public Property Get ToString() As String
    Debug.Print "(" & Join(pArray, ", ") & ")"
End Property
