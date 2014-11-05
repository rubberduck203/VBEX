Attribute VB_Name = "seq"
Option Explicit

' seq
' ===
'
' This has become a dumping ground for useful methods that don't
' really have a home.  PLEASE FIND OR MAKE THEM ONE!
'
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

