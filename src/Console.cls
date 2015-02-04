VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Console
' =======
'
' Representation of the VBE Immediate Pane
'
Implements IOutput
'
' Interfaces
' ----------
'
' ### IOutput
'
' #### Prints
'
Public Sub Prints(Optional ByVal x As Variant)

    If Not IsMissing(x) Then
        Debug.Print Printable.ToString(x);
    End If
    
End Sub
Private Sub IOutput_Prints(ByVal x As Variant)
  
    Prints x
  
End Sub
'
' #### PrintLine
'
Public Sub PrintLine(Optional ByVal x As Variant)

    Prints x
    Debug.Print
    
End Sub
Private Sub IOutput_PrintLine(ByVal x As Variant)

    PrintLine x
    
End Sub
'
' #### PrintLines
'
Public Sub PrintLines(ByRef xs() As Variant)

    Dim x As Variant
    For Each x In xs
        PrintLine x
    Next x

End Sub
Public Sub IOutput_PrintLines(ByRef xs() As Variant)

    PrintLines xs

End Sub

