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
Implements Output
'
' Interfaces
' ----------
'
' ### Output
'
' #### Prints
'
Public Sub Prints(ByVal x As Variant)

    Debug.Print defShow.Show(x);
    
End Sub
Private Sub Output_Prints(ByVal x As Variant)
  
    Prints x
  
End Sub
'
' #### PrintLine
'
Public Sub PrintLine(ByVal x As Variant)

    Prints x
    Debug.Print
    
End Sub
Private Sub Output_PrintLine(ByVal x As Variant)

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
Private Sub Output_PrintLines(ByRef xs() As Variant)

    PrintLines xs

End Sub

