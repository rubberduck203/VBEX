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

'Implements IOutput
' Prints `x` to Immediate Pane in Office VBE

Public Sub PrintLine(Optional ByVal x As Variant)

    If IsMissing(x) Then
        Debug.Print vbNullString
    Else
        Debug.Print cast.ToString(x)
    End If
    
End Sub
Public Sub Puts(Optional ByVal x As Variant)

    If Not IsMissing(x) Then Debug.Print cast.ToString(x);
    
End Sub
'
' IOutput Interfaces.
'
Private Sub IOutput_PrintLine(Optional ByVal x As Variant)

    PrintLine x
    
End Sub
Private Sub IOutput_Puts(ByVal x As Variant)
  
    Puts x
  
End Sub
