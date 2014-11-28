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

Implements IOutput
' Prints `object` to Immediate Pane in Office VBE

Public Sub PrintLine(Optional ByVal object As Variant)
    If IsMissing(object) Then
        'newline
        Debug.Print vbNullString
    Else
        Debug.Print object
    End If
End Sub
Public Sub Put(Optional ByVal object As Variant)
    If Not IsMissing(object) Then Debug.Print object;
End Sub
'
' IOutput Interfaces.
'
Private Sub IOutput_PrintLine(Optional ByVal object As Variant)
    PrintLine object
End Sub
Private Sub IOutput_Put(ByVal object As Variant)
    Put object
End Sub
