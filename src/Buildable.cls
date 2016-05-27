VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Buildable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Buildable
' ==========
'
Public Function MakeEmpty() As Buildable
    Exceptions.NotImplementedError Me, "MakeEmpty"
End Function
Public Sub AddItem(ByVal x)
    Exceptions.NotImplementedError Me, "AddItem"
End Sub
Public Sub AddItems(ByVal x)
    Exceptions.NotImplementedError Me, "AddItems"
End Sub
