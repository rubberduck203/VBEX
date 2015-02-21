VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMutableIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Property Get DeRef() As Variant
End Property

Public Property Let DeRef(ByVal x As Variant)
End Property

Public Property Set DeRef(ByVal x As Variant)
End Property

Public Function Inc(Optional ByVal i As Long = 1) As Boolean
End Property

Public Function Dec(Optional ByVal i As Long = 1) As Boolean
End Property
