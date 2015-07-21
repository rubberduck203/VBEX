VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Assoc Class
' ===========
'
'
' Copywrite (C) 2014 Philip Wales
' This file (Assoc.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
'
Implements IShowable
'
' Private Members
' ---------------
'
Private pPair As Tuple
'
'
'
Public Function Make(ByVal k, ByVal v) As Assoc

    Dim result As New Assoc
    result.AssignKeyVal k, v
    Set Make = result

End Function
Friend Sub AssignKeyVal(ByVal k, ByVal v)

    Set pPair = Tuple.Pack(k, v)
    
End Sub
Public Property Get key() As Variant
    
    Assign key, pPair.Item(1)
    
End Property
Public Property Get Value() As Variant
    
    Assign Value, pPair.Item(2)
    
End Property
'
'
'
Public Function Show() As String

    Show = defShow.Show(key) & " -> " & defShow.Show(Value)

End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
