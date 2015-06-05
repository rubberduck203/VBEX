VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Str"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Str Class
' =========
'
'
' Copywrite (C) 2014 Philip Wales
' This file (Str.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
'
Implements IBuildable
Implements IShowable
Implements IIterable
'
Private pString As String
Public Function Create(ParamArray strings()) As Str
    
    Set Create = Copy(CVar(strings))

End Function
Public Function Copy(ByVal strings) As Str
    
    Dim result As New Str
    result.Concat strings
    Set Copy = result

End Function
Public Property Get NewEnum() As IUnknown

    Static exploded As Collection
    Set exploded = Explode
    Set NewEnum = exploded.[_NewEnum]
    
End Property
Private Function Explode() As Collection

    Dim length As Long
    length = Len(pString)

    Dim result As New Collection
    
    If length > 0 Then
    
        Dim i As Long
        For i = 1 To length
            result.Add Mid$(pString, i, 1)
        Next
        
    End If
    
    Set Explode = result

End Function
Public Sub Append(ParamArray strings())
    
    Concat CVar(strings)
    
End Sub
Public Sub Concat(ByVal strings)
    
    Dim s
    For Each s In strings
        pString = pString & defShow.Show(s)
    Next
    
End Sub
Public Function Count() As Long
    
    Count = Len(pString)
    
End Function
'
' IBuildable
' ----------
'
Private Function IBuildable_MakeEmpty() As IBuildable
    
    Set IBuildable_MakeEmpty = Create()
    
End Function
Private Sub IBuildable_AddItem(ByVal element)
    
    Append element
    
End Sub
Private Sub IBuildable_AddItems(ByVal sequence)
    
    Concat sequence
    
End Sub
'
' IShowable
' ---------
'
Public Function Show() As String
Attribute Show.VB_UserMemId = 0
   
    Show = pString
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
'
' IIterable
' ---------
'
Public Function Item(ByVal index As Long) As String
    
    Item = Mid(pString, index, 1)
    
End Function

Public Function LowerBound() As Long

    LowerBound = 1
    
End Function
Public Function UpperBound() As Long

    UpperBound = Count
    
End Function
Private Function IIterable_Item(ByVal index As Long) As Variant

    IIterable_Item = Item(index)
    
End Function
Private Function IIterable_LowerBound() As Long

    IIterable_LowerBound = LowerBound
    
End Function
Private Function IIterable_UpperBound() As Long

    IIterable_UpperBound = UpperBound
    
End Function

