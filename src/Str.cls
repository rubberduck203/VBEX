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
'
' Constructors
' ------------
'
Public Function Make(ByVal s As String) As Str
    
    Dim result As New Str
    result.Encapsulated = s
    Set Make = result

End Function
Public Function Join(ByVal strings, _
        Optional ByVal delim As String = vbNullString) As Str

    Dim strList As List
    Set strList = List.Copy(Str)

    Set Join = Make(Join(strList.ToArray, delim))

End Function
Public Function Format(ByVal fmt As String, ParamArray vals()) As Str
  
    Dim result As Str
    Set result = Str.Make(fmt)

    Dim index As Long
    index = 0

    Dim placeHolder As String
    placeHolder = "{0}"

    Dim val
    For Each val In vals
        result.Substitute placeHolder, defShow.Show(val)
        index = index + 1
        Mid$(placeHolder, 2, 1) = index
    Next

End Function
Friend Property Let Encapsulated(ByVal s As String)

    pString = s

End Property
Public Property Get NewEnum() As IUnknown

    Static exploded As Collection
    Set exploded = Explode
    Set NewEnum = exploded.[_NewEnum]
    
End Property
'
' Private Methods
' ---------------
'
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
'
' Public Methods
' --------------
'
Public Sub Append(ParamArray strings())
    
    Concat CVar(strings)
    
End Sub
Public Sub Concat(ByVal strings)
    
    Dim s
    For Each s In strings
        pString = pString & defShow.Show(s)
    Next
    
End Sub
Public Sub Substitute(ByVal oldString As String, ByVal newString As String)

    pString = Replace(pString, oldString, newString)
    'Alt
    'Dim loc As Long
    'loc = InStr(1, pString, oldString)
    '
    'If loc > 0 Then
    '    Mid$(pString, loc, len(newString) = newString
    'End If

End Sub
Public Function Count() As Long
    
    Count = Len(pString)
    
End Function
'
' IBuildable
' ----------
'
Private Function IBuildable_MakeEmpty() As IBuildable
    
    Set IBuildable_MakeEmpty = Make(vbNullString)
    
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

