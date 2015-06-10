VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit
'
' Window
' ======
'
' Wrapper for a IIterable with mutable bounds
'
Implements IShowable
Implements IIterable
'
' Private Members
' ---------------
'
Private pIterable As IIterable
Private pLower As Long
Private pUpper As Long
'
'
' Private Methods
' ---------------
'
' ### Errors
'
'
' Public Methods
' --------------
'
Public Function ToArray() As Variant()

    Dim upper As Long
    upper = UpperBound
    
    Dim results()
    If upper >= 0 Then
        
        ReDim results(upper)
        
        Dim i As Long
        For i = 0 To upper
            Assign results(i), Item(i)
        Next
    
    Else
        results = Array()
    End If
    
    ToArray = results
    
End Function
'
' Constructors
' ------------
'
Public Function Make(ByVal iterable As IIterable, _
        ByVal lower As Long, ByVal upper As Long) As Window
    
    Dim result As New Window
    Set result.Target = iterable
    result.SetLowerBound lower
    result.SetUpperBound upper

    Set Make = result

End Function
Public Function FromLeft(ByVal iterable As IIterable, _
        ByVal upper As Long) As Window

    Set FromLeft = Make(iterable, iterable.LowerBound, upper)

End Function
Public Function FromRight(ByVal iterable As IIterable, _
        ByVal lower As Long) As Window

    Set FromRight = Make(iterable, lower, iterable.UpperBound)

End Function
Friend Property Set Target(ByVal iterable As IIterable)

    Set pIterable = iterable

End Property
Friend Sub SetLowerBound(ByVal lower As Long)

    CheckIndexRange pIterable, "Item", lower
    pLower = lower

End Sub
Friend Sub SetUpperBound(ByVal upper As Long)

    CheckIndexRange pIterable, "Item", upper
    pUpper = upper

End Sub
''
' IIterable
' ---------
'
Public Property Get Item(ByVal index As Long) As Variant

    CheckIndexRange Me, "Item", index
    Assign Item, pIterable.Item(index + pLower)
    
End Property
Public Function LowerBound() As Long

    LowerBound = 0

End Function
''
' [ 0 | 1 | 2 | 3 | 4 ]
'     [ 0 | 1 ]
Public Function UpperBound() As Long

    UpperBound = pUpper - pLower
    
End Function
Private Function IIterable_Item(ByVal index As Long) As Variant

    Assign IIterable_Item, Item(index)

End Function
Private Function IIterable_LowerBound() As Long

    IIterable_LowerBound = LowerBound

End Function
Private Function IIterable_UpperBound() As Long

    IIterable_UpperBound = UpperBound

End Function
'
' IShowable
' ---------
'
Public Function Show() As String

    Show = defShow.ParamShowableObject(Me, pIterable, pLower, pUpper)
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show
    
End Function
