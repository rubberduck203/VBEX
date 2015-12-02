VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Frame
' =====
'
' Wrapper for a IIterable with mutable bounds
'
' Copywrite (C) 2014 Philip Wales
' This file (Frame.cls) is distributed under the GPL-3.0 license
' Obtain a copy of the GPL-3.0 license <http://opensource.org/licenses/GPL-3.0>
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
' Constructors
' ------------
'
Public Function Make(ByVal iterable As IIterable, _
        ByVal lower As Long, ByVal upper As Long) As Frame
    
    Dim result As New Frame
    Set result.Target = iterable
    result.SetLowerBound lower
    result.SetUpperBound upper

    Set Make = result

End Function
Public Function FromLeft(ByVal iterable As IIterable, _
        ByVal upper As Long) As Frame

    Set FromLeft = Make(iterable, iterable.LowerBound, upper)

End Function
Public Function FromRight(ByVal iterable As IIterable, _
        ByVal lower As Long) As Frame

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
' Public Methods
' --------------
'
Public Function ConvertTo(ByVal seed As IBuildable) As IBuildable

    Dim result As IBuildable
    Set result = seed.MakeEmpty

    Dim i As Long
    For i = pLower To pUpper
        seed.AddItem Item(i)
    Next

    Set ConvertTo = result

End Function
Public Function MapTo(ByVal seed As IBuildable, ByVal op As IApplicable) As IBuildable

    Set MapTo = defMap.IterableMap(seed, op, Me)

End Function
Public Function FlatMapTo(ByVal seed As IBuildable, ByVal op As IApplicable) As IBuildable

    Set FlatMapTo = defMap.IterableBind(seed, op, Me)
    
End Function
' FilterTo and all the others.....
''
' Transversable
' -------------
'
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Static copyCollec As Collection
    Set copyCollec = ToCollection
    Set NewEnum = copyCollec.[_NewEnum]

End Property
''
' IIterable
' ---------
'
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0

    CheckIndexRange Me, "Item", index
    Assign Item, pIterable.Item(index + pLower)
    
End Property
Public Function LowerBound() As Long

    LowerBound = 1

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
Public Function ToArray() As Variant()

    ToArray = defIterable.ToArray(Me)

End Function
Private Function IIterable_ToArray() As Variant()

    IIterable_ToArray = ToArray()

End Function
Public Function ToCollection() As Collection

    Set ToCollection = defIterable.ToCollection(Me)

End Function
Private Function IIterable_ToCollection() As Collection

    Set IIterable_ToCollection = ToCollection
    
End Function
Public Function ToBuildable(ByVal seed As IBuildable) As IBuildable

    Set ToBuildable = defIterable.ToBuildable(seed, Me)

End Function
Private Function IIterable_ToBuildable(ByVal seed As IBuildable) As IBuildable

    Set IIterable_ToBuildable = ToBuildable(seed)

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
