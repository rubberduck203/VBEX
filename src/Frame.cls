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
' Wrapper for a Linear with mutable bounds
'
' Copywrite (C) 2014 Philip Wales
' This file (Frame.cls) is distributed under the GPL-3.0 license
' Obtain a copy of the GPL-3.0 license <http://opensource.org/licenses/GPL-3.0>
'
Implements Showable
Implements Linear
'
' Private Members
' ---------------
'
Private pLinearSequence As Linear
Private pLower As Long
Private pUpper As Long
'
' Constructors
' ------------
'
Public Function Make(ByVal linSeq As Linear, _
        ByVal lower As Long, ByVal upper As Long) As Frame
    
    Dim result As New Frame
    result.Inject linSeq, lower, upper
    Set Make = result

End Function
Public Function FromLeft(ByVal linSeq As Linear, _
        ByVal upper As Long) As Frame

    Set FromLeft = Make(linSeq, linSeq.LowerBound, upper)

End Function
Public Function FromRight(ByVal linSeq As Linear, _
        ByVal lower As Long) As Frame

    Set FromRight = Make(linSeq, lower, linSeq.UpperBound)

End Function
Friend Sub Inject(ByVal linSeq As Linear, ByVal lower As Long, ByVal upper As Long)

    CheckIndexRange linSeq, "Item", lower
    CheckIndexRange linSeq, "Item", upper
    Set pLinearSequence = linSeq
    pLower = lower
    pUpper = upper

End Sub
''
' Public Methods
' --------------
'
Public Function ConvertTo(ByVal seed As Buildable) As Buildable

    Dim result As Buildable
    Set result = seed.MakeEmpty

    Dim i As Long
    For i = pLower To pUpper
        seed.AddItem Item(i)
    Next

    Set ConvertTo = result

End Function
Public Function MapTo(ByVal seed As Buildable, ByVal op As Applicable) As Buildable

    Set MapTo = defMap.IterableMap(seed, op, Me)

End Function
Public Function FlatMapTo(ByVal seed As Buildable, ByVal op As Applicable) As Buildable

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
' Linear
' ---------
'
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0

    CheckIndexRange Me, "Item", index
    Assign Item, pLinearSequence.Item(index + pLower)
    
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

Private Function Linear_Item(ByVal index As Long) As Variant

    Assign Linear_Item, Item(index)

End Function
Private Function Linear_LowerBound() As Long

    Linear_LowerBound = LowerBound

End Function
Private Function Linear_UpperBound() As Long

    Linear_UpperBound = UpperBound

End Function
Public Function ToArray() As Variant()

    ToArray = defIterable.ToArray(Me)

End Function
Private Function Linear_ToArray() As Variant()

    Linear_ToArray = ToArray()

End Function
Public Function ToCollection() As Collection

    Set ToCollection = defIterable.ToCollection(Me)

End Function
Private Function Linear_ToCollection() As Collection

    Set Linear_ToCollection = ToCollection
    
End Function
Public Function ToBuildable(ByVal seed As Buildable) As Buildable

    Set ToBuildable = defIterable.ToBuildable(seed, Me)

End Function
Private Function Linear_ToBuildable(ByVal seed As Buildable) As Buildable

    Set Linear_ToBuildable = ToBuildable(seed)

End Function
'
' Showable
' ---------
'
Public Function Show() As String

    Show = defShow.ParamShowableObject(Me, pLinearSequence, pLower, pUpper)
    
End Function
Private Function Showable_Show() As String

    Showable_Show = Show
    
End Function
