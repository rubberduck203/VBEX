VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit
'
' Vector
' =====
'
Implements IBuildable
Implements IShowable
Implements IIterable

'
' Private Members
' ---------------
'
''
' All memebers are stored in an Array
Private Const START_SIZE As Long = 32
Private pArray() As Variant
Private pCount As Long
'
' Private Methods
' ---------------
'
Private Sub PushBack(ByVal x)

    Dim upper As Long
    upper = UBound(pArray)
    
    pCount = pCount + 1
    
    If pCount > upper Then
        ReDim Preserve pArray(0 To 2 * upper)
    End If
    
    Assign pArray(pCount - 1), x

End Sub
'Private Sub PopBack()
'
'    Dim nextDown As Long
'    nextDown = UBound(pArray) \ 2
'
'    If pCount.PreDec(1) < nextDown And nextDown > START_SIZE Then
'        ReDim Preserve pArray(0 To nextDown)
'    End If
'
'End Sub
'
'
' Class
' -----
'
Private Sub Class_Initialize()

    ReDim pArray(0 To START_SIZE)
    pCount = 0
    
End Sub
Public Function Create(ParamArray Values()) As Vector

   Set Create = Copy(CVar(Values))
   
End Function
Public Function Copy(ByVal sequence) As Vector

    Dim result As New Vector
    result.Extend sequence
    Set Copy = result
    
End Function

Public Sub Append(ParamArray Elements())

    Extend CVar(Elements)
    
End Sub
Public Sub Extend(ByVal sequence)
    
    Dim element
    For Each element In sequence
        PushBack element
    Next element

End Sub
Public Function Count() As Long

    Count = pCount
    
End Function
'
' IBuildable
' ----------
'
Private Function IBuildable_MakeEmpty() As IBuildable
    
    Set IBuildable_MakeEmpty = Vector.Create
    
End Function
Private Sub IBuildable_AddItem(ByVal Elements)

    Extend CVar(Elements)
    
End Sub
Private Sub IBuildable_AddItems(ByVal sequence)

    IBuildable_AddItem sequence

End Sub
'
' IIterable
' ---------
'
Public Property Get Item(ByVal index As Long) As Variant

    CheckIndexRange Me, "Item", index
    Assign Item, pArray(index)
    
End Property
Public Property Let Item(ByVal index As Long, ByVal val)

    CheckIndexRange Me, "Item", index
    pArray(index) = val
    
End Property
Public Property Set Item(ByVal index As Long, ByVal val)

    CheckIndexRange Me, "Item", index
    Set pArray(index) = val
    
End Property
Public Function LowerBound() As Long

    LowerBound = 0

End Function
Public Function UpperBound() As Long

    UpperBound = Count - 1
    
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
    
    IIterable_ToArray = ToArray
    
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

    Show = defShow.ShowableObject(Me, ToArray)
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show
    
End Function
