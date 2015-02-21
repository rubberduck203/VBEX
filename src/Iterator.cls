VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Iterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit
'
' Iterator
' ========
'
' Basic Mutable iterator class. 
Implements IMutableIterator
Implements ICountable
Implements IPrintable
Implements ICloneable
'
' Private Members
' ---------------
'
Private pIndex As Long
Private pTarget As Object
'
' Constructor
' -----------
'
Private Sub Class_Initialize()

    pIndex = 1

End Sub
Public Function Create(ByVal target As ICountable) As Iterator

    Dim result As New Iterator
    result.Assign target
    Set Create = result

End Function
'
' ### Friend Properties
'
Friend Sub Assign(ByVal aTarget As ICountable)
    
    Set pTarget = aTarget
    
End Sub
'
' Interfaces
' ----------
'
' ### IMutableIterator
'
Public Property Get DeRef() As Variant

    Dim result As Variant
    cast.Assign result, cast.Clone(pTarget(pIndex))

End Property
Public Property Get IMutableIterator_DeRef() As Variant

    cast.Assign IMutableIterator_DeRef, DeRef

End Property
Public Property Let DeRef(ByVal x As Variant)

    pTarget(pIndex) = x

End Property
Public Property Let IMutableIterator_DeRef(ByVal x As Variant)

    DeRef = x

End Property
Public Property Set DeRef(ByVal x As Variant)

    Set pTarget(pIndex) = x

End Property
Public Property Set IMutableIterator_DeRef(ByVal x As Variant)

    Set DeRef = x

End Property
Public Function Inc(Optional ByVal i As Long = 1) As Boolean

    pIndex = pIndex + i
    Inc = (pIndex < pTarget.Count)
    
End Property
Public Function IMutableIterator_Inc(Optional ByVal i As Long = 1) As Boolean
    
    IMutableIterator_Inc = Inc
    
End Property
Public Function Dec(Optional ByVal i As Long = 1) As Boolean

    pIndex = pIndex - i
    Dec = (pIndex > 1)
    
End Property
Public Function IMutableIterator_Dec(Optional ByVal i As Long = 1) As Boolean
    
    IMutableIterator_Dec = Dec
    
End Property
'
' ### ICountable
'
Public Property Get Count() As Long
    
    Count = pTarget.Count
  
End Property
Public Property Get ICountable_Count() As Long
    
  ICountable_Count = Count

End Property
'
' ### ICloneable
'
Public Property Get Clone() As Iterator
      
    Set Clone = Create(pTarget)
    
End Property
Public Property Get ICloneable_Clone() As Object
      
    Set ICloneable_Clone = Clone
    
End Property
'
' ### IPrintable
'
Public Property Get ToString() As String
      
    ToString = Printable.ParamObjectToString(Me, pIndex, pTarget)
    
End Property
Public Property Get IPrintable_ToString() As String
      
    IPrintable_ToString = ToString
    
End Property
