VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit
'
' Tuple
' =====
'
' Implementation of Tuple in VB
'
'Implements IEquatable ' but only against other tuples...
Implements ICountable
Implements IPrintable
Implements ISequence
'
' Private Members
' ---------------
'
''
' All memebers are stored in an Array
' ...simple enough
Private pArray() As Variant
'
' Private Methods
' ---------------
'
Private Sub CopyArray(ByRef dest() As Variant, ByRef src() As Variant)
    
    Dim lower As Long
    lower = LBound(src)
    
    Dim upper As Long
    upper = UBound(src)
    
    Dim offset As Long
    offset = LBound(dest) - lower
    
    Dim i As Long
    For i = lower To upper
        cast.Assign dest(i + offset), src(i)
    Next i
    
End Sub
Private Sub RaiseIndexError(ByVal method As String, ByRef offending() As Variant)

    Err.Raise 9, TypeName(Me) & "." & method, _
        "Offending sequence is of size: " & ArraySize(offending) + 1 & " not " & ArraySize(pArray) + 1
        
End Sub
Private Function ArraySize(ByRef a() As Variant) As Long
    
    ArraySize = UBound(a) - LBound(a)
    
End Function
'
' Friend Methods
' --------------
'
''
'
Friend Sub Contain(ByVal Values As Variant)

    pArray = Values
    
End Sub
'
' Constructors
' ------------
'
''
' Pack: create a tuple from a parameter array
'
'    Tuple.Pack(1, 2, "A") ' (1, 2, A)
Public Function Pack(ParamArray Values() As Variant) As Tuple

   Set Pack = Implode(cast.CArray(Values))
   
End Function
''
' Implode: Transform an Iterable into a Tuple
' offsetting required for indexes
'
'     Tuple.Implode(Array(1, 2, "A")) ' (1, 2, A)
Public Function Implode(ByRef arr As Variant) As Tuple

    Dim result As New Tuple
    
    Dim a() As Variant
    a = cast.CArray(arr)
    
    Dim size As Long
    size = ArraySize(a)
    
    If Not ((IsEmpty(arr)) Or (size = -1)) Then
        
        Dim newArray() As Variant
        ReDim newArray(0 To size) As Variant
        
        CopyArray newArray, a
        result.Contain newArray
        
    Else
        result.Contain Array()
    End If
    
    Set Implode = result
    
End Function
''
' Zip, create a list of tuples...
' TODO: use param array instead of just 2
Public Function Zip(ByVal xs As List, ByVal ys As List) As List
    
    Dim result As List
    Set result = List.Create

    Dim i As Long
    For i = 1 To srch.Min(xs.Count, ys.Count)
        result.Append Pack(xs(i), ys(i))
    Next i
    
    Set Zip = result
    
End Function
'
' Public Methods
' --------------
'
Public Property Get Item(ByVal index As Variant)

    cast.Assign Item, pArray(index - 1)
    
End Property
''
' Load elements into a parameter array of varaibles
'
'   Set t = Tuple.Pack(1, 2, "A")
'   t.Unpack x, y, z ' x = 1, y = 2, z = "A"
'
' In other languages this is the same as
'
'     x, y, z = t
'
Public Sub Unpack(ParamArray elements() As Variant)

    If ArraySize(pArray) = ArraySize(cast.CArray(elements)) Then
        Dim i As Long
        For i = 0 To ArraySize(pArray)
            cast.Assign elements(i), pArray(i)
        Next i
    Else
        RaiseIndexError "Unpack", cast.CArray(elements)
    End If
    
End Sub
''
' Explode: Load elements into array `arr`
' arr must be an array
'
Public Sub Explode(ByRef elements() As Variant)
    
    If ArraySize(pArray) = ArraySize(elements) Then
        CopyArray elements, pArray
    Else
        RaiseIndexError "Explode", elements
    End If
    
End Sub
'
'Public Function Equals(ByVal other As Tuple) As Boolean
'    Equals = False
'    If Count <> other.Count Then Exit Function
'    Dim i As Long
'    For i = 0 To UBound(pArray)
'        If pArray(i) <> other.Encapsulated(i) Then Exit Function
'    Next i
'    Equals = True
'End Function
'
' Interfaces
' ----------
'
'
' ### ISequence
'
Public Function ToArray() As Variant()
     
    Dim size As Long
    size = ArraySize(pArray)
    
    If size <> -1 Then
    
        Dim result() As Variant
        ReDim result(0 To size)
        CopyArray result, pArray
        ToArray = result
        
    Else
        ToArray = Array()
    End If

End Function
Public Function ISequence_ToArray() As Variant()

    ISequence_ToArray = ToArray

End Function
'
' ### ICountable
'
Public Property Get Count() As Long

    Count = ArraySize(pArray) + 1
    
End Property
Public Property Get ICountable_Count() As Long

    ICountable_Count = Count
    
End Property
'
' ### IPrintable
'
Public Property Get ToString(Optional delim As String = ", ") As String

    ToString = cast.SequenceToString(Me, delim)
    
End Property
Public Property Get IPrintable_ToString() As String

    IPrintable_ToString = ToString
    
End Property
