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
Implements IShowable
Implements IEquatable
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
Private Sub Class_Initialize()

    pArray = Array()

End Sub
Private Sub CopyArray(ByRef dest(), ByRef src())
    
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
Private Sub TupleSizeError(ByVal method As String, ByRef offending())

    Dim msg As String
    msg = "Offending sequence is of size: " & ArraySize(offending) + 1 & " not " & ArraySize(pArray) + 1

    IndexError Me, method, msg
        
End Sub
Private Function ArraySize(ByRef a()) As Long
    
    ArraySize = UBound(a) - LBound(a)
    
End Function
'
' Friend Methods
' --------------
'
Friend Sub Contain(ByVal vs)

    pArray = vs
    
End Sub
'
' Constructors
' ------------
'
''
' Pack: create a tuple from a parameter array
'
'    Tuple.Pack(1, 2, "A") ' (1, 2, A)
Public Function Pack(ParamArray vals()) As Tuple

   Set Pack = Implode(cast.CArray(vals))
   
End Function
''
' Implode: Transform an Iterable into a Tuple
' offsetting required for indexes
'
'     Tuple.Implode(Array(1, 2, "A")) ' (1, 2, A)
Public Function Implode(ByRef arr) As Tuple

    Dim result As New Tuple
    
    Dim a()
    a = cast.CArray(arr)
    
    Dim size As Long
    size = ArraySize(a)
    
    If Not ((IsEmpty(arr)) Or (size = -1)) Then
        
        Dim newArray()
        ReDim newArray(0 To size)
        
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
Public Function Zip(ByVal seed As IBuildable, ByVal xs As IIterable, _
        ByVal ys As IIterable) As IBuildable

    Dim result As IBuildable
    Set result = seed.MakeEmpty

    Dim lower As Long
    lower = srch.Max(xs.LowerBound, ys.LowerBound)
    
    Dim upper As Long
    upper = srch.Min(xs.UpperBound, ys.UpperBound)
    
    Dim i As Long
    For i = 1 To upper
        result.AddItem Pack(xs.Item(i), ys.Item(i))
    Next i

    Set Zip = result

End Function
'
' Public Methods
' --------------
'
Public Property Get Item(ByVal index)

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
Public Sub Unpack(ParamArray Elements())

    If ArraySize(pArray) = ArraySize(cast.CArray(Elements)) Then
        Dim i As Long
        For i = 0 To ArraySize(pArray)
            cast.Assign Elements(i), pArray(i)
        Next i
    Else
        TupleSizeError "Unpack", cast.CArray(Elements)
    End If
    
End Sub
''
' Explode: Load elements into array `arr`
' arr must be an array
'
Public Sub Explode(ByRef Elements())
    
    If ArraySize(pArray) = ArraySize(Elements) Then
        CopyArray Elements, pArray
    Else
        TupleSizeError "Explode", Elements
    End If
    
End Sub
Public Property Get Count() As Long

    Count = ArraySize(pArray) + 1
    
End Property
'
' ### IEquatable
'
Public Function Equals(ByVal other) As Boolean

    Equals = False
    
    If TypeName(other) <> TypeName(Me) Then
        Exit Function
    ElseIf Count <> other.Count Then
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To UBound(pArray)
        If Not defEquals.Equals(pArray(i), other.Item(i + 1)) Then
            Exit Function
        End If
    Next i
    
    Equals = True
    
End Function
Private Function IEquatable_Equals(ByVal other) As Boolean

    IEquatable_Equals = Equals(other)
    
End Function
'
'
' ### IShowable
'
Public Function Show() As String

    Show = defShow.ShowableObject(Me, pArray)
    
End Function
Public Function IShowable_Show() As String

    IShowable_Show = Show
    
End Function
