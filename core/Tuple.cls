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
'
' Private Members
' ---------------
'
''
' All memebers are stored in an Array
' ...simple enough
Private pArray As Variant

'
' Friend Methods
' --------------
'
''
'
Friend Sub Contain(ByVal Values As Variant)
    pArray = Values
End Sub
Friend Property Get Encapsulated() As Variant()
    Encapsulated = pArray
End Property
'
' Constructors
' ------------
'
''
' Pack: create a tuple from a parameter array
'
'    Tuple.Pack(1, 2, "A") ' (1, 2, A)
Public Function Pack(ParamArray values() As Variant) As Tuple
   Set Pack = Implode(values)
End Function
''
' Implode: Transform an Iterable into a Tuple
' offsetting required for indexes
'
'     Tuple.Implode(Array(1, 2, "A")) ' (1, 2, A) 
Public Function Implode(ByVal arr As Variant) As Tuple
    
    Dim arrLen As Long
    arrLen = UBound(arr)
    
    Dim dest() As Variant
    ReDim dest(arrLen) As Variant
    CopyArray dest, arr, arrLen
    
    Dim result As New Tuple
    result.Contain dest
    
    Set Implode = result
    
End Function
'
' Accessors
' ---------
'
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

    If UBound(pArray) = UBound(elements) Then
        
        Dim i As Long
        For i = 0 To UBound(pArray)
            elements(i) = pArray(i)
        Next i
        
    Else
        RaiseIndexError "Explode", elements
    End If
    
End Sub
''
' Explode: Load elements into array `arr`
' arr must be an array
'
Public Sub Explode(ByRef arr As Variant)
    Debug.Assert LBound(arr) = 0
    
    If UBound(pArray) = UBound(arr) Then
        CopyArray arr, pArray, UBound(pArray)
    Else
        RaiseIndexError "Explode", arr
    End If
    
End Sub
'
' Public Properties
' -----------------
'

'
' Public Methods
' --------------
'
Public Function Equals(ByVal other As Tuple) As Boolean
    Equals = False
    If Count <> other.Count Then Exit Function
    Dim i As Long
    For i = 0 To Ubound(pArray)
        If pArray(i) <> other.Encapsulated(i) Then Exit Function
    Next i
    Equals = True
End Function
'
' Private Methods
' ---------------
'
Private Sub CopyArray(ByRef dest As Variant, ByRef src As Variant, ByVal size As Long)
    Dim i As Long
    For i = 0 To size
        dest(i) = src(i)
    Next i
End Sub
Private Sub RaiseIndexError(ByVal method As String, ByVal offending As Variant)
    Err.Raise 9, TypeName(Me) & "." & method, _
        "Offending sequence is of size: " & UBound(offending) + 1 & " not " & UBound(pArray) + 1
End Sub
'
' Interfaces
' ----------
'
' ### ICountable
'
Public Property Get Count() As Long
    Count = Ubound(pArray) + 1
End Property
Public Property Get ICountable_Count() As Long
    ICountable_Count = Count
End Property
'
' ### IPrintable
'
Public Property Get ToString(Optional delim As String = ", ") As String
    ToString = "(" & Join(pArray, delim) & ")"
End Property
Public Property Get IPrintable_ToString() As String
    IPrintable_ToString = ToString
End Property
