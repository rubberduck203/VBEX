Attribute VB_Name = "castTest"
'@TestModule
Private Assert As New Rubberduck.AssertClass
Option Explicit

'
' cast Test
' =========
'
' Assign
' ------
'
'@TestMethod
Public Sub CastAssignPrimative()

    Dim x As Integer
    x = 1
    
    cast.Assign x, 2
    Assert.AreEqual 2, x
    
End Sub
'@TestMethod
Public Sub CastAssignArray()

    Dim a() As Variant
    a = List.Create(1, 2, 3).ToArray
    Assert.AreEqual "1 2 3", Join(a), "proving givens"
    
    Dim lower As Integer
    lower = LBound(a)
    
    cast.Assign a(lower), "x"
    Assert.AreEqual "x", a(lower)

End Sub
'@TestMethod
Public Sub CastAssignObject()

    Dim xs As List
    Set xs = List.Create(1, 2, 3)
    
    Dim ys As List
    Set ys = List.Create("A", "B", "C")
    
    ' Assign uses default "Set"
    cast.Assign xs, ys
    Assert.IsTrue xs.Equals(ys)
    Assert.AreEqual ObjPtr(ys), ObjPtr(xs)
    
    'Double Check
    ys.append "D"
    Assert.IsTrue xs.Equals(ys)
    
End Sub
'
' CArray
' ------
'
'@TestMethod
Public Sub CastCArray()

End Sub
'
' Interfaces
' ----------
'
' ### IPrintable
'
'@TestMethod
Public Sub CastToStringPrimativeData()

    With Assert
        .AreEqual "x", cast.ToString("x")
        .AreEqual "1", cast.ToString(CInt(1))
        .AreEqual "1", cast.ToString(CDbl(1))
    End With

End Sub
'@TestMethod
Public Sub CastToStringIPrintable()

    Dim flatList As List
    Set flatList = List.Create(1, 2, 3)
    
    Dim nestList As List
    Set nestList = List.Create(flatList, flatList)

    With Assert
        .AreEqual flatList.ToString, cast.ToString(flatList)
        .AreEqual nestList.ToString, cast.ToString(nestList)
    End With

End Sub
'@TestMethod
Public Sub CastSequenceToString()

    Dim flat As List
    Set flat = List.Create(1, 2, 3)
    
    Assert.AreEqual "List(1, 2, 3)", SequenceToString(flat, delim:=", ")
    
    Dim nested As List
    Set nested = List.Create(flat)
    
    Assert.AreEqual "List(List(1, 2, 3))", SequenceToString(nested, delim:=", ")

End Sub
'
' ### ICloneable
'
'
' ### IEquatable
'
'@TestMethod
Public Sub CastEqualsPrimatives()

    With Assert
        
        .IsTrue Equals(1, 1)
        .IsTrue Equals("a", "a")
        .IsFalse Equals(1, "a")
    
    End With

End Sub
'@TestMethod
Public Sub CastEqualsObjects()

    Dim flat As List
    Set flat = List.Create(1, 2, 3)
    
    Dim nested As List
    Set nested = List.Create(flat, flat, flat)

    With Assert
        
        .IsTrue Equals(flat, flat)
        .IsTrue Equals(nested, nested)
        .IsFalse Equals(flat, List.Create(1, 2))
        .IsFalse Equals(nested, List.Create(flat, flat))
    
    End With
    
End Sub
'@TestMethod
Public Sub CastEqualsMixed()

    With Assert
        .IsFalse Equals(List.Create(), "x")
        .IsTrue Equals(Application, "Microsoft Excel")
    End With

End Sub
'
' ### ICountable
'
'
' ### ISequence
'

