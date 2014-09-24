VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeProcedures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' *******************************************************************
' * Collection Class Wrapper for vbeProcedure objects
' * Author: Christopher J. McClellan
' * DateCreated: 5/30/2014
' *******************************************************************

Private mCollection As Collection

Public Function Create() As vbeProcedures
    Set Create = New vbeProcedures
End Function

Public Sub Clear()
Attribute Clear.VB_Description = "Clears the collection and removes any procedure in it from memory."
    killVbeProcs
    Set mCollection = New Collection
End Sub

Public Function Add(ByRef vbeProc As vbeProcedure, Optional ByVal Key As Variant)
    If IsMissing(Key) Then
        mCollection.Add vbeProc
    Else
        mCollection.Add vbeProc, Key
    End If
End Function

Public Function Remove(ByVal Index As Variant)
    mCollection.Remove (Index)
End Function

Public Function item(ByVal Index As Variant) As vbeProcedure
Attribute item.VB_Description = "Gets the procedure at the specified index."
Attribute item.VB_UserMemId = 0
    Set item = mCollection(Index)
End Function

Public Function Count() As Long
    Count = mCollection.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCollection.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set mCollection = New Collection
End Sub

Private Sub Class_Terminate()
    killVbeProcs
    Set mCollection = Nothing
End Sub

Private Sub killVbeProcs()
    Dim proc As vbeProcedure
    If Not mCollection Is Nothing Then
        For Each proc In mCollection
            Set proc = Nothing
        Next proc
    End If
End Sub

