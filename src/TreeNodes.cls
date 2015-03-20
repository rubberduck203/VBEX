VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TreeNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private this As Collection

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all of the child nodes and frees the circular reference to their Parent."
	Dim node As TreeNode
	For Each node In this
		' Yes, I could call `Remove`, but this is more efficient.
		Set node.Parent = Nothing 'release circular reference
	Next
	
	Set this = New Collection
End Sub

Public Sub Add(ByVal Item As TreeNode)
	this.Add Item
End Sub

Public Sub Remove(ByVal Item As TreeNode)
Attribute Remove.VB_Description = "Removes a TreeNode from the collection and frees its circular reference to its Parent."
	Set Item.Parent = Nothing 'release circular reference
	this.Remove IndexOf(Item)
End Sub

Public Function Item(ByVal index As Long) As TreeNode
Attribute Item.VB_UserMemId = 0
	Set Item = this(index)
End Function

Public Function Count() As Long
	Count = this.Count
End Function

' Returns the index of item if found, otherwise returns 0.
Public Function IndexOf(ByVal node As TreeNode) As Long
Attribute IndexOf.VB_Description = "Returns the index of item if found, otherwise returns 0."
	Dim i As Long
	For i = 1 To this.Count
		If this.Item(i) Is node Then
			IndexOf = i
			Exit Function
		End If
	Next i
End Function

Public Function Exists(ByVal Name As String) As Boolean
	Dim i As Long
	For i = 1 To this.Count
		If this.Item(i).Name = Name Then
			Exists = True
			Exit Function
		End If
	Next
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
	Set NewEnum = this.[_NewEnum]
End Function

Private Sub Class_Initialize()
	Set this = New Collection
End Sub

Private Sub Class_Terminate()
	Set this = Nothing
End Sub
