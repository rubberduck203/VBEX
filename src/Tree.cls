VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Type TTree
	Root As TreeNode
End Type

Private this As TTree
Public Function Create() As Tree
	
	Dim result As New Tree
	Set result.Root = New TreeNode
	Set Create = result

End Function
Friend Property Set Root(ByVal node As TreeNode)
	
	Set this.Root = node
	
End Property
Public Property Get Root() As TreeNode
Attribute Root.VB_UserMemId = 0
	
    Set Root = this.Root
	
End Property

