VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TTree
    Root As TreeNode
End Type

Private this As TTree

Public Property Get Root() As TreeNode
    Set Root = this.Root
End Property

Public Property Set Root(ByVal Value As TreeNode)
    Set this.Root = Value
End Property

Private Sub Class_Initialize()
    Set this.Root = New TreeNode
End Sub
