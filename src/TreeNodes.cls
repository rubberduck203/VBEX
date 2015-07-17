VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TreeNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private this As Dictionary

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all of the child nodes and frees the circular reference to their Parent."
        this.Clear
End Sub

Public Sub Add(ByVal node As TreeNode)
        this.Add node.name, node.Value
End Sub

Public Sub Remove(ByVal node As TreeNode)
Attribute Remove.VB_Description = "Removes a TreeNode from the collection and frees its circular reference to its Parent."
        Set Item.Parent = Nothing 'release circular reference
        this.Remove node.name
End Sub

Public Function Item(ByVal aName As String) As TreeNode
Attribute Item.VB_UserMemId = 0
        Set Item = this(aName)
End Function

Public Function Count() As Long
        Count = this.Count
End Function

Public Function Exists(ByVal aName As String) As Boolean
        Exists = this.Exists(aName)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
        Set NewEnum = this.Values.[_NewEnum]
End Function

Private Sub Class_Initialize()
        Set this = New Collection
End Sub

Private Sub Class_Terminate()
        Set this = Nothing
End Sub
