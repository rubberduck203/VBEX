VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TTreenode
	Name As String
	Value As Variant
	Children As TreeNodes
	Parent As TreeNode
End Type

Private this As TTreenode

Public Property Get Name() As String
	Name = this.Name
End Property

Public Property Let Name(ByVal Value As String)
	this.Name = Value
End Property

Public Property Get Value() As Variant
	AssignUnknown Value, this.Value
End Property

Public Property Set Value(ByVal newValue As Variant)
	Set this.Value = newValue
End Property

Public Property Let Value(ByVal newValue As Variant)
	this.Value = newValue
End Property

Public Property Get Children() As TreeNodes
	Set Children = this.Children
End Property

Public Property Get Parent() As TreeNode
	Set Parent = this.Parent
End Property

Public Property Set Parent(ByVal Value As TreeNode)
	Set this.Parent = Value
End Property

' If the argument already has a Parent, AddChild creates a shallow copy of the node to be added.
Public Function AddChild(ByVal node As TreeNode) As TreeNode
Attribute AddChild.VB_Description = "If the argument already has a Parent, AddChild creates a shallow copy of the node to be added."
	If node.Parent Is Nothing Then
		this.Children.Add node
		Set node.Parent = Me

		Set AddChild = node
	Else
		Dim copyOfNode As New TreeNode
		copyOfNode.Name = node.Name
		
		If Not IsEmpty(node.Value) Then
			'note: this should really use assign unknown, but it doesn't actually assign the value.
			'  I'm really not sure why, but I suspect it is because I'm trying to set properties.
		
			'AssignUnknown copyOfNode.Value, node.Value
			
			If IsObject(node.Value) Then
				Set copyOfNode.Value = node.Value
			Else
				copyOfNode.Value = node.Value
			End If
		End If
		
		Dim child As TreeNode
		For Each child In node.Children
			copyOfNode.AddChild child
		Next
		
		this.Children.Add copyOfNode
		Set copyOfNode.Parent = Me
		
		Set AddChild = copyOfNode
	End If

End Function

Public Function AddNewChild(ByVal Name As String) As TreeNode
Attribute AddNewChild.VB_Description = "Creates and Adds a New child node with the given Name."
	Dim child As TreeNode
	Set child = Me.AddChild(New TreeNode)
	child.Name = Name
	
	Set AddNewChild = child
End Function

Public Sub RemoveChild(ByVal node As TreeNode)
Attribute RemoveChild.VB_Description = "Removes the child node from this node's Children."
	With this.Children
		Set .Item(.IndexOf(node)).Parent = Nothing
	End With
	
	this.Children.Remove node
End Sub

Public Function HasChildren() As Boolean
	HasChildren = (this.Children.Count <> 0)
End Function

Public Function Path(Optional ByVal separator As String = "\") As String
Attribute Path.VB_Description = "Uses the node names to build a Path String. If a node name is empty, the path will have consecutive separators."
	Dim result As String
	result = Me.Name
	
	If Me.HasChildren Then
		result = result & separator
	End If
	
	If Not Me.Parent Is Nothing Then
		result = Me.Parent.Path(separator) & result
	End If
	
	Path = result
	
End Function

Public Function ToString() As String
	ToString = "Name: " & this.Name & "; ValueType: " & TypeName(this.Value)
End Function

Private Sub AssignUnknown(ByRef destination As Variant, ByVal source As Variant)
	If IsObject(source) Then
		Set destination = source
	Else
		destination = source
	End If
End Sub

Private Sub Class_Initialize()
	Set this.Children = New TreeNodes
End Sub
