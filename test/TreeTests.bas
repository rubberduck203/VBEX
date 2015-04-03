Attribute VB_Name = "TreeTests"
Option Explicit

Option Private Module

'@TestModule
Private Assert As New Rubberduck.AssertClass

Private t As Tree

'@TestInitialize
Public Sub TestInitialize()
	Set t = New Tree
	t.Root.Name = "C:"
End Sub

'@TestCleanup
Public Sub TestCleanup()
	Set t = Nothing
End Sub

'@TestMethod
Public Sub RootNodeIsNotNothingOnTreeCreation()
	On Error GoTo TestFail
	
	'Arrange:
		Dim myTree As Tree
		Set myTree = New Tree
	'Act:

	'Assert:
	Assert.IsNotNothing myTree.Root

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub RootIsNotNothingAfterSetting()
	'Arrange:
	Set t = New Tree
	
	'Act:
	Set t.Root = New TreeNode
	
	'Assert
	Assert.IsNotNothing t.Root
End Sub

'@TestMethod
Public Sub AddingAChildToRoot()
	On Error GoTo TestFail
	
	'Arrange:
		Dim child As New TreeNode
	'Act:
		t.Root.AddChild child
	
	'Assert:
	Assert.AreSame child, t.Root.Children(1)

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub AddChildToChild()
	On Error GoTo TestFail
	
	Const expected As Long = 1
	
	'Arrange:
		Dim child As TreeNode
		Set child = t.Root.AddChild(New TreeNode)
		child.Name = "Users"
	'Act:
		Set child = child.AddChild(New TreeNode)
		child.Name = "username"

	'Assert:
	Assert.AreEqual expected, t.Root.Children.Count
	Assert.AreEqual expected, t.Root.Children(1).Children.Count
	Assert.AreEqual "Users", t.Root.Children(1).Name
	Assert.AreEqual "username", t.Root.Children(1).Children(1).Name

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub ChildTracksParent()
	On Error GoTo TestFail
	
	'Arrange:
		Dim child As TreeNode
	'Act:
		Set child = t.Root.AddChild(New TreeNode)
		child.Name = "Users"
		
	'Assert:
	Assert.AreEqual "C:", child.Parent.Name

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub ParentIsNotNothingAfterRemovingChild() 'TODO: Rename test
	On Error GoTo TestFail
	
	'Arrange:
		Const expectedCount As Long = 0
		
		Dim child As TreeNode
		Set child = t.Root.AddChild(New TreeNode)
	'Act:
		t.Root.RemoveChild child
	
	'Assert:
	Assert.AreEqual expectedCount, t.Root.Children.Count
	Assert.IsNotNothing t.Root
	Assert.AreEqual "C:", t.Root.Name

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub HasChildrenTrue()
	On Error GoTo TestFail
	
	'Arrange:
		Set t.Root = New TreeNode

	'Act:
		t.Root.AddChild New TreeNode
	'Assert:
	Assert.IsTrue t.Root.HasChildren

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub HasChildrenFalseOnCreation()
	On Error GoTo TestFail
	
	'Arrange:
	'Act:
	
	'Assert:
	Assert.IsFalse t.Root.HasChildren

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub LeafPathToString()
	On Error GoTo TestFail
	
	'Arrange:
		Const expected As String = "C:\Users\username\test.txt"
		
		Dim child As TreeNode
		Set child = t.Root.AddChild(New TreeNode)
		child.Name = "Users"
		
		Set child = child.AddChild(New TreeNode)
		child.Name = "username"
		
		Set child = child.AddChild(New TreeNode)
		child.Name = "test.txt"
	'Act:
		Dim actual As String
		actual = child.Path
	'Assert:
	Assert.AreEqual expected, actual

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub NodePathToString()
	On Error GoTo TestFail
	
	'Arrange:
		Const expected As String = "C:\Users\"
		
		Dim child As TreeNode
		Set child = t.Root.AddChild(New TreeNode)
		child.Name = "Users"
		
		Set child = child.AddChild(New TreeNode)
		child.Name = "username"
	'Act:
		Dim actual As String
		actual = t.Root.Children(1).Path
	'Assert:
	Assert.AreEqual expected, actual

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub LeafPathToStringWithOptionalSeparator()
	On Error GoTo TestFail
	
	'Arrange:
		Const expected As String = "C:/Users/username/test.txt"
		
		Dim child As TreeNode
		Set child = t.Root.AddChild(New TreeNode)
		child.Name = "Users"
		
		Set child = child.AddChild(New TreeNode)
		child.Name = "username"
		
		Set child = child.AddChild(New TreeNode)
		child.Name = "test.txt"
	'Act:
		Dim actual As String
		actual = child.Path("/")
	'Assert:
	Assert.AreEqual expected, actual

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub NodePathToStringWithOptionalSeparator()
	On Error GoTo TestFail
	
	'Arrange:
		Const expected As String = "C:/Users/"
		
		Dim child As TreeNode
		Set child = t.Root.AddChild(New TreeNode)
		child.Name = "Users"
		
		Set child = child.AddChild(New TreeNode)
		child.Name = "username"
	'Act:
		Dim actual As String
		actual = t.Root.Children(1).Path("/")
	'Assert:
	Assert.AreEqual expected, actual

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod
Public Sub AddingANodeToSecondParentCopiesNode()
	On Error GoTo TestFail
	
	'Arrange:
		Dim parent1 As TreeNode
		Dim parent2 As TreeNode
		
		Set parent1 = t.Root.AddNewChild("parent 1")
		Set parent2 = t.Root.AddNewChild("parent 2")
		
		Dim child As New TreeNode
		child.Name = "child"
		
	'Act:
		parent1.AddChild child
		parent2.AddChild child
	'Assert:
	Assert.AreNotSame parent1.Children(1), parent2.Children(1)

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub CanAddChildToTwoParents()
	On Error GoTo TestFail
	
	'Arrange:
		Dim parent1 As TreeNode
		Dim parent2 As TreeNode
		
		Set parent1 = t.Root.AddNewChild("parent 1")
		Set parent2 = t.Root.AddNewChild("parent 2")
		
		Dim child As New TreeNode
		child.Name = "child"
		
	'Act:
		parent1.AddChild child
		parent2.AddChild child

	'Assert:
	Assert.AreSame parent1, parent1.Children(1).Parent
	Assert.AreSame parent2, parent2.Children(1).Parent

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub CanAddObjectToValue()
	On Error GoTo TestFail
	
	'Arrange:
		Dim expected As New Collection
	'Act:
		Set t.Root.Value = expected
	'Assert:
	Assert.AreSame expected, t.Root.Value

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub CanAddValueToValue()
	On Error GoTo TestFail
	
	'Arrange:
		Const expected As Integer = 42
	'Act:
		t.Root.Value = expected
	'Assert:
	Assert.AreEqual expected, t.Root.Value

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ShallowCopyOfValueValue()
	On Error GoTo TestFail
	
	'Arrange:
		Dim parent1 As TreeNode
		Dim parent2 As TreeNode
		
		Set parent1 = t.Root.AddNewChild("parent 1")
		Set parent2 = t.Root.AddNewChild("parent 2")
		
		Dim child As New TreeNode
		child.Name = "child"
		Const expected As Integer = 42
		child.Value = expected
	'Act:
		parent1.AddChild child
		parent2.AddChild child
		
	'Assert:
	Assert.AreNotSame parent1.Children(1), parent2.Children(1)
	Assert.AreEqual parent1.Children(1).Value, parent2.Children(1).Value

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub ShallowCopyOfObjectValue()
	On Error GoTo TestFail
	
	'Arrange:
		Dim parent1 As TreeNode
		Dim parent2 As TreeNode
		
		Set parent1 = t.Root.AddNewChild("parent 1")
		Set parent2 = t.Root.AddNewChild("parent 2")
		
		Dim child As New TreeNode
		child.Name = "child"
		Dim expected As New Collection
		Set child.Value = expected
	'Act:
		parent1.AddChild child
		parent2.AddChild child
		
	'Assert:
	Assert.AreNotSame parent1.Children(1), parent2.Children(1)
	Assert.AreSame parent1.Children(1).Value, parent2.Children(1).Value

TestExit:
	Exit Sub
TestFail:
	Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


