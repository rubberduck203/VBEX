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
'
' Tree
' ====
'
' A TypedList of type Tree with an additional data member
'
' Private Members
' ---------------
'
Private pData As Variant
''
' Type is Tree Enforced in constructor
Private pBranches As TypedList
'
' Public Members
' --------------
'
' None
'
' Private Methods
' ---------------
'
''
' Set TypeName of branches to Tree
' Each tree has a value and a list of
' other trees
Private Sub Class_Initialize()
    
    Set pBranches = New TypedList
    pBranches.TypedName = TypeName(Me)
    
End Sub
Private Sub Class_Terminate()
    Set pBranches = Nothing
End Sub
'
' Public Methods
' --------------
'
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    
    Set NewEnum = Leaves.NewEnum
    
End Property
'
' ### Access
'
Public Property Get Value() As Variant
    seq.Assign Item, pData
End Property
Public Property Get Branches() As TypedList
    Set Branches = subTrees
End Property
Public Property Get Branch(ByVal index As Long) As Tree
Attribute Branch.VB_UserMemID = 0
    Set Branch = pBranches(index)
End Property
''
' TEST!
' Should return a 2D list only!
Public Function Paths() As TypedList
    
    Set Paths = New TypedList
    Paths.TypedName = "List"
    
    If IsLeaf Then
        ' This branch is a leaf
        Dim leaf As New List
        leaf.Append pData
        Paths.Append leaf
        ' return a 2D List
    Else
        
        Dim arm As Variant
        ' concatenate each branches leaves
        For Each arm In pBranches
            Paths.Extend arm.Paths
        Next arm
        ' emplace our data at the front of each.
        For Each arm In Paths ' This should not create recursion.
            arm.Emplace 1, pData
        Next arm
                
    End If
    
End Function
Public Function Leaves() As List
    
    Set Leaves = New List
    
    If IsLeaf Then
        Leaves.Append pData
    Else
        
        Dim arm As Variant
        For Each arm in pBranches
            Leaves.Extend arm.Leaves
        Next arm
  
    End If

End Function
'
' ### Replace
'
Public Property Let Value(ByVal data As Variant)
    pData = data
End Property
Public Property Set Value(ByVal data As Variant)
    Set pData = data
End Property
Public Property Set Branch(ByVal index As Long, ByVal aTree As Tree)
Attribute Branch.VB_UserMemID = 0
    
    Set pBranches(index) = aTree
    
End Property
'
' ### Aux
'
Public Sub Create(ByVal data As Variant)
    
    Dim newBranch as New Tree
    newBranch.Value = data
    pBranches.Append newBranch
    
End Sub
Public Function IsLeaf() As Boolean
    IsLeaf = (pBranches.Count = 0)
End Function
Public Function BranchCount() As Long

    BranchCount = pBranches.Count

    Dim arm As Variant
    For Each arm in pBranches
        BranchCount = BranchCount + arm.BranchCount
    Next arm

End Function
'
' Inherit TypedList
' -----------------
'
'
' ### Insert
'
Public Sub Emplace(ByVal index As Long, ByVal aTree As Tree)
    pBranches.Emplaceindex, aTree
End Sub
Public Sub Insert(ByVal index As Long, ByVal trees As Variant)
    pBranches.Insert index, trees
End Sub
'
' ### Extend
'
Public Sub Append(ByVal aTree As Tree)
    pBranches.Append aTree
End Sub
Public Sub Extend(ByVal trees As Variant)
    pBranches.Extend trees
End Sub
'
' ### Remove
'
Public Sub Remove(ByVal index As Tree)
    pBranches.Remove index
End Sub
Public Sub Clear(ByVal start As Long, ByVal size As Long)
    pBranches.Remove start, size
End Sub
