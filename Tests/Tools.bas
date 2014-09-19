Attribute VB_Name = "Tools"
Option Explicit

Private Const ProjectName = "VBEXTests"

Public Sub NewUnitTest()
    VBAUnit.UnitTestModule.Add VBE.VBProjects(ProjectName)
End Sub

Public Sub NewTestMethod(methodName As String)
    VBAUnit.UnitTestModule.AddNewTestMethod VBE.VBProjects(ProjectName), methodName
End Sub



