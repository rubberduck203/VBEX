Attribute VB_Name = "Tools"
Option Explicit

Private Const ProjectName = "VBEXTests"

Public Sub NewUnitTest()
    VBAUnit.UnitTestModule.Add Application.VBE.VBProjects(ProjectName)
End Sub

Public Sub NewTestMethod(methodName As String)
    VBAUnit.UnitTestModule.AddNewTestMethod methodName
End Sub


