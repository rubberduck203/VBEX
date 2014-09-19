Attribute VB_Name = "DevTools"
Option Explicit

Public Sub ImportSourceFiles(sourcePath As String)
    Dim file As String
    file = Dir(sourcePath)
    While (file <> "")
        Application.VBE.ActiveVBProject.VBComponents.Import sourcePath & file
        file = Dir
    Wend
End Sub

Public Sub ExportSourceFiles(destPath As String)
    
    Dim comp As VBComponent
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        comp.Export destPath & comp.Name & ToFileExtension(comp.Type)
    Next
    
End Sub

Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
        Case Else
            ToFileExtension = vbNullString
    End Select
    
End Function
