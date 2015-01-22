Attribute VB_Name = "DevTools"
Option Explicit

''
' Extensibility Library: For Meta stuff
Private Const VBA_EXTENSIBILITY_LIB As String = "C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
Private Const VBA_EXTENSIBILITY_NAME As String = "VBIDE"

''
' Scripting Runtime: For Dictionary
Private Const VBA_SCRIPTING_LIB As String = "C:\Windows\system32\scrrun.dll"
Private Const VBA_SCRIPTING_NAME As String = "C:\Windows\system32\scrrun.dll"


''
' This should mimic vbext_ComponentType.  The idea is that this module
' could operate with as little setup as possible.
'
Private Enum CompenentType

    stdModule = 1
    classModule = 2
    msForm = 3
    activeXDesigner = 11
    document = 100
    
End Enum
'
' Building VBEX
' -------------
'
Public Sub BuildVBEX(ByVal sourceDir As String, ByVal buildDir As String)

    Dim buildPath As String
    buildPath = buildDir & "VBEX.xlam"

    Dim testPath As String
    testPath = buildDir & "VBEX-Testing.xlam"

    BuildAddin sourceDir & "src\", buildPath, "VBEX"
    BuildAddin sourceDir & "test\", testPath, "Testing"
    
    Dim vbexWb As Workbook
    Set vbexWb = Workbooks.Open(buildPath)
    
    AddReference vbexWb.VBProject, VBA_EXTENSIBILITY_NAME, VBA_EXTENSIBILITY_LIB
    AddReference vbexWb.VBProject, VBA_SCRIPTING_NAME, VBA_SCRIPTING_LIB
    
    vbexWb.Close savechanges:=True
    
End Sub
Private Sub BuildAddin(ByVal sourceDir As String, _
        ByVal buildPath As String, ByVal projectName As String)
        
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    Dim prj As Object
    Set prj = wb.VBProject
    prj.name = projectName
    
    ImportSourceFiles prj, sourceDir
    
    wb.SaveAs buildPath, FileFormat:=55
    wb.Close savechanges:=False
    
End Sub
'
' Importing VBA Files
' -------------------
'
''
' `project` is `Object` to avoid dependence
Private Sub ImportSourceFiles(ByVal project As Object, ByVal sourceDir As String)

    Dim file As String
    file = Dir(sourceDir)
    
    While (file <> "")
        project.VBComponents.Import sourcePath & file
        file = Dir
    Wend
    
End Sub
''
'
Private Function HasReference(ByVal project As Object, ByVal refName As String) As Boolean

    Dim ref As Variant
    For Each ref In project.Reference
    
        If ref.name = refName Then
            HasReference = True
            Exit Function
        End If
        
    Next ref
    
    HasReference = False

End Function
''
'
Private Sub AddReference(ByVal project As Object, ByVal refName As String, _
        ByVal dllPath As String)

    If Not HasReference(project, refName) Then
        project.References.AddFromFile dllPath
    End If

End Sub
'
' Exporting VBEX
' --------------
'
' Maybe separate this so that Extensibility Lib can be used
'
Public Sub ExportSourceFiles(ByVal project As Object, ByVal destPath As String)

    Dim component As Variant
    For Each component In project.VBComponents
        
        Dim compType As Integer
        compType = component.Type
    
        If OughtExport(compType) Then
        
            Dim exportPath As String
            exportPath = destPath & component.name & ToFileExtension(component.Type)
            component.Export exportPath
            
        End If
    Next

End Sub
Private Function OughtExport(ByVal compType As CompenentType) As Boolean

    OughtExport = ((compType = stdModule) Or (compType = classModule))
    
End Function
Private Function ToFileExtension(ByVal compType As CompenentType) As String

    Select Case compType
        Case CompenentType.classModule
            ToFileExtension = ".cls"
            
        Case CompenentType.stdModule
            ToFileExtension = ".bas"
            
        Case CompenentType.msForm
            ToFileExtension = ".frm"
            
        Case CompenentType.activeXDesigner
            '?
            
        Case CompenentType.document
            '?
            
        Case Else
            ToFileExtension = vbNullString
    End Select

End Function
'
'Public Sub RemoveAllModules()
'    Dim project As VBProject
'    Set project = Application.VBE.ActiveVBProject
'
'    Dim comp As VBComponent
'    For Each comp In project.VBComponents
'        If Not comp.Name = "DevTools" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
'            project.VBComponents.Remove comp
'        End If
'    Next
'End Sub
'

