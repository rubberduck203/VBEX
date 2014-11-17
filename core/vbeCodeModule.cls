VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeCodeModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' requires Microsoft Visual Basic for Applications Extensibility 5.3 library
Option Explicit

Private mCodeModule As CodeModule
Private mVbeProcedures As vbeProcedures

Public Property Get CodeModule() As CodeModule
    Set CodeModule = mCodeModule
End Property

Public Property Let CodeModule(ByRef codeMod As CodeModule)
    Me.Initialize codeMod
End Property

Public Property Get vbeProcedures()
    Set vbeProcedures = mVbeProcedures
End Property

Public Sub Insert(ComponentType As vbext_ComponentType)
    'Dim project As VBProject
    'Set project = VBIDE.VBE
    'project.VBComponents.Add ComponentType
End Sub

Public Function Create(codeMod As CodeModule) As vbeCodeModule
' allows calls from other projects without breaking the exisiting API

    Set Create = New vbeCodeModule
    Create.Initialize codeMod
End Function

Public Sub Initialize(codeMod As CodeModule)
    Set mCodeModule = codeMod
    Set mVbeProcedures = GetProcedures(mCodeModule)
End Sub

Private Sub Class_Terminate()
    Set mVbeProcedures = Nothing
    Set mCodeModule = Nothing
End Sub

Private Function GetProcedures(codeMod As CodeModule) As vbeProcedures
    Dim procName As String
    Dim procs As New vbeProcedures
    Dim proc As vbeProcedure
    Dim line As String
    Dim procKind As vbext_ProcKind
    
    Dim lineNumber As Long
    lineNumber = codeMod.CountOfDeclarationLines + 1
    While (lineNumber < codeMod.CountOfLines)
        procName = codeMod.ProcOfLine(lineNumber, procKind) 'procKind is an OUT param
    
        Set proc = New vbeProcedure
        proc.Initialize procName, codeMod, procKind
        
        procs.Add proc
    
        ' add current start line to the number of lines in current procedure to get the start line of the next procedure.
        lineNumber = lineNumber + codeMod.ProcCountLines(procName, procKind) + 1
    Wend

    Set GetProcedures = procs
    
End Function
