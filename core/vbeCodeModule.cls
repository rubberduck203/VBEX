VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeCodeModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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


Public Sub Initialize(codeMod As CodeModule)
    Set mCodeModule = codeMod
    Set mVbeProcedures = getProcedures(mCodeModule)
End Sub

Private Sub Class_Terminate()
    Set mVbeProcedures = Nothing
    Set mCodeModule = Nothing
End Sub

Private Function getProcedures(codeMod As CodeModule) As vbeProcedures
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Returns collection of all vbeProcedures in a CodeModule            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim procName As String
    Dim lastProcName As String
    Dim procs As New vbeProcedures
    Dim proc As vbeProcedure
    
    Dim i As Long
    ' Skip past any Option statement
    '   and any module-level variable declations.

    For i = codeMod.CountOfDeclarationLines + 1 To codeMod.CountOfLines
        ' get procedure name
        procName = codeMod.ProcOfLine(i, vbext_pk_Proc)
        If Not procName = lastProcName Then
            ' create new procedure object
            Set proc = New vbeProcedure
            proc.Initialize procName, codeMod
            ' add it to collection
            procs.Add proc
            ' reset lastProcName
            lastProcName = procName
        End If
    Next i
    Set getProcedures = procs
    
End Function
