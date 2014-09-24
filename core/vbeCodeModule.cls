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
    For lineNumber = 1 To codeMod.CountOfLines
        line = codeMod.Lines(lineNumber, 1)
        If IsSignature(line) Then
            procKind = GetProcedureType(line)
            procName = codeMod.ProcOfLine(lineNumber, procKind)
            
            Set proc = New vbeProcedure
            proc.Initialize procName, codeMod, procKind
        End If
    Next lineNumber

    Set GetProcedures = procs
End Function

Private Function GetProcedureType(signatureLine As String) As vbext_ProcKind
    If InStr(1, signatureLine, "Property Get") > 0 Then
        GetProcedureType = vbext_pk_Get
    ElseIf InStr(1, signatureLine, "Property Let") > 0 Then
        GetProcedureType = vbext_pk_Let
    ElseIf InStr(1, signatureLine, "Property Set") > 0 Then
        GetProcedureType = vbext_pk_Set
    ElseIf InStr(1, signatureLine, "Sub") > 0 Or InStr(1, signatureLine, "Function") > 0 Then
        GetProcedureType = vbext_pk_Proc
    Else
        Const InvalidProcedureCallOrArgument As Long = 5
        Err.Raise InvalidProcedureCallOrArgument
    End If
End Function

Private Function IsSignature(line As String) As Boolean
    If line = vbNullString Then Exit Function
    
    Dim propertyPosition As Long
    Dim functionPosition As Long
    Dim subPosition As Long
    Dim commentPosition As Long
    
    propertyPosition = InStr(1, line, "Property")
    functionPosition = InStr(1, line, "Function")
    subPosition = InStr(1, line, "Sub")
    commentPosition = InStr(1, line, "'")
    
    If propertyPosition > 0 Or functionPosition > 0 Or subPosition > 0 Then
        If InStr(1, line, "End") = 0 Then
            If commentPosition > propertyPosition Then
                Exit Function
            ElseIf commentPosition > functionPosition Then
                Exit Function
            ElseIf commentPosition > subPosition Then
                Exit Function
            Else
                IsSignature = True
            End If
        End If
    End If
    
End Function
