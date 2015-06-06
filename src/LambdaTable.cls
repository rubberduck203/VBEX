VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LambdaTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Constants
' ---------
'
Private Const DEFAULT_LAMBDA_ID As String = "x___"
Private Const CANNOT_FIND_MACRO As Integer = 1004
Private Const LAMBDA_MODULE As String = "LambdasModule" '?
'
' Private Members
' ---------------
'
Private pLambdaCount As Long
'
' Class
' -----
'
Private Sub Class_Initialize()
    pLambdaCount = 0
    ClearModule
End Sub
Private Sub Class_Terminate()
    ClearModule
End Sub
Private Sub ClearModule()
    
    With LambdaCodeModule
        .DeleteLines 1, .CountOfLines
    End With
    
    pLambdaCount = 0
    
End Sub

'
' Add lambda to table return lambda name
'
Public Function AddLambda(ByRef args() As String, _
        ByVal expression As String) As String

    Dim id As String
    id = GenerateLambdaId
    
    Dim funCode As String
    funCode = BuildFunction(id, expression, args)
    
    LambdaCodeModule.AddFromString funCode
    
    AddLambda = id

End Function
Public Sub RemoveLambda(ByVal id As String)

    With LambdaCodeModule
        
        Dim lineStart As Long
        lineStart = .ProcStartLine(id, vbext_pk_Proc)
        
        Dim lineCount As Long
        lineCount = .ProcCountLines(id, vbext_pk_Proc)
        
        .DeleteLines lineStart, lineCount

    End With

End Sub
Private Function GenerateLambdaId() As String
    
    pLambdaCount = pLambdaCount + 1
    GenerateLambdaId = DEFAULT_LAMBDA_ID & Hex(pLambdaCount)

End Function
'
' ### Code Building
'
Private Function BuildFunction(ByVal id As String, ByVal exp As String, ByRef args() As String) As String

    Dim bod As String
    bod = Body(id, exp)
    
    Dim parms As String
    parms = Parameters(args)
    
    Dim sig As String
    sig = Signature(id, parms)

    BuildFunction = Content(sig, bod)

End Function
Private Function Content(ByVal sig As String, ByVal bod As String) As String

    Content = Join(Array(sig, bod, "End Function"), vbNewLine)
    
End Function
Private Function Signature(ByVal id As String, ByVal parms As String) As String
    
    Signature = "Public Function " & id & "(" & parms & ") As Variant"
    
End Function
Private Function Body(ByVal id As String, ByVal exp As String) As String
    
    Body = "cast.Assign " & id & ", " & exp
    
End Function
Private Function Parameters(ByRef args() As String) As String
    
    Parameters = Join(args, ", ")
    
End Function
'
' ### VBE access
'
Private Property Get LambdaCodeModule() As CodeModule
    
     Set LambdaCodeModule = ThisWorkbook.VBProject.VBComponents(LAMBDA_MODULE).CodeModule
     
End Property
