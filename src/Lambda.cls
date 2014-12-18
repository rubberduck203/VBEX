VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Lambda
' ========
'
'
' Constants
' ---------
'
Private Const DEFAULT_LAMBDA_ID As String = "LambdaFunction"
Private Const CANNOT_FIND_MACRO As Integer = 1004
Private Const LAMBDA_PROJECT As String = "LambdasProject" 'Naming convention?
Private Const LAMBDA_MODULE As String = "LambdasModule" '?
'
' Private Members
' ---------------
'
Private pId As String ' name of the lamda
Private pExpression As String ' expression
Private pParameters() As String ' name of the
'
' Constructors
' ------------
'
''
' Create an Lambda Function
Public Function Create(ByRef Parameters() As String, ByVal Expression As String) As Lambda
    Dim anon As New Lambda
    anon.Assign Parameters, Expression
    Set Create = anon
End Function
Public Function FromStr(ByVal Expression As String) As Lambda
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "\(([a-zA-Z-9_]*)\)\s\=\>\s(.*)" ' Doesn't accept multiple arguements

    If Not regex.test(Expression) Then
        Err.Raise 5, "Lambda", "Invalid Lambda function expression."
    End If

    Dim regexMatch As Variant
    Set regexMatch = regex.Execute(Expression)(0)
    ' Nope, do not like. should be
    If regexMatch.SubMatches(0) = vbNullString Then
        Set FromStr = ArgLessLambda(regexMatch.SubMatches(1))
    Else
        Set FromStr = ArgedLambda(regexMatch.SubMatches(0), regexMatch.SubMatches(1))
    End If

End Function
'
' #### Private Constructors?
'
' I don't like these and they should leave my sight.  But I believe we
' need to first remove the regex dependency.
Private Function ArgLessLambda(ByVal expr As String) As Lambda
    
    Dim params() As String
    Dim result As New Lambda
    result.Assign DEFAULT_LAMBDA_ID, Right(expr, Len(expr) - Len("() => ")), params
    Set ArgLessLambda = result
    
End Function
Private Function ArgedLambda(ByVal args As String, ByVal expr As String) As Lambda
    
    Dim params() As String
    params = Split(Trim(args), ",")
    
    Dim result As New Lambda
    result.Assign DEFAULT_LAMBDA_ID, expr, params
    Set ArgedLambda = result
    
End Function
'
' Usage
' -----
'
Public Function Exec(ParamArray params() As Variant) As Variant
    Exec = Apply(cast.CArray(params))
End Function
Public Function Apply(ByRef args() As Variant) As Variant
    
    Dim result As Variant
    On Error GoTo ErrHandler
    GenerateFunction
    Select Case UBound(args) + 1
        Case 0
            result = Application.Run(pId)
        Case 1
            result = Application.Run(pId, args(0))
        Case 2
            result = Application.Run(pId, args(0), args(1))
        Case 3
            result = Application.Run(pId, args(0), args(1), args(2))
        Case 4
            result = Application.Run(pId, args(0), args(1), args(2), args(3))
        Case 5
            result = Application.Run(pId, args(0), args(1), args(2), args(3), args(4))
        Case 6
            result = Application.Run(pId, args(0), args(1), args(2), args(3), args(4), args(5))
        Case 7
            result = Application.Run(pId, args(0), args(1), args(2), args(3), args(4), args(5), args(6))
        Case 8
            result = Application.Run(pId, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case 9
            result = Application.Run(pId, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
        Case 10
            result = Application.Run(pId, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
        Case Else
            Err.Raise 5, TypeName(Me) & ".Apply", "Too many parameters. Max is 10" ' too many parameters errors
    End Select

    Apply = result
CleanExit:
    DeleteFunction
    Exit Function

ErrHandler:

    If Err.number = CANNOT_FIND_MACRO Then ' more descriptive later
        Err.Raise Err.number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    Resume CleanExit
    
End Function
'
' Private Methods
' ---------------
'
' ### Code Manipulation
'
Private Sub GenerateFunction()
    With LambdaCodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString Content
    End With
End Sub
Private Sub DeleteFunction()
    With LambdaCodeModule
        .DeleteLines 1, .CountOfLines
    End With
End Sub
'
' ### Private Properties
'
''
' Initializtion wrapper.  Everything is needed at once,  we cannot do them individually
Friend Sub Assign(ByVal id As String, ByVal expr As String, params() As String)
    pExpression = expr
    pId = id
    pParameters = params
End Sub

Public Property Get Parameters() As String
    Parameters = Join(pParameters, ", ")
End Property
Public Property Get Signature() As String
    Signature = "Public Function " & pId & "(" & Parameters & ") As Variant"
End Property
Public Property Get Content() As String
    Content = Join(Array("", Signature, Body, "End Function", ""), vbNewLine)
End Property
Public Property Get Expression() As String
    Expression = pExpression
End Property
Public Property Get Body() As String
    Body = vbTab & pId & " = " & pExpression
End Property
Private Property Get LambdaCodeModule() As CodeModule
    Set LambdaCodeModule = Application.VBE.VBProjects(LAMBDA_PROJECT).VBComponents(LAMBDA_MODULE).CodeModule
End Property
Public Property Get ToString() As String
    ToString = "(" & Parameters & ") => " & Expression
End Property
'
' ### Error Raising
'
