VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'
' Lambda
' ======
'
' Create lambda expressions from strings using
' either C# notation
'
'       "(x, y, z) => x * y + z"
'
' Or Scala short hand
'
'      "( _ * _ ) + _"
'
' both produce the same result
'
' RESTRICTIONS
' ------------
'
'   Lambdas currently must be completely independent
' of external definitions.  One cannot reference
' external varaibles or functions from a lambda.  This
' is because Lambda simple translates the strings to
' a proper function defintion that are written to an
' external blank module.  That module is reference by
' the client code but not visa versa.
'
' _eg_ both of these are invalid
'
'     Public Function Intersect(ByVal other As SortedSet) As SortedSet
'        Set Intersect = Filter(Lambda.FromString("(x) => other.Exists(x)")
'     End Function
'
' `other` is not decalared in the scope of the created lambda function
'
'     Public Function Foo() As Integer
'         Foo = 12
'     End Function
'     Public Sub Bar()
'         Debug.Print List.Create(1,2,3).Map(Lambda.FromShortHand("_ + Foo()")).ToString
'     End Sub
'
' `Foo` is not declared in the scope of the lambda.
'
'  Currently, the external module is located inside of
' the VBEX project.  So it should be able to reference
' VBEX methods.  So this is valid.
'
'    Private Function BaseNames(ByVal filepaths As List) As List
'        Set BaseNames = filepaths.Map(Lambda.FromString("(f) => path.BaseName(f)"))
'    End Function
'
'
' TODO:
' -----
'
'  - Generate code on class init and delete on destruction
'    + Don't delete other lambdas
'  - Delegate as just reference to client code procedure
'  - pass local variables to the lambda.
'
Option Explicit
Implements IPrintable
Implements IApplicable
'
' Constants
' ---------
'
Private Const DEFAULT_LAMBDA_ID As String = "LambdaFunction"
Private Const CANNOT_FIND_MACRO As Integer = 1004
Private Const INVALID_LAMBDA_EXPRESSION As Integer = 666 ' TODO:
'Private Const LAMBDA_PROJECT As String = "LambdasProject" 'Naming convention?
Private Const LAMBDA_MODULE As String = "LambdasModule" '?
Private Const SHORT_HAND_CHAR As String = "_" ' TODO: it's acutal name
'
' Private Members
' ---------------
'
Private pId As String ' name of the lamda
Private pExpression As String ' expression
Private pParameters() As String ' name of the args
'
' Constructors
' ------------
'
''
' Create an Lambda Function
Public Function Create(ByRef params() As String, ByVal expr As String) As Lambda
    
    Dim result As New Lambda
    result.Assign DEFAULT_LAMBDA_ID, params, expr
    Set Create = result
    
End Function
''
' Initializtion wrapper.  Everything is needed at once,  we cannot do them individually
Friend Sub Assign(ByVal id As String, ByRef params() As String, ByVal expr As String)
    
    pExpression = expr
    pId = id ' TODO: Autogenerate so we can have multiple lambdas in registry.
    pParameters = params 'TODO: VALIDATE these, All alpha characters, and trim them
    
End Sub
Public Function FromString(ByVal lambdaExpression As String) As Lambda
    
    Set FromString = FromProper(lambdaExpression)
     
End Function
Private Function FromProper(ByVal proper As String) As Lambda

    Dim splitExpr() As String
    splitExpr = Split(proper, ") => ")
    
    Dim args As String
    args = Mid$(splitExpr(0), 2) ' skip "("
    
    Dim expr As String
    expr = splitExpr(1)
    
    Dim params() As String
    params = Split(Replace$(args, " ", ""), ",")
    Set FromProper = Lambda.Create(params, expr)
    
End Function
Public Function FromShortHand(ByVal shorthand As String) As Lambda

    Dim aCode As Integer
    aCode = Asc("a")
    
    Dim expr As String
    expr = shorthand
    
    Dim paramSize As Integer
    paramSize = UBound(Split(expr, SHORT_HAND_CHAR)) - 1
    
    Dim params() As String
    ReDim params(0 To paramSize)
    
    Dim i As Integer
    For i = 0 To paramSize
    
        Dim paramChar As String
        paramChar = Chr(aCode + i)
        
        expr = Replace(expr, SHORT_HAND_CHAR, paramChar, Count:=1)
        
        params(i) = paramChar
        
    Next i
    
    Set FromShortHand = Create(params, expr)
    
End Function
'
' Usage
' -----
'
Public Function Exec(ParamArray params() As Variant) As Variant
Attribute Exec.VB_UserMemId = 0
    
    cast.Assign Exec, Apply(cast.CArray(params))
    
End Function
Public Function Apply(ByRef args() As Variant) As Variant
    
    Dim result As Variant
    On Error GoTo ErrHandler
    GenerateFunction
    Select Case UBound(args) + 1
        Case 0
            cast.Assign result, Application.Run(pId)
        Case 1
            cast.Assign result, Application.Run(pId, args(0))
        Case 2
            cast.Assign result, Application.Run(pId, args(0), args(1))
        Case 3
            cast.Assign result, Application.Run(pId, args(0), args(1), args(2))
        Case 4
            cast.Assign result, Application.Run(pId, args(0), args(1), args(2), args(3))
        Case 5
            cast.Assign result, Application.Run(pId, args(0), args(1), args(2), args(3), args(4))
        Case 6
            cast.Assign result, Application.Run(pId, args(0), args(1), args(2), args(3), args(4), args(5))
        Case 7
            cast.Assign result, Application.Run(pId, args(0), args(1), args(2), args(3), args(4), args(5), args(6))
        Case 8
            cast.Assign result, Application.Run(pId, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case 9
            cast.Assign result, Application.Run(pId, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
        Case 10
            cast.Assign result, Application.Run(pId, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
        Case Else
            Err.Raise 5, TypeName(Me) & ".Apply", "Too many parameters. Max is 10" ' too many parameters errors
    End Select

    cast.Assign Apply, result
    
CleanExit:
    DeleteFunction
    Exit Function
    
ErrHandler:
    
    If Err.Number = CANNOT_FIND_MACRO Then ' more descriptive later
        Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    Resume CleanExit
    
End Function
'
' Public Properties
' -----------------
'
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
Private Property Get ParameterCount() As String
    
    ParameterCount = UBound(pParameters) - LBound(pParameters) + 1
    
End Property
Private Property Get Parameters() As String
    
    Parameters = Join(pParameters, ", ")
    
End Property
Private Property Get Signature() As String
    
    Signature = "Public Function " & pId & "(" & Parameters & ") As Variant"
    
End Property
Private Property Get Content() As String

    Content = Join(Array("", Signature, Body, "End Function", ""), vbNewLine)
    
End Property
Private Property Get Expression() As String
    
    Expression = pExpression
    
End Property
Private Property Get Body() As String
    
    Body = vbTab & pId & " = " & pExpression
    
End Property
Private Property Get LambdaCodeModule() As CodeModule
    
     Set LambdaCodeModule = ThisWorkbook.VBProject.VBComponents(LAMBDA_MODULE).CodeModule
     
End Property
'
' Error Raising
' -------------
'
'
' Interfaces
' ----------
'
' ### IApplicable
'
Public Function IApplicable_Apply(ByRef args() As Variant) As Variant

    cast.Assign IApplicable_Apply, Apply(args)

End Function
Public Function IApplicable_Exec(ParamArray args() As Variant) As Variant

<<<<<<< HEAD
    cast.Assign IApplicable_Exec, Apply(cast.CArray(args))
=======
    IApplicable_Exec = Apply(cast.CArr(args))
>>>>>>> origin/master

End Function
'
' ### IPrintable
'
Public Property Get ToString() As String
    
    ToString = TypeName(Me) & "[(" & Parameters & ") => " & Expression & "]"
    
End Property
Public Property Get IPrintable_ToString() As String

    IPrintable_ToString = ToString()
    
End Property
