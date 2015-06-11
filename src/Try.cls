VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Try"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Try Class
' ===========
'
'
' Copywrite (C) 2014 Philip Wales
' This file (Try.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
' An object that optionally contains an item.  It's use is an abstraction from
' checking if a method returned a value or not.
'
'
Implements IEquatable
Implements IShowable
Implements IMonadic
'
' Private Members
' ---------------
'
Private pResult As Variant
Private pError As ErrObject
Private pSuccess As Boolean
'
' Constructor
' -----------
'
Public Function Make(ByVal attempt As IDelayed) As Try

    Dim result As Try
    On Error GoTo Fail
    Set result = Success(attempt.Evaluate())
    On Error GoTo 0

CleanExit:
    Set Try = result

Exit Function

Fail:
    Set result = Failure(Err)
    Resume CleanExit

End Function
Private Function Success(ByVal x) As Try
    
    Dim result As New Try
    result.AssignResult x
    Set Some = result
    
End Function
Private Function Failure(ByVal e As ErrObject) As Try

    Dim result As New Try
    result.AssingError e
    Set Failure = result

End Function
'
' ### Friend Methods
'
Friend Sub AssignResult(ByVal x)

    pSuccess = True
    Assign pResult, x

End Sub
Friend Sub AssignError(ByVal e As ErrObject)

    pSuccess = False
    Set pError = e

End Sub
'
' Public Methods
' --------------
'
Public Function IsSucces() As Boolean
    
    IsSuccess = pSuccess
    
End Function
Public Function IsFailure() As Boolean

    IsFailure = (Not IsSuccess)

End Function
Public Function GetItem() As Variant
    
    If IsSuccess Then
        Assign GetItem, pResult
    Else
        Exceptions.BubbleError Me, "GetItem", pError
    End If

End Function
'
' IEquatable
' ----------
'
Public Function Equals(ByVal x As Try) As Boolean

    If x.IsSuccess And IsSuccess Then
        Equals = (x.GetItem = pResult)
    Else
        Equals = False
    End If

End Function
Private Function IEquatable_Equals(ByVal x) As Boolean

    IEquatable_Equals = Equals(x)

End Function
'
' IMonadic
' --------
'
Public Function Bind(ByVal op As IApplicable) As Try

    Dim result As Try
    If IsSuccess Then
    
        On Error GoTo ErrHandler
        Dim opResult
        Assign opResult, op.Apply(pVar)
        Set result = opResult
        On Error GoTo 0
        
    Else
        Set result = Me
    End If
    
    Set Bind = result
Exit Function
ErrHandler:
    Dim msg As String
    Select Case Err.Number
        Case vbErrorNums.TYPE_ERROR, vbErrorNums.OBJECT_REQUIRED
            msg = defShow.Show(op) & " did not return a Try object"
            Exceptions.TypeError Me, "Bind", msg
        Case Else
            Exceptions.BubbleError Me, "Bind", Err
    End Select
End Function
Public Function Map(ByVal op As IApplicable) As Try

    Dim result As Try
    If IsSuccess Then

        Dim delay As ByName
        Set delay = ByName.Make(op, pResult)
        Set result = Make(delay)

    Else
        Set result = Me
    End If
    Set Map = result
    
End Function
Private Function IMonadic_Bind(ByVal op As IApplicable) As IMonadic

    Set IMonadic_Bind = Bind(op)

End Function
Private Function IMonadic_Map(ByVal op As IApplicable) As IMonadic

    Set IMonadic_Map = Map(op)

End Function
Private Function IMonadic_Unit(ByVal x) As IMonadic

    Set IMonadic_Unit = Make(x)

End Function
'
' IShowable
' ---------
'
Public Function Show() As String

    Dim result As String
    If IsSuccess Then
        result = ParamShowableObject(Me, pResult)
    Else
        result = ParamShowableObject(Me, pError)
    End If

    Show = result
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
