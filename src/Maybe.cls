VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Maybe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Maybe Class
' ===========
'
'
' Copywrite (C) 2014 Philip Wales
' This file (Maybe.cls) is distributed under the MIT license
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
Private pVar As Variant
'
' Constructor
' -----------
'
Public Function Some(ByVal x) As Maybe
    
    Dim result As New Maybe
    result.AssignOpt x
    Set Some = result
    
End Function
Public Function None() As Maybe

    Dim result As New Maybe
    Set None = result

End Function
Public Function MakeIf(ByVal pred As Boolean, ByVal x) As Maybe

    If pred Then
        Set MakeIf = Some(x)
    Else
        Set MakeIf = None
    End If
    
End Function
'
' ### Friend Methods
'
Friend Sub AssignOpt(ByVal x)

    Assign pVar, x

End Sub
'
' Public Methods
' --------------
'
Public Function IsNone() As Boolean
    
    IsNone = IsEmpty(pVar)
    
End Function
Public Function IsSome() As Boolean

    IsSome = (Not IsNone)

End Function
Public Function GetItem() As Variant

    If IsSome Then
        Assign GetItem, pVar
    Else
        ValueError Me, "GetItem", "Cannot get None"
    End If
    
End Function
Public Function GetOrElse(ByVal other) As Variant

    If IsSome Then
        Assign GetOrElse, pVar
    Else
        Assign GetOrElse, other
    End If

End Function
'
' IEquatable
' ----------
'
Public Function Equals(ByVal x) As Boolean

    Equals = False
    
    If TypeName(x) <> TypeName(Me) Then
        Exit Function
    End If
    
    Dim mx As Maybe
    Set mx = x
    
    If IsSome And mx.IsSome Then
        Equals = (defEquals.Equals(pVar, mx.GetItem))
    Else
        Equals = (IsNone And mx.IsNone)
    End If

End Function
Private Function IEquatable_Equals(ByVal x) As Boolean

    IEquatable_Equals = Equals(x)

End Function
'
' IMonadic
' --------
'
Public Function Bind(ByVal f As IApplicable) As Maybe

    Dim result As Maybe
    If IsSome Then
        Set result = f.Apply(pVar)
    Else
        Set result = None
    End If
    
    Set Bind = result

End Function
Private Function IMonadic_Bind(ByVal f As IApplicable) As IMonadic

    Set IMonadic_Bind = Bind(f)

End Function
Private Function IMonadic_Unit(ByVal x) As IMonadic

    Set IMonadic_Unit = Some(x)

End Function
'
' IShowable
' ---------
'
Public Function Show() As String

    Dim result As String
    If IsNone Then
        result = ParamShowableObject(Me)
    Else
        result = ParamShowableObject(Me, pVar)
    End If

    Show = result
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
