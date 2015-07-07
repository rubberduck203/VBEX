VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lazy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Lazy Class
' ============
'
'
' Copywrite (C) 2014 Philip Wales
' This file (Lazy.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
Implements IShowable
Implements IDelayed
'
' Private Members
' ---------------
'
Private pDelayed As IDelayed
Private pValue As Variant
'
' Constructors
' ------------
'
Public Function Create(ByVal op As IApplicable, ParamArray args()) As Lazy

    Set Create = Make(op, Tuple.Implode(CArray(args)))

End Function
Public Function Make(ByVal op As IApplicable, ByVal args As Tuple) As Lazy
    
    Dim delayed As ByName
    Set delayed = ByName.Make(op, args)

    Set Make = FromDelayed(delayed)

End Function
Public Function FromDelayed(ByVal delayed As IDelayed) As Lazy

    Dim result As New Lazy
    Set result.DelayedOp = delayed
    Set FromDelayed = result
    
End Function
Friend Property Set DelayedOp(ByVal delayed As IDelayed)

    Set pDelayed = delayed

End Property
'
' Public Methods
' --------------
'
Public Function IsDelayed() As Boolean

    IsDelayed = IsEmpty(pValue)

End Function
Public Function IsEvaluated() As Boolean

    IsEvaluated = (Not IsDelayed)

End Function
'
' IDelayed
' --------
'
Public Property Get Evaluate() As Variant
Attribute Evaluate.VB_UserMemId = 0

    If IsDelayed Then
        Assign pValue, pDelayed.Evaluate
    End If

    Assign Evaluate, pValue

End Property
Private Property Get IDelayed_Evaluate() As Variant

    Assign IDelayed_Evaluate, Evaluate

End Property
Public Function Map(ByVal op As IApplicable) As Lazy

    Dim result As IDelayed
    If IsDelayed Then
        'Set result = ByName.Create(op, Me) ' Evaluates `Me` before returning result
        'Set result = ByName.Make(op, Tuple.Pack(Me)) ' Same result: Parameter Array + Me => default property evaluated.
        'Set result = ByName.Make(op, Tuple.Implode(Array(Me))) ' Same result
        Set result = ByName.Make(op, Tuple.SingleValue(Me)) ' Evaluating result does not evaluate `Me` when using a CallByName IApplicable.
        Debug.Assert IsDelayed
    Else
        Set result = ByName.Create(op, pValue)
    End If
    
    Set Map = Lazy.FromDelayed(result)

End Function
Private Function IDelayed_Map(ByVal op As IApplicable) As IDelayed

    Set IDelayed_Map = Map(op)

End Function
'
' IShowable
' ---------
'
Public Function Show() As String

    Dim repr As String
    If IsDelayed Then
        repr = defShow.Show(pDelayed)
    Else
        repr = defShow.Show(pValue)
    End If
    Show = defShow.ParamShowableObject(Me, repr)

End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
