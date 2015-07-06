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
Private Sub Class_Initialize()

    Set pValue = Maybe.None

End Sub
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

    IsDelayed = pValue.IsNone

End Function
Public Function IsEvaluated() As Boolean

    IsEvaluated = pValue.IsSome

End Function
'
' IDelayed
' --------
'
Public Property Get Evaluate() As Variant
Attribute Evaluate.VB_UserMemId = 0

    If IsDelayed Then
        Set pValue = Maybe.Some(pDelayed.Evaluate)
    End If

    Assign Evaluate, pValue.GetItem

End Property
Private Property Get IDelayed_Evaluate() As Variant

    Assign IDelayed_Evaluate, Evaluate

End Property
Public Function Map(ByVal op As IApplicable) As Lazy

    Dim result As IDelayed
    If IsDelayed Then
        ' result will still need evaluation even after I am evaluated.
        ' the result must have some reference to myself.
        'Set result = pDelayed.Map(op)
        Dim indirect As Lazy
        Set indirect = Me
        Set result = ByName.Create(op, indirect)
    Else
        ' Correctly done
        Set result = ByName.Create(op, pValue.GetItem)
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
    repr = defShow.Show(pValue.GetOrElse(pDelayed))

    Show = defShow.ParamShowableObject(Me, repr)

End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
