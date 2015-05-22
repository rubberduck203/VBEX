Attribute VB_Name = "defApply"
Option Explicit

Public Function Compose(ByVal outer As IApplicable, ByVal inner As IApplicable) As IApplicable

    Set Compose = Composed.Make(outer, inner)

End Function
Public Function AndThen(ByVal inner As IApplicable, ByVal outer As IApplicable) As IApplicable

    Set AndThen = Composed.Make(outer, inner)

End Function

