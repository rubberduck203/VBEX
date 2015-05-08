Attribute VB_Name = "defApply"
Option Explicit

Public Function Compose(ByVal f As IApplicable, ByVal g As IApplicable) As IApplicable

    Set Compose = Composed.Make(f, g)

End Function
Public Function AndThen(ByVal f As IApplicable, ByVal g As IApplicable) As IApplicable

    Set Compose = Composed.Make(g, f)

End Function

