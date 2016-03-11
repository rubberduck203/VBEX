Attribute VB_Name = "defMonad"
Option Explicit

Public Function ApplyUnit(ByVal m As Monadic) As OnArgs

    Set ApplyUnit = OnArgs.Make("Unit", VbMethod, m)

End Function
