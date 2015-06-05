Attribute VB_Name = "defApply"
Option Explicit

Public Function Compose(ByVal outer As IApplicable, ByVal inner As IApplicable) As IApplicable

    Set Compose = Composed.Make(outer, inner)

End Function
Public Function AndThen(ByVal inner As IApplicable, ByVal outer As IApplicable) As IApplicable

    Set AndThen = Composed.Make(outer, inner)

End Function
Public Function ApplicationRunOnArray(ByVal id As String, ByRef args() As Variant) As Variant

    Dim result
    Select Case UBound(args) + 1
        Case 0
            Assign result, Application.Run(id)
        Case 1
            Assign result, Application.Run(id, args(0))
        Case 2
            Assign result, Application.Run(id, args(0), args(1))
        Case 3
            Assign result, Application.Run(id, args(0), args(1), args(2))
        Case 4
            Assign result, Application.Run(id, args(0), args(1), args(2), args(3))
        Case 5
            Assign result, Application.Run(id, args(0), args(1), args(2), args(3), args(4))
        Case 6
            Assign result, Application.Run(id, args(0), args(1), args(2), args(3), args(4), args(5))
        Case 7
            Assign result, Application.Run(id, args(0), args(1), args(2), args(3), args(4), args(5), args(6))
        Case 8
            Assign result, Application.Run(id, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case 9
            Assign result, Application.Run(id, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
        Case 10
            Assign result, Application.Run(id, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
        Case Else
            NotImplementedError "defApply", "ApplicationRunOnArray"
    End Select

    Assign ApplicationRunOnArray, result

End Function
Public Function CallByNameOnArray(ByVal obj As Object, ByVal method As String, ByVal clltype As VbCallType, ByRef args() As Variant) As Variant

    Dim result
    Select Case UBound(args) + 1
        Case 0
            Assign result, CallByName(obj, method, clltype)
        Case 1
            Assign result, CallByName(obj, method, clltype, args(0))
        Case 2
            Assign result, CallByName(obj, method, clltype, args(0), args(1))
        Case 3
            Assign result, CallByName(obj, method, clltype, args(0), args(1), args(2))
        Case 4
            Assign result, CallByName(obj, method, clltype, args(0), args(1), args(2), args(3))
        Case 5
            Assign result, CallByName(obj, method, clltype, args(0), args(1), args(2), args(3), args(4))
        Case 6
            Assign result, CallByName(obj, method, clltype, args(0), args(1), args(2), args(3), args(4), args(5))
        Case 7
            Assign result, CallByName(obj, method, clltype, args(0), args(1), args(2), args(3), args(4), args(5), args(6))
        Case 8
            Assign result, CallByName(obj, method, clltype, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case 9
            Assign result, CallByName(obj, method, clltype, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
        Case 10
            Assign result, CallByName(obj, method, clltype, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
        Case Else
            NotImplementedError "defApply", "CallByNameOnArray"
    End Select

    Assign CallByNameOnArray, result

End Function

