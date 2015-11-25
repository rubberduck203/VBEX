Attribute VB_Name = "Exceptions"
Option Explicit

Public Enum vbErrorNums
    TYPE_ERROR = 13
    OBJECT_REQUIRED = 424
    INDEX_ERROR = 9
    VALUE_ERROR = 380
    ' TODO: list more values
End Enum

Public Enum exErrorNums
    TYPE_ERROR = vbErrorNums.TYPE_ERROR
    OBJECT_REQUIRED = vbErrorNums.OBJECT_REQUIRED
    INDEX_ERROR = vbErrorNums.INDEX_ERROR
    VALUE_ERROR = vbErrorNums.VALUE_ERROR
    UNIMPLEMENTED = 1 ' TODO: use more non-conflicting values
    ILLEGAL_ARGUMENT = 666
    KEY_ERROR
    OS_ERROR
End Enum
'
' Exceptions
' ==========
'
Public Sub BubbleError(ByVal raiser, ByVal method As String, _
        ByVal e As ErrObject)

    Dim trace As String
    trace = MakeDescription(raiser, method, e.description)

    Err.Raise e.Number, e.source, trace ', e.HelpFile, e.HelpContext

End Sub
Public Sub IllegalArgument(ByVal raiser, ByVal method As String, _
        Optional ByVal msg As String)

    RaiseError exErrorNums.ILLEGAL_ARGUMENT, raiser, method, msg

End Sub
Public Sub IndexError(ByVal raiser, ByVal method As String, _
        Optional ByVal msg As String)
    
    RaiseError exErrorNums.INDEX_ERROR, raiser, method, msg
 
End Sub
Public Sub KeyError(ByVal raiser, ByVal method As String, _
        Optional ByVal msg As String)
        
    RaiseError exErrorNums.KEY_ERROR, raiser, method, msg
    
End Sub
Public Sub NotImplementedError(ByVal raiser, ByVal method As String)

    Dim source As String
    source = MakeSource(raiser, method)
    
    Dim msg As String
    msg = source & " Not implemented."

    RaiseError exErrorNums.UNIMPLEMENTED, raiser, method, msg
    
End Sub
Public Sub OSError(ByVal raiser, ByVal method As String, _
        Optional ByVal msg As String)
        
    RaiseError exErrorNums.OS_ERROR, raiser, method, msg
    
End Sub
Public Sub TypeError(ByVal raiser, ByVal method As String, _
        Optional ByVal msg As String)
        
    RaiseError exErrorNums.TYPE_ERROR, raiser, method, msg
    
End Sub
Public Sub ValueError(ByVal raiser, ByVal method As String, _
        Optional ByVal msg As String)
        
     RaiseError exErrorNums.VALUE_ERROR, raiser, method, msg
    
End Sub
'
' Private Methods
' ---------------
'
Private Sub RaiseError(ByVal errNum As exErrorNums, ByVal raiser, _
        ByVal method As String, ByVal msg As String)

    Err.Raise errNum, description:=MakeDescription(raiser, method, msg)

End Sub
Private Function MakeDescription(ByVal raiser, ByVal method As String, _
        ByVal msg As String) As String
    
    MakeDescription = AddTrace(MakeSource(raiser, method), msg)
    
End Function
Private Function MakeSource(ByVal raiser, ByVal method As String) As String

    Dim result As String
    If IsObject(raiser) Then
        result = TypeName(raiser) & "." & method
    Else
        result = raiser & "." & method
    End If

    MakeSource = result

End Function
Private Function AddTrace(ByVal source As String, _
        ByVal description As String) As String

    AddTrace = source & " >> " & description

End Function
