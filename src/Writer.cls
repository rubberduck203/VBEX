VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Writer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
' File Writer
' ===========
'
' author: Matt's Mug
' url: http://codereview.stackexchange.com/questions/52306/file-reader-writer-text
'
'
Option Explicit
Implements IOutput

'
' Private Members
' ---------------
'
Private Const ERROR_BASE As Long = &HFADE
Private pFilePath As String
Private pFileNumber As Long
'
' Public Memebers
' ---------------
'
'expose raised errors to clients:
' TODO: not exposed yet Public?
Public Enum FileWriterError
    FileNotOpened = vbObjectError + ERROR_BASE + 42
    FileAlreadyOpened
    FileDoesNotExist
    FileAlreadyExists
End Enum
'
' Constructor
' -----------
'
Public Function OpenFile(ByVal filePath As String, _
        Optional ByVal overwrite As Boolean = True) As Writer

    Dim result As New Writer
    result.SetFile filePath, False, overwrite
    Set OpenFile = result
    
End Function
Public Function NewFile(ByVal filePath As String) As Writer

    Dim result As New Writer
    result.SetFile filePath, True, False
    Set NewFile = result
    
End Function
'
' ### Friend Methods
'
Friend Sub SetFile(ByVal filePath As String, ByVal create As Boolean, _
        ByVal overwrite As Boolean)
    
On Error GoTo ErrHandler

    If pFileNumber <> 0 Then
        OnFileAlreadyOpenedError "SetFile", filePath
    End If
    
    Dim filePathExists As Boolean
    filePathExists = fsview.fileExists(filePath)
    
    If create And filePathExists Then
        OnFileAlreadyExistsError "SetFile", filePath
    ElseIf Not (create Or filePathExists) Then
        OnFileDoesNotExistError "SetFile", filePath
    End If
    
    pFileNumber = FreeFile
    pFilePath = filePath

    If overwrite Or create Then
        Open filePath For Output As #pFileNumber
    Else
        Open filePath For Append As #pFileNumber
    End If

Exit Sub

ErrHandler:
    OnUnHandledError "SetFile"

End Sub
'
' Destructors
' -----------
'
Private Sub Class_Terminate()

    CloseFile
    
End Sub
'
' Public Methods
' --------------
'
Public Sub CloseFile()
On Error GoTo ErrHandler

    Close #pFileNumber
    pFileNumber = 0

Exit Sub

ErrHandler:
    OnUnHandledError "CloseFile"
        
End Sub
'
' Private Methods
' ---------------
'
' ### Writing
'
Private Sub AppendData(ByVal data As String, ByVal method As String)
On Error GoTo ErrHandler

    If pFileNumber = 0 Then
        OnFileNotOpenedError method
    End If

    Print #pFileNumber, data;

CleanExit:
    Exit Sub

ErrHandler:

    'handle "52: Bad file name or number" by raising a FileWriterError.FileNotOpened instead:
    If Err.Number = 52 Then OnFileNotOpenedError method

    'close file it *any* error occurs writing to it:
    CloseFile

    OnUnHandledError "AppendData"
    
End Sub
'
' ### Errors
'
Private Function GetErrorSource(ByVal method As String) As String

    GetErrorSource = TypeName(Me) & "." & method
    
End Function
Private Sub OnUnHandledError(ByVal method As String)

    Err.Raise Err.Number, GetErrorSource(method), Err.Description, _
        Err.HelpFile, Err.HelpContext
        
End Sub
Private Sub OnFileNotOpenedError(ByVal method As String)

    Err.Raise FileWriterError.FileNotOpened, GetErrorSource(method), _
        "File #" & pFileNumber & "(" & pFilePath & ") was unexpectedly closed."
        
End Sub
Private Sub OnFileAlreadyOpenedError(ByVal method As String, ByVal fileName As String)

    Err.Raise FileWriterError.FileAlreadyOpened, GetErrorSource(method), _
        "File '" & fileName & _
        "' cannot be opened with this instance at this point. A file is already opened."
        
End Sub
Private Sub OnFileDoesNotExistError(ByVal method As String, ByVal fileName As String)

    Err.Raise FileWriterError.FileDoesNotExist, GetErrorSource(method), _
        "File '" & fileName & "' cannot be opened as it does not exist."
        
End Sub
Private Sub OnFileAlreadyExistsError(ByVal method As String, ByVal fileName As String)

    Err.Raise FileWriterError.FileAlreadyExists, GetErrorSource(method), _
        "File '" & fileName & "' cannot be created as it already exists."
        
End Sub
'
' Interfaces
' ----------
'
' ### IOutput
'
'
' #### Prints
'
Public Sub Prints(ByVal data As Variant)

    AppendData Printable.ToString(data), method:="Prints"
    
End Sub
Public Sub IOutput_Prints(ByVal data As Variant)

    Prints data

End Sub
'
' #### PrintLine
'
Public Sub PrintLine(ByVal line As Variant)
    
    AppendData Printable.ToString(line), method:="PrintLine"
    AppendData vbNewLine, method:="PrintLine"
    
End Sub
Public Sub IOutput_PrintLine(ByVal line As Variant)

    PrintLine line

End Sub
'
' #### PrintLiness
'
Public Sub PrintLines(ByRef lines() As Variant)
    
    Dim line As Variant
    For Each line In lines
        PrintLine line
    Next line
    
End Sub
Public Sub IOutput_PrintLines(ByRef lines() As Variant)

    PrintLine lines

End Sub



