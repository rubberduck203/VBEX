Attribute VB_Name = "shutilE"
'
' shutilE
' =======
'
' Advanced filesystem operations for VBA. This variant raises errors if 
' attempted operation fails.
'
' Copyright (c) 2014 Philip Wales
' This file (shutilE.bas) is distributed under the MIT license.
' Obtain a copy of the license here: http://opensource.org/licenses/MIT
'
' Scripting.FileSystemObject is slow and unstable since it relies on sending
' signals to ActiveX objects across the system.  This module only uses built-in
' functions of Visual Basic, such as `Dir`, `Kill`, `Name`, etc.

'
'
' File System Modifications
' -------------------------
'
'
Public Sub Move(ByVal src_path As String, ByVal dest_path As String, _
        Optional create_parent As Boolean = False)

    On Error GoTo ErrHandler

    DestIsFolderFeature dest_path, src_path
    
    If create_parent Then CreateRootPath dest_path
    
    Name src_path As dest_path
    If Not Exists(dest_path) Then Error_FailedCreate "Move", "Name As"
    If Exists(src_path) Then Error_FailedDestroy "Move", "Name As"
    
CleanExit:
    Exit Function
  
ErrHandler:
    Select Case Err.Number
    Case Else
        Err.Raise Err.Number
    End Select

End Function
Public Sub Remove(ByVal file_path As String)
    On Error GoTo ErrHandler
    
    Kill file_path
    If Exists(dest_path) Then Error_FailedDestroy "Remove", "Kill"
    
CleanExit:
    Exit Function

ErrHandler:
    Select Case Err.Number
    Case Else
        Err.Raise Err.Number
    End Select
    
End Function
Public Sub MakeDir(ByVal folder_path As String, ByVal Optional create_parent As Boolean = False)
    
    Dim check As Boolean
    On Error GoTo ErrHandler
        
    If create_parent Then CreateRootPath folder_path
    MkDir folder_path
    
    If Not FolderExists(dest_path) Then Error_FailedCreate "MakeDir", "MkDir"
    
CleanExit:
    Exit Function
    
ErrHandler:
    Select Case Err.Number
    Case Else
        Err.Raise Err.Number
    End Select
    
End Function
Public Sub CopyFile(ByVal src_path As String, ByVal dest_path As String, _
      Optional create_parent As Boolean = False)
    
    On Error GoTo ErrHandler
    
    DestIsFolderFeature dest_path, src_path
    If FileExists(dest_path) Then Error_NoOverwrite "CopyFile"
    
    If create_parent Then CreateRootPath dest_path
    FileCopy src_path, dest_path
    
    If Not FileExists(dest_path) Then Error_FailedCreate "CopyFile", "FileCopy"

CleanExit:
   Exit Function

ErrHandler:
    Select Case Err.Number
    Case Else
       Err.Raise Err.number, GetErrorSource(method), Err.Description, Err.HelpFile, Err.HelpContext
    End Select
    
End Function
Private Sub CreateRootPath(ByVal path As String)
    
    Dim parent_folder As String
    parent_folder = RootName(path)
    
    If Not FolderExists(parent_folder) Then
    
        MakeDir parent_folder, create_parent:=True
        
    End If
    
End Function
Private Sub DestIsFolderFeature(ByRef dest_path As String, _
        ByVal src_path As String)
    
    If right$(dest_path, 1) = SEP Or FolderExists(dest_path) Then 
        dest_path = pJoin(dest_path, BaseName(src_path))
    End If
    
End Sub
'
' ### Custom Error Messages
'
Private Sub Error_FailedCreate(Byval method as String, ByVal operation As String)
    Err.Raise osErrNums.unknown, method, "Destination does not exist after errorless `" & operation &"`"
End Sub
Private Sub Error_FailedDestroy(ByVal method As String, ByVal operation As String)
    Err.Raise osErrNums.unknown, method, "Destination still exists after errorless `" & operation &"`"
End Sub
Private Sub Error_NoOverwrite(ByVal method As String)
    Err.Raise osErrNums.overwriteRefusal, method, "Will not overwrite file at destination.  Remove it first if desired."
End Sub
