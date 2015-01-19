Attribute VB_Name = "shutilB"
'
' shutilB
' =======
'
' Advanced filesystem operations for VBA. Routines of this variant will return `true` if
' operation is successful or `false` if failed.
'
' Copyright (c) 2014 Philip Wales
' This file (shutilB.bas) is distributed under the MIT license.
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
Function Move(ByVal src_path As String, ByVal dest_path As String, _
        Optional create_parent As Boolean = False) As Boolean
              
    Dim check As Boolean
    On Error GoTo ErrHandler

    DestIsFolderFeature dest_path, src_path
    
    If create_parent Then CreateRootPath dest_path
    
    Name src_path As dest_path
    check = Exists(dest_path)
    
CleanExit:
    Move = check
    Exit Function
  
ErrHandler:
    Err.Clear
    Debug.Assert (Not check)
    Resume CleanExit
    
End Function
Function Rename(ByVal path As String, ByVal new_name As String) As Boolean
    
    Debug.Assert BaseName(new_name) = new_name
    
    Rename = Move(path, pJoin(RootName(path), new_name))

End Function
Function Remove(ByVal file_path As String) As Boolean
    
    Dim check As Boolean
    On Error GoTo ErrHandler
    
    Kill file_path
    check = (Not FileExists(file_path))
    
CleanExit:
    Remove = check
    Exit Function

ErrHandler:
    Err.Clear
    Debug.Assert (Not check)
    Resume CleanExit

End Function
Function MakeDir(ByVal folder_path As String, _
        Optional create_parent As Boolean = False) As Boolean
                
    Dim check As Boolean
    On Error GoTo ErrHandler
        
    If create_parent Then CreateRootPath folder_path
    MkDir folder_path
    check = FolderExists(folder_path)
    
CleanExit:
    MakeDir = check
    Exit Function
    
ErrHandler:
    Err.Clear
    Debug.Assert (Not check)
    Resume CleanExit
    
End Function
Function CopyFile(ByVal src_path As String, ByVal dest_path As String, _
        Optional create_parent As Boolean = False) As Boolean
    
    Dim check As Boolean
    On Error GoTo ErrHandler
    
    DestIsFolderFeature dest_path, src_path
    
    If FileExists(dest_path) Then GoTo CleanExit:
    
    If create_parent Then CreateRootPath dest_path
    FileCopy src_path, dest_path
    check = FileExists(dest_path)

CleanExit:
    CopyFile = check
    Exit Function
    
ErrHandler:
    Err.Clear
    Debug.Assert (Not check)
    Resume CleanExit
    
End Function
Private Function CreateRootPath(ByVal path As String) As Boolean

    Dim parent_folder As String
    parent_folder = RootName(path)
    
    If Not FolderExists(parent_folder) Then
    
        CreateRootPath = MakeDir(parent_folder, create_parent:=True)
        
    End If
    
End Function
Private Sub DestIsFolderFeature(ByRef dest_path As String, ByVal src_path As String)
    If right$(dest_path, 1) = SEP Or FolderExists(dest_path) Then
        dest_path = pJoin(dest_path, BaseName(src_path))
    End If
End Sub
