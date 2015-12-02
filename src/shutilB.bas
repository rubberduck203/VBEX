Attribute VB_Name = "shutilB"
Option Explicit
'
' shutilB
' =======
'
' Advanced filesystem operations for VBA. Routines of this variant will return `true` if
' operation is successful or `false` if failed.
'
' Copyright (c) 2014 Philip Wales
' This file (shutilB.bas) is distributed under the GPL-3.0 license.
' Obtain a copy of the license here: http://opensource.org/licenses/GPL-3.0
'
' It simply calls the respective `shutilE` method and uses boiler plate
' code to evaluate succes.  Until procedures can be treated as data types
' better this boiler-plate code will remain.
'
Public Function Move(ByVal src As String, ByVal dest As String, _
        Optional createParent As Boolean = False) As Boolean
              
    Dim noError As Boolean
    On Error GoTo ErrHandler

    shutilE.Move src, dest, createParent
    noError = True
    
CleanExit:
    Move = noError
    Exit Function
  
ErrHandler:
    Err.Clear
    noError = False
    Resume CleanExit
    
End Function
Public Function Rename(ByVal aPath As String, ByVal newName As String) As Boolean
    
    Dim noError As Boolean
    On Error GoTo ErrHandler

    shutilE.Rename aPath, newName
    noError = True
    
CleanExit:
    Rename = noError
    Exit Function
  
ErrHandler:
    Err.Clear
    noError = False
    Resume CleanExit
    
End Function
Public Function Remove(ByVal filePath As String) As Boolean
    
    Dim noError As Boolean
    On Error GoTo ErrHandler

    shutilE.Remove filePath
    noError = True
    
CleanExit:
    Remove = noError
    Exit Function
  
ErrHandler:
    Err.Clear
    noError = False
    Resume CleanExit
    
End Function
Public Function MakeDir(ByVal filePath As String, _
        Optional createParent As Boolean = False) As Boolean
    
    Dim noError As Boolean
    On Error GoTo ErrHandler

    shutilE.MakeDir filePath, createParent
    noError = True
    
CleanExit:
    MakeDir = noError
    Exit Function
  
ErrHandler:
    Err.Clear
    noError = False
    Resume CleanExit
    
End Function
Public Function CopyFile(ByVal src As String, ByVal dest As String, _
        Optional createParent As Boolean = False) As Boolean
    
    Dim noError As Boolean
    On Error GoTo ErrHandler

    shutilE.CopyFile src, dest, createParent
    noError = True
    
CleanExit:
    CopyFile = noError
    Exit Function
  
ErrHandler:
    Err.Clear
    noError = False
    Resume CleanExit
    
End Function

