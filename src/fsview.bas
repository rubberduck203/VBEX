Attribute VB_Name = "fsview"
Option Explicit
'
' fsview
' ======
'
' Introspect the file system.
' 1. Path exists
' 2. Sub Items of path
' 3. Recursive Find
' 4. Glob search (Only uses VB `?` and `*` wild cards)
'
' Copyright (c) 2014 Philip Wales
' This file (fsview.bas) is distributed under the MIT license.
' Obtain a copy of the license here: http://opensource.org/licenses/MIT

Private Const ALLPAT As String = "*"
Public Const PARDIR As String = ".."
Public Const CURDIR As String = "."
'
' Introspect FileSystem
' ---------------------
''
' returns whether file or folder exists or not.
' Use `vbType` argument to filter/include files.
' See <http://msdn.microsoft.com/en-us/library/dk008ty4(v=vs.90).aspx>
' for more types
Public Function Exists(ByVal file_path As String, _
        Optional ByVal vbType As Integer = vbDirectory) As Boolean

    If Not file_path = vbNullString Then
    
        Exists = Not (Dir$(path.RTrimSep(file_path), vbType) = vbNullString)
        
    End If
    
End Function
''
' Will not return true if a folder exists of the same name
Public Function fileExists(ByVal file_path As String)

    fileExists = Exists(file_path, vbNormal)
    
End Function
''
' vbDirectory option still includes files.
' FML
Public Function FolderExists(ByVal folder_path As String)

    FolderExists = Exists(folder_path, vbDirectory) _
                   And Not Exists(folder_path, vbNormal)
    
End Function
''
' returns a List of strings that are paths of subitems in root which
' match pat.
Public Function SubItems(ByVal root As String, Optional ByVal pat As String = ALLPAT, _
        Optional ByVal vbType As Integer = vbDirectory) As List
                  
    Set SubItems = List.create
    
    Dim sub_item As String
    sub_item = Dir$(JoinPath(root, pat), vbType)
    
    While sub_item <> vbNullString
    
        SubItems.Append JoinPath(root, sub_item)
        sub_item = Dir$()
        
    Wend
    
End Function
Public Function SubFiles(ByVal root As String, _
        Optional pat As String = ALLPAT) As List

    Set SubFiles = SubItems(root, pat, vbNormal)
    
End Function
''
' Why on earth would I want . and .. included in sub folders?
' When vbDirectory is passed to dir it still includes files.  Why the would
' anyone want that?  Now there is no direct way to actually list subfolders
' only get a list of both files and folders and filter out files
Public Function SubFolders(ByVal root As String, Optional ByVal pat As String = vbNullString, _
        Optional ByVal skipDots As Boolean = True) As List
                    
    Dim result As List
    Set result = SubItems(root, pat, vbDirectory)
    
    If skipDots And result.Count > 0 Then

        If result(1) = JoinPath(root, CURDIR) Then ' else root
            result.Remove 1
            If result(1) = JoinPath(root, PARDIR) Then  ' else mountpoint
                result.Remove 1
            End If
        End If
        
    End If
    
    ' filter method!
    Dim i As Long
    For i = result.Count To 1 Step -1
        If fileExists(result(i)) Then
            result.Remove i
        End If
    Next i
    
    Set SubFolders = result
    
End Function
Public Function Find(ByVal root As String, Optional ByVal pat As String = "*", _
        Optional ByVal vbType As Integer = vbNormal) As List

    Set Find = List.create
    
    FindRecurse root, Find, pat, vbType
    
End Function
Private Sub FindRecurse(ByVal root As String, ByRef Items As List, _
        Optional pat As String = "*", Optional ByVal vbType As Integer = vbNormal)
    
    Dim folder As Variant
    For Each folder In SubFolders(root)
        FindRecurse folder, Items, pat, vbType
    Next folder
    
    Items.Extend SubItems(root, pat, vbType)
    
End Sub
Public Function Glob(ByVal pattern As String, Optional ByVal vbType As Integer = vbNormal) As List
    
    Dim root As String
    root = LongestRoot(pattern)
    
    Dim patterns() As String
    patterns = Split(right$(pattern, Len(pattern) - Len(root) - 1), SEP)
    
    Set Glob = GlobRecurse(root, patterns, 0, vbType)
    
End Function
Private Function GlobRecurse(ByVal root As String, ByRef patterns() As String, _
        ByVal index As Integer, ByVal vbType As Integer) As List
   
    If index = UBound(patterns) Then
        Set GlobRecurse = SubItems(root, patterns(index), vbType)
    Else
        
        Set GlobRecurse = List.create
        
        Dim folder As Variant
        For Each folder In SubFolders(root, patterns(index))
            GlobRecurse.Extend GlobRecurse(folder, patterns, index + 1, vbType)
        Next folder
        
    End If
    
End Function
