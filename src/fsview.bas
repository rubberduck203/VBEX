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
' This file (fsview.bas) is distributed under the GPL-3.0 license.
' Obtain a copy of the license here: http://opensource.org/licenses/GPL-3.0

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
Public Function Exists(ByVal filePath As String, _
        Optional ByVal vbType As VbFileAttribute = vbDirectory) As Boolean

    If Not filePath = vbNullString Then
    
        Exists = Not (Dir$(Path.RTrimSep(filePath), vbType) = vbNullString)
        
    End If
    
End Function
''
' Will not return true if a folder exists of the same name
Public Function FileExists(ByVal filePath As String)

    FileExists = Exists(filePath, vbNormal)
    
End Function
''
' vbDirectory option still includes files.
' FML
Public Function FolderExists(ByVal folderPath As String)

    FolderExists = Exists(folderPath, vbDirectory) _
                   And Not Exists(folderPath, vbNormal)
    
End Function
''
' returns a List of strings that are paths of subitems in root which
' match pat.
Public Function SubItems(ByVal Root As String, Optional ByVal pat As String = ALLPAT, _
        Optional ByVal vbType As VbFileAttribute = vbDirectory) As List
                  
    Set SubItems = List.Create
    
    Dim subItem As String
    subItem = Dir$(JoinPath(Root, pat), vbType)

    Do While subItem <> vbNullString
    
        SubItems.Append JoinPath(Root, subItem)
        subItem = Dir$()
        
    Loop
    
End Function
Public Function SubFiles(ByVal Root As String, _
        Optional pat As String = ALLPAT) As List

    Set SubFiles = SubItems(Root, pat, vbNormal)
    
End Function
''
' Why on earth would I want . and .. included in sub folders?
' When vbDirectory is passed to dir it still includes files.  Why the would
' anyone want that?  Now there is no direct way to actually list subfolders
' only get a list of both files and folders and filter out files
Public Function SubFolders(ByVal Root As String, Optional ByVal pat As String = vbNullString, _
        Optional ByVal skipDots As Boolean = True) As List
                    
    Dim result As List
    Set result = SubItems(Root, pat, vbDirectory)
    
    If skipDots And result.Count > 0 Then

        If result(1) = JoinPath(Root, CURDIR) Then ' else root
            result.Remove 1
            If result(1) = JoinPath(Root, PARDIR) Then  ' else mountpoint
                result.Remove 1
            End If
        End If
        
    End If

    Set SubFolders = result.FilterNot(InternalDelegate.Make("FileExists"))
    
End Function
Public Function Find(ByVal Root As String, Optional ByVal pat As String = "*", _
        Optional ByVal vbType As VbFileAttribute = vbNormal) As List

    Dim result As List
    Set result = List.Create
    
    FindRecurse Root, result, pat, vbType
    
    Set Find = result
    
End Function
Private Sub FindRecurse(ByVal Root As String, ByRef foundItems As List, _
        Optional pat As String = "*", _
        Optional ByVal vbType As VbFileAttribute = vbNormal)
    
    Dim folder As Variant
    For Each folder In SubFolders(Root)
        FindRecurse folder, foundItems, pat, vbType
    Next folder
    
    foundItems.Extend SubItems(Root, pat, vbType)
    
End Sub
Public Function Glob(ByVal pattern As String, _
    Optional ByVal vbType As VbFileAttribute = vbNormal) As List
    
    Dim Root As String
    Root = Path.LongestRoot(pattern)
    
    Dim patterns() As String
    patterns = Split(right$(pattern, Len(pattern) - Len(Root) - 1), Path.SEP)
    
    Set Glob = GlobRecurse(Root, patterns, 0, vbType)
    
End Function
Private Function GlobRecurse(ByVal Root As String, ByRef patterns() As String, _
        ByVal index As Integer, ByVal vbType As VbFileAttribute) As List
    
    Dim result As List
    
    If index = UBound(patterns) Then
        Set result = SubItems(Root, patterns(index), vbType)
    Else
        
        Set result = List.Create
        
        Dim folder As Variant
        For Each folder In SubFolders(Root, patterns(index))
            result.Extend GlobRecurse(folder, patterns, index + 1, vbType)
        Next folder
        
    End If
    
    Set GlobRecurse = result
    
End Function
