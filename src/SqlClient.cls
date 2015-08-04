VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
' SqlClient
' =========
'
' Establishes an ADODB connection to a source and provides a rich interface to
' that source.
'
' Private Member
' --------------
'
Private pConnection As ADODB.Connection
'
' Error Values
' ------------
'
Public Enum SqlClientErrors
    NO_RECORDS = 770
    TOO_MANY_RECORDS
    TOO_MANY_FIELDS
End Enum
'
' Constructor
' -----------
'
' ### Constructors
'
Public Function Make(ByVal connectString As String) As SqlClient
    
    Dim result As New SqlClient
    result.OpenConnection connectString
    Set Make = result
    
End Function
Public Function ToAccess(ByVal filePath As String) As SqlClient
    
    Dim connectString As String
    connectString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & filePath & ";" & _
        "Jet OLEDB:Engine Type=5;" & _
        "Persist Security Info=False;"

    Set ToAccess = Make(connectString)

End Function
Public Function ToExcel(ByVal filePath As String) As SqlClient
    
    Dim connectString As String
    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & filePath & ";" & _
        "Extended Properties=Excel 8.0;"

    Set ToExcel = Make(connectString)

End Function
'
' ### Class Events
'
Private Sub Class_Initialize()
    
    Set pConnection = New ADODB.Connection

End Sub
Private Sub Class_Terminate()
    
    pConnection.Close
    
End Sub
'
' ### Friend Properties
'
Friend Sub OpenConnection(ByVal connectStr As String)
    
    pConnection.Open connectStr
    
End Sub
Friend Property Get DbConnection() As ADODB.Connection
    
    Set DbConnection = pConnection
    
End Property
'
' Public Methods
' --------------
'
' ### Generic Methods
'

Public Function MakeCommand(ByVal query As String) As ADODB.Command
    
    Dim result As New ADODB.Command
    result.ActiveConnection = DbConnection
    result.CommandText = query
    Set MakeCommand = result

End Function
Public Function GetQuery(ByVal queryName As String) As ADODB.Command
    
    Dim result As ADODB.Command
    Set result = MakeCommand(queryName)
    result.CommandType = adCmdStoredProc
    Set GetQuery = result

End Function
Public Function BuildQuery(ByVal queryName As String, ByVal args As Dict) As ADODB.Command
    
    On Error GoTo SimpleBubble
    Dim queryCmd As ADODB.Command
    Set queryCmd = GetQuery(queryName)
    AppendOrderedArgs queryCmd, OrderArgs(args)
    Set BuildQuery = queryCmd
    
Exit Function
SimpleBubble:
    Dim showArgs As Dict
    Set showArgs = Dict.Create( _
        Assoc.Make("queryName", queryName), _
        Assoc.Make("args", args) _
    )
    Exceptions.BubbleError Me, "BuildQuery", Err
    
End Function
Public Function QueryResult(ByVal queryName As String, ByVal args As Dict) As ADODB.Recordset

    On Error GoTo SimpleBubble
    Dim queryCmd As ADODB.Command
    Set queryCmd = BuildQuery(queryName, args)
    Set QueryResult = queryCmd.Execute

Exit Function
SimpleBubble:
    Dim showArgs As Dict
    Set showArgs = Dict.Create( _
        Assoc.Make("queryName", queryName), _
        Assoc.Make("args", args) _
    )
    Exceptions.BubbleError Me, "QueryResult", Err
    
End Function
'
' ### Result Handling
'
' #### Single Results
'
Public Function SimpleSelectLookup(ByVal sqlName As String, ByVal arg As Assoc) As Maybe
    
    On Error GoTo SimpleBubble
    Set SimpleSelectLookup = SingleValueQuery(sqlName, Dict.Create(arg))
    
Exit Function
SimpleBubble:
    Exceptions.BubbleError Me, "SimpleSelectLookup", Err
    
End Function
Public Function SingleValueQuery(ByVal queryName As String, ByVal args As Dict) As Maybe
    
    On Error GoTo SimpleBubble
    Dim sqlResult As ADODB.Recordset
    Set sqlResult = QueryResult(queryName, args)
    Set SingleValueQuery = GetSingleResult(sqlResult)
    
Exit Function
SimpleBubble:
    Dim showArgs As Dict
    Set showArgs = Dict.Create( _
        Assoc.Make("queryName", queryName), _
        Assoc.Make("args", args) _
    )
    Exceptions.BubbleError Me, "SingleValueQuery", Err
    
End Function
'
' #### Dictionary
'
''
' HDict:
Public Function HDict(ByVal recSet As ADODB.Recordset) As Dict
    
    Dim recs() As Variant
    recs = recSet.GetRows
    
    ' Check for only 2 fields
    If UBound(recs, 1) > 1 Then
        Err.Raise TOO_MANY_FIELDS, _
            Description:=TypeName(Me) & ".HDict: Query returned more than 2 fields."
    End If
    
    Dim keys As List
    Set keys = List.Copy(Application.Index(recs, 1, 0))
    
    Dim vals As List
    Set vals = List.Copy(Application.Index(recs, 2, 0))
    
    Dim result As Dict
    Set result = Dict.FromLists(keys, vals)
    
    Set HDict = result
    
End Function
'
' Private Methods
' ---------------
'
' ### Result Handling
'
Private Function GetSingleResult(ByVal recSet As ADODB.Recordset) As Maybe
    
    Dim result As Maybe
    ' Will need a RecordSet wrapper class this is nonesense
    If (recSet.BOF Or recSet.EOF) Then
        Set result = Maybe.None
    Else
    
        Dim recRows() As Variant
        recRows = recSet.GetRows
        
        If UBound(recRows, 2) > 0 Then
            Exceptions.KeyError Me, "GetSingleResult", "Query returned multiple records."
        ElseIf UBound(recRows, 1) > 0 Then
            Exceptions.KeyError Me, "GetSingleResult", "Query returned multiple fields."
        Else
            Set result = Maybe.Some(recRows(0, 0))
        End If
        
    End If
    
    Set GetSingleResult = result

End Function
'
' ### Append safely
'
Private Function SafeAdd(ByVal getArgs As Dict, ByVal getQuerysName As String, _
        ByVal addArgs As Dict, ByVal addQuerysName As String) As Long
    
    Dim autoId As Maybe
    Set autoId = SingleValueQuery(getQuerysName, getArgs)
    
    If autoId.IsNone Then
    
        BuildQuery(addQuerysName, addArgs).Execute
        Set autoId = SingleValueQuery(getQuerysName, getArgs)
        
    End If
    On Error GoTo ErrHandler
    SafeAdd = autoId
Exit Function
ErrHandler:
    Dim showArgs As Dict
    Set showArgs = Dict.Create( _
        Assoc.Make("getArgs", getArgs), _
        Assoc.Make("getQuerysName", getQuerysName), _
        Assoc.Make("addArgs", addArgs), _
        Assoc.Make("addQuerysName", addQuerysName) _
    )
    Exceptions.ValueError Me, "SafeAdd", "Failed to get autoId with " & showArgs.Show
End Function
Private Function SimpleSafeAdd(ByVal arg As Assoc, ByVal getQuerysName As String, _
        ByVal addQuerysName As String) As Long
    
    Dim unifiedArgs As Dict
    Set unifiedArgs = Dict.Create(arg)
    
    SimpleSafeAdd = SafeAdd(unifiedArgs, getQuerysName, unifiedArgs, addQuerysName)
    
End Function
'
' ### Parameter Handling
'
Private Sub AppendArgs(ByVal sqlCmd As ADODB.Command, ByVal args As List)
    
    Dim arg
    For Each arg In args
        With sqlCmd
            .Parameters.Append .CreateParameter(Type:=adVariant, Value:=arg)
        End With
    Next
    
End Sub
''
' TODO: Move to VBEX.Dict somehow.
Private Function OrderArgs(ByVal args As Dict) As List
    
    Dim orderedNames As SortedSet
    Set orderedNames = SortedSet.Copy(args.keys)

    Dim argItemGetter As OnArgs
    Set argItemGetter = OnArgs.Make("Item", VbGet, args)

    Dim result As List
    Set result = Map(List, argItemGetter, orderedNames)

    Set OrderArgs = result

End Function
