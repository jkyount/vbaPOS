VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public dbTable As New ADODB.Connection
Public rsTable As New ADODB.RecordSet
Private pTable As String
Private pParentTable As String
Private pServerName As String
Private pHasOpenCheck As Boolean
Private zclsOrder_pCheck As String
Private pSplit As Boolean
Private pChecks As New Collection
Private pServerNum As Variant
Private pTableCheck As String


Public Property Get Db() As String
Db = GetDb
End Property
Public Property Let Checks(value As Collection)
Set pChecks = value
End Property
Public Property Get Checks() As Collection
Set Checks = GetChecks
End Property
Public Property Let ServerNum(value As Variant)
pServerNum = value
End Property
Public Property Get ServerNum() As Variant
ServerNum = Me.GetTableParam("ServerNum")
End Property
Public Property Let Table(value As String)
pTable = value
End Property
Public Property Get Table() As String
Table = pTable
End Property
Public Property Let ParentTable(value As String)
pParentTable = value
End Property
Public Property Get ParentTable() As String
ParentTable = pParentTable
End Property
Public Property Let ServerName(value As String)
pServerName = value
End Property
Public Property Get ServerName() As String
ServerName = Me.GetTableParam("ServerName")
End Property
Public Property Let HasOpenCheck(value As Boolean)
pHasOpenCheck = value
End Property
Public Property Get HasOpenCheck() As Boolean
HasOpenCheck = Me.GetTableParam("HasOpenCheck")
End Property

'==========================================================================
'==========================================================================
Public Function Wrap(obj As aclsDataObject) As aclsDataObject
Dim iDataObj As New aclsDataObject
Set iDataObj = obj
iDataObj.rs = Me.GetRs
iDataObj.Conn = Me.GetConn
iDataObj.Db = Me.GetDb
iDataObj.DbFile = Me.GetDbFile
Set Wrap = iDataObj
Set iDataObj = Nothing
End Function



Public Function GetDb() As String
GetDb = "TableStates"
End Function
Public Function GetDbFile() As String
GetDbFile = "TableStates"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbTable
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsTable
End Function

Public Sub OpenDbs()
Dim iDataObject As New aclsDataObject
Set iDataObject = Wrap(iDataObject)
iDataObject.OpenDbs iDataObject
Set iDataObject = Nothing
End Sub

Public Sub CloseDbs()
Dim iDataObject As New aclsDataObject
Set iDataObject = Wrap(iDataObject)
iDataObject.CloseDbs iDataObject
Set iDataObject = Nothing
End Sub


'==========================================================================








'==========================================================================
Public Sub Assign(check As String, ServerNum As Integer, ServerName As String)
Me.Table = GetNextTable(Me.ParentTable)
Me.AssignCheck check
Me.AssignServer ServerNum, ServerName
End Sub

Public Function GetTableParam(param As String) As Variant
OpenDbs
rsTable.Source = "SELECT * FROM TableStates WHERE Table = """ & Me.Table & """"
rsTable.Open
GetTableParam = rsTable.Fields(param).value
rsTable.Close
dbTable.Close
End Function

Public Sub SetTableInUse(check As String, state As String)
pSetTableInUse check, state
End Sub

Private Sub pSetTableInUse(check As String, state As String)
OpenDbs
rsTable.Source = "SELECT * From TableStates WHERE CheckNumber = """ & check & """"
rsTable.Open
rsTable.Fields("InUse").value = state
rsTable.Fields("CheckNumber").value = check
rsTable.Update
rsTable.Close
dbTable.Close
End Sub

Public Function GetTablesInUse() As Collection
Set GetTablesInUse = pGetTablesInUse
End Function

Private Function pGetTablesInUse() As Collection
Dim coll As New Collection
Dim x As String
OpenDbs
rsTable.Source = "SELECT * From TableStates WHERE InUse = True ORDER BY Table"
rsTable.Open
Do Until rsTable.EOF
    If Not rsTable.Fields("ParentTable") = x Then
        x = rsTable.Fields("ParentTable")
        coll.Add x
    End If
    rsTable.MoveNext
Loop
Set pGetTablesInUse = coll
rsTable.Close
dbTable.Close
End Function

Public Function GetServerTables(ServerNum As Integer) As Collection
Set GetServerTables = pGetServerTables(ServerNum)
End Function
Private Function pGetServerTables(ServerNum As Integer) As Collection
Dim coll As New Collection
Dim x As String
OpenDbs
rsTable.Source = "SELECT * From TableStates WHERE InUse = True AND ServerNum = " & ServerNum & " ORDER BY Table"
rsTable.Open
Do Until rsTable.EOF
    If Not rsTable.Fields("ParentTable") = x Then
        x = rsTable.Fields("ParentTable")
        coll.Add x
    End If
    rsTable.MoveNext
Loop
Set pGetServerTables = coll
rsTable.Close
dbTable.Close
End Function


Public Function GetNextTable(ParentTable As String) As String
GetNextTable = pGetNextTable(ParentTable)
End Function
Private Function pGetNextTable(ParentTable As String) As String
OpenDbs
rsTable.Source = "SELECT * FROM TableStates WHERE ParentTable = """ & ParentTable & """ ORDER BY LEN(Table), Table ASC"
rsTable.Open
Do Until rsTable.Fields("InUse").value = False Or rsTable.EOF
    rsTable.MoveNext
Loop

If rsTable.EOF = True Then
    MsgBox "Could not assign check to table"
    Exit Function
End If
pGetNextTable = rsTable.Fields("Table").value
rsTable.Close
dbTable.Close
End Function

Public Sub AssignCheck(check As String)
pAssignCheck check
End Sub
Private Sub pAssignCheck(check As String)
OpenDbs
rsTable.Source = "SELECT * FROM TableStates WHERE Table = """ & Me.Table & """"
rsTable.Open
rsTable.Fields("CheckNumber").value = check
rsTable.Fields("InUse").value = True
rsTable.Update
rsTable.Close
dbTable.Close
End Sub

Public Sub UnassignCheck(check As String)
pUnassignCheck check
End Sub
Private Sub pUnassignCheck(check As String)
OpenDbs
rsTable.Source = "SELECT * FROM TableStates WHERE CheckNumber = """ & check & """"
rsTable.Open
rsTable.Fields("CheckNumber").value = ""
rsTable.Fields("InUse").value = False
rsTable.Fields("ServerName").value = ""
rsTable.Fields("ServerNum").value = 0
rsTable.Update
rsTable.Close
dbTable.Close
End Sub

Public Sub AssignServer(ServerNum As Integer, ServerName As String)
pAssignServer ServerNum, ServerName
End Sub
Private Sub pAssignServer(ServerNum As Integer, ServerName As String)
OpenDbs
rsTable.Source = "SELECT * FROM TableStates WHERE Table = """ & Me.Table & """"
rsTable.Open
rsTable.Fields("ServerNum").value = ServerNum
rsTable.Fields("ServerName").value = ServerName
rsTable.Update
rsTable.Close
dbTable.Close
End Sub

Private Function GetChecks() As Collection
Dim coll As New Collection
OpenDbs
rsTable.Source = "SELECT * FROM TableStates WHERE ParentTable = """ & Me.ParentTable & """ AND NOT CheckNumber = """""
rsTable.Open
Do Until rsTable.EOF
    coll.Add rsTable.Fields("CheckNumber").value
    rsTable.MoveNext
Loop
rsTable.Close
dbTable.Close
Set GetChecks = coll
Set coll = Nothing
End Function

Public Sub ClearTableStates()
pClearTableStates
End Sub
Private Sub pClearTableStates()
OpenDbs
rsTable.Source = "TableStates"
rsTable.Open
Do Until rsTable.EOF
    rsTable.Fields("ServerName").value = ""
    rsTable.Fields("CheckNumber").value = ""
    rsTable.Fields("ServerNum").value = 0
    rsTable.Fields("InUse").value = False
    rsTable.MoveNext
Loop
rsTable.UpdateBatch
CloseDbs
End Sub
Public Sub RecallTableState(check As String)
Dim dict As New Dictionary
Dim iDataObject As New aclsDataObject
Set iDataObject = Wrap(GetNewMatchObj("CheckNumber", check))
Set dict = CDictCollection(GetRecordsetMatch(iDataObject, ConstructMatchQuery(iDataObject)))(1)

SetTableState dict("ParentTable"), dict("Table")

Set dict = Nothing
Set iDataObject = Nothing

End Sub

