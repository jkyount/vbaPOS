Attribute VB_Name = "i_Table"
Option Explicit


Public Sub SetTableInUse(check As String, state As String)
Dim iTable As New zclsTable
iTable.SetTableInUse check, state
Set iTable = Nothing
End Sub

Public Function GetTablesInUse() As Collection
Dim iTable As New zclsTable
Set GetTablesInUse = iTable.GetTablesInUse
Set iTable = Nothing
End Function

Public Function GetServerTables(ServerNum As Integer) As Collection
Dim iTable As New zclsTable
Set GetServerTables = iTable.GetServerTables(ServerNum)
Set iTable = Nothing
End Function

Public Function GetNextTable(ParentTable As String) As String
Dim iTable As New zclsTable
GetNextTable = iTable.GetNextTable(ParentTable)
Set iTable = Nothing
End Function

Public Sub ClearTableStates()
Dim iTable As New zclsTable
iTable.ClearTableStates
Set iTable = Nothing
End Sub

Public Sub SetTableState(ParentTable As String, Table As String)
ThisTable.ParentTable = ParentTable
ThisTable.Table = Table
End Sub


Public Sub RecallTableState(check As String)
Dim iTable As New zclsTable
iTable.RecallTableState check
Set iTable = Nothing
End Sub

Public Sub UnassignCheck(check As String)
Dim iTable As New zclsTable
iTable.Table = ThisTable.Table
iTable.UnassignCheck check
Set iTable = Nothing
End Sub

