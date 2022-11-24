VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsMonthlyTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public dbMonthlyTotals As New ADODB.Connection
Public rsMonthlyTotals As New ADODB.RecordSet

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
GetDb = "DailyTotals"
End Function

Public Function GetDbFile() As String
GetDbFile = "ReportsDB"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbMonthlyTotals
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsMonthlyTotals
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








Public Sub Update()
pUpdate
End Sub

Private Sub pUpdate()

Dim DailyTotals As New Dictionary
Set DailyTotals = GetCurrentDayTotals
Dim iMon As New aclsMonthlyTotals
AddNewRecord iMon, DailyTotals


End Sub

Private Function GetCurrentDayTotals() As Dictionary
Dim iIndex As New zclsDailyCheckIndex
Dim iDataObj As New aclsDataObject
Set iDataObj = Wrap(GetNewMatchObj)
Dim rs As ADODB.RecordSet
Set rs = GetRecordsetMatch(iDataObj, ConstructMatchQuery(iDataObj))
Dim fld As Object
Dim val As Variant
Dim dict As New Dictionary
For Each fld In rs.Fields
    If fld.Type = 5 Then
        val = FilteredMatch(iIndex.Wrap(GetNewMatchObj("NOT " & fld.name, 0)), "Sum(" & fld.name & ")")(1)(0, 0)
        dict.Add fld.name, val
    End If
Next fld
iDataObj.CloseDbs iDataObj
dict.Add "TotalsDate", Format(Now, "mmddyy")
Set iIndex = Nothing
Set iDataObj = Nothing
Set rs = Nothing
Set fld = Nothing
Set GetCurrentDayTotals = dict
End Function
