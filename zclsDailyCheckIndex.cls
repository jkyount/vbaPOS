VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsDailyCheckIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit




Public dbDailyCheckIndex As New ADODB.Connection
Public rsDailyCheckIndex As New ADODB.RecordSet

Public Property Get Db() As String
Db = GetDb
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
iDataObj.Archive = GetArchive
iDataObj.ArchiveDbFile = GetArchiveDbFile
Set Wrap = iDataObj
Set iDataObj = Nothing
End Function


Public Function GetDb() As String
GetDb = "DailyCheckIndex"
End Function

Public Function GetDbFile() As String
GetDbFile = "CheckDb"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbDailyCheckIndex
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsDailyCheckIndex
End Function

Public Function GetArchive() As String
GetArchive = "ArchivedCheckIndex"
End Function
Public Function GetArchiveDbFile() As String
GetArchiveDbFile = "ReportsDb"
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




Public Sub UpdateTax(check As String)
pUpdateTax check
End Sub

Private Sub pUpdateTax(check As String)
Dim Ord As New aclsOrder
Ord.ImportCheckDetails Ord, check
Ord.OrderType = Ord.SameOrderType

Update Wrap(GetNewUpdateObj(, check, "Tax", Ord.OrderType.GetTax(check)))
Set Ord = Nothing
'OpenDbs
'rsDailyCheckIndex.Source = "SELECT * From DailyCheckIndex WHERE CheckNumber = """ & Check & """"
'rsDailyCheckIndex.Open
'If rsDailyCheckIndex.Fields("DineIn").Value = True Then
'    rsDailyCheckIndex.Fields("Tax").Value = GetTax_DineIn(Check)
'End If
'If rsDailyCheckIndex.Fields("DineIn").Value = False Then
'    rsDailyCheckIndex.Fields("Tax").Value = GetTax_Carryout(Check)
'End If
'rsDailyCheckIndex.Update
'rsDailyCheckIndex.Close
'dbDailyCheckIndex.Close
End Sub


Public Sub UpdateTotal(check As String)
pUpdateTotal check
End Sub

Private Sub pUpdateTotal(check As String)
OpenDbs
rsDailyCheckIndex.Source = "SELECT * From DailyCheckIndex WHERE CheckNumber = """ & check & """"
rsDailyCheckIndex.Open
rsDailyCheckIndex.Fields("Total").value = rsDailyCheckIndex.Fields("Subtotal").value + rsDailyCheckIndex.Fields("Tax").value + rsDailyCheckIndex.Fields("ServiceCharge").value
rsDailyCheckIndex.Update
rsDailyCheckIndex.Close
dbDailyCheckIndex.Close
End Sub


'Public Function GetTable(Check As String) As String
'GetTable = pGetTable(Check)
'End Function
'
'Private Function pGetTable(Check As String) As String
'pGetTable = ValueMatch(Wrap(GetNewMatchObj(, Check)), "Table")
'End Function

Public Sub UpdateDailyCheckIndex(check As String, Params As Dictionary)
pUpdateDailyCheckIndex check, Params
End Sub
Private Sub pUpdateDailyCheckIndex(check As String, Params As Dictionary)
OpenDbs
rsDailyCheckIndex.Source = "SELECT * From DailyCheckIndex WHERE CheckNumber = """ & check & """"
rsDailyCheckIndex.Open
Dim member As Variant
For Each member In Params
    If Not member = "Taxable" Then
    rsDailyCheckIndex.Fields(member).value = Params(member)
    End If
Next member
rsDailyCheckIndex.Fields("CheckNumber").value = check

rsDailyCheckIndex.Update
rsDailyCheckIndex.Close
dbDailyCheckIndex.Close
End Sub

Public Function FormatRecordset(ChecksRecordset As ADODB.RecordSet) As Variant

FormatRecordset = pFormatRecordset(ChecksRecordset)
End Function
Private Function pFormatRecordset(ChecksRecordset As ADODB.RecordSet) As Variant
'OpenDbs
'rsDailyCheckIndex.Source = strQuery
'ChecksRecordset.Open
Dim arrCheck() As Variant
Dim arrTotal() As Variant
Dim arrName() As Variant
Dim arrTime() As Variant
Dim arrServer() As Variant
Dim i As Integer
Dim TempArray() As Variant

If ChecksRecordset.RecordCount = 0 Then
    
    ReDim TempArray(0 To 0, 0 To 4)
    TempArray(0, 0) = ""
    TempArray(0, 1) = "No"
    TempArray(0, 2) = "Open"
    TempArray(0, 3) = "Checks"
    TempArray(0, 4) = ""
    pFormatRecordset = TempArray()
    CloseDbs
    Exit Function
End If

arrCheck() = ChecksRecordset.GetRows(, , "CheckNumber")
rsDailyCheckIndex.MoveFirst
arrName() = ChecksRecordset.GetRows(, , "OrderName")
rsDailyCheckIndex.MoveFirst
arrTime() = ChecksRecordset.GetRows(, , "PickupTime")
rsDailyCheckIndex.MoveFirst
arrTotal() = ChecksRecordset.GetRows(, , "SubTotal")
rsDailyCheckIndex.MoveFirst
arrServer = ChecksRecordset.GetRows(, , "ServerName")
CloseDbs

ReDim TempArray(0 To UBound(arrCheck, 2), 0 To 4)
For i = 0 To UBound(arrCheck, 2)
    TempArray(i, 0) = arrCheck(0, i)
Next i
For i = 0 To UBound(arrCheck, 2)
    TempArray(i, 1) = arrName(0, i)
Next i
For i = 0 To UBound(arrCheck, 2)
    TempArray(i, 2) = arrTime(0, i)
Next i
For i = 0 To UBound(arrCheck, 2)
    TempArray(i, 3) = arrTotal(0, i)
Next i
For i = 0 To UBound(arrCheck, 2)
    TempArray(i, 4) = arrServer(0, i)
Next i
pFormatRecordset = TempArray()
End Function

Public Sub CloseOrder(check As String)
pCloseOrder check
End Sub
Private Sub pCloseOrder(check As String)
Dim iIndex As New zclsDailyCheckIndex

Update iIndex.Wrap(GetNewUpdateObj(, check, "Closed", True))
Update iIndex.Wrap(GetNewUpdateObj(, check, "CheckClose", Format(Now, "hh:mm")))

End Sub






