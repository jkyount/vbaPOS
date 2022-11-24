VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsDailyCheckDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public dbDailyCheckDetail As New ADODB.Connection
Public rsDailyCheckDetail As New ADODB.RecordSet

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
GetDb = "DailyCheckDetail"
End Function

Public Function GetDbFile() As String
GetDbFile = "CheckDb"
End Function

Public Function GetArchive() As String
GetArchive = "ArchivedCheckDetail"
End Function
Public Function GetArchiveDbFile() As String
GetArchiveDbFile = "ReportsDb"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbDailyCheckDetail
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsDailyCheckDetail
End Function


'==========================================================================







Public Function GetValueDict() As Dictionary
OpenDbs
rsDailyCheckDetail.Source = "DailyCheckDetail"
rsDailyCheckDetail.Open
Dim fld As Object
Dim dict As New Dictionary
For Each fld In rsDailyCheckDetail.Fields
    dict.Add fld.name, ""
Next fld
Set GetValueDict = dict
CloseDbs
Set fld = Nothing
Set dict = Nothing
End Function
'==========================================================================






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


Public Function GetNextLocalGroup(check As String) As Integer
OpenDbs
rsDailyCheckDetail.Source = "SELECT LocalGroup FROM DailyCheckDetail WHERE CheckNumber = """ & check & """ ORDER BY LocalGroup ASC"
rsDailyCheckDetail.Open
If rsDailyCheckDetail.RecordCount = 0 Then
    GetNextLocalGroup = 1
    rsDailyCheckDetail.Close
    dbDailyCheckDetail.Close
    Exit Function
End If
rsDailyCheckDetail.MoveLast
GetNextLocalGroup = rsDailyCheckDetail.Fields("LocalGroup").value + 1
CloseDbs
End Function

Public Function GetNextEntityGroup(check As String) As Integer
OpenDbs
rsDailyCheckDetail.Source = "SELECT EntityGroup FROM DailyCheckDetail WHERE CheckNumber = """ & check & """ ORDER BY EntityGroup ASC"
rsDailyCheckDetail.Open
If rsDailyCheckDetail.RecordCount = 0 Then
    GetNextEntityGroup = 1
    rsDailyCheckDetail.Close
    dbDailyCheckDetail.Close
    Exit Function
End If
rsDailyCheckDetail.MoveLast
GetNextEntityGroup = rsDailyCheckDetail.Fields("EntityGroup").value + 1
CloseDbs
End Function

'==========================================================================


Public Sub AddCurrentItems(coll As Collection)
Dim member As aclsItem
Dim LocalGroup As Integer, EntityGroup As Integer

EntityGroup = GetNextEntityGroup(currentcheck)
For Each member In coll

    'If Not TypeName(member) = "zclsEmpty" Then
    LocalGroup = GetNextLocalGroup(currentcheck)
        'NormalizeParentIDs member, LocalGroup
    AddItem member, LocalGroup, EntityGroup
        'AdvCell
    'End If
Next member
End Sub

'Public Sub NormalizeParentIDs(member As aclsItem, LocalGroup As Integer)
'Dim child As bclsChild
'For Each child In member.Children.coll
'    GetItemByID(child.ID).Parent.ID = LocalGroup
'Next child
'Set child = Nothing
'End Sub

Public Sub AddItem(item As aclsItem, LocalGroup As Integer, EntityGroup As Integer)
'==========================================================================


'==========================================================================

OpenDbs
rsDailyCheckDetail.Source = "SELECT * From DailyCheckDetail WHERE CheckNumber = """ & currentcheck & """ ORDER BY Seat ASC"
rsDailyCheckDetail.Open
rsDailyCheckDetail.AddNew
rsDailyCheckDetail.Fields("CheckNumber").value = currentcheck
rsDailyCheckDetail.Fields("ItemID").value = item.ItemID
rsDailyCheckDetail.Fields("ItemIndicator").value = item.ItemIndicator
rsDailyCheckDetail.Fields("ItemName").value = item.ItemName
rsDailyCheckDetail.Fields("Price").value = item.Price
rsDailyCheckDetail.Fields("Family").value = item.Family
rsDailyCheckDetail.Fields("Category").value = item.Category
rsDailyCheckDetail.Fields("AlwaysTax").value = item.AlwaysTax
rsDailyCheckDetail.Fields("PrintKitchen").value = item.PrintKitchen
rsDailyCheckDetail.Fields("PrintPantry").value = item.PrintPantry
rsDailyCheckDetail.Fields("Seat").value = GetCurrentSeat
rsDailyCheckDetail.Fields("LocalGroup").value = LocalGroup
rsDailyCheckDetail.Fields("IsPrimaryItem").value = item.IsPrimaryItem
rsDailyCheckDetail.Fields("EntityGroup").value = EntityGroup
rsDailyCheckDetail.Update
CloseDbs

'==========================================================================

'==========================================================================

End Sub

'==========================================================================



'Public Function GetTax_DineIn(Check As String) As Currency
'GetTax_DineIn = pGetTax_DineIn(Check)
'End Function
'
'Private Function pGetTax_DineIn(Check As String) As Currency
'pGetTax_DineIn = ((SumMatch(Wrap(GetNewMatchObj(, Check)), "Price")) * 0.08)
'End Function

'Public Function GetTax_Carryout(Check As String) As Currency
'GetTax_Carryout = pGetTax_Carryout(Check)
'End Function
'
'Private Function pGetTax_Carryout(Check As String) As Currency
'pGetTax_Carryout = ((SumMatch(Wrap(GetNewMatchObj(, Check, "AlwaysTax", True)), "Price")) * 0.08)
'End Function

Public Function IsValidCheck(check As String) As Boolean
IsValidCheck = pIsValidCheck(check)
End Function
Private Function pIsValidCheck(check As String) As Boolean

If check = "" Then
    pIsValidCheck = False
    Exit Function
End If
OpenDbs
rsDailyCheckDetail.Source = "SELECT * From DailyCheckDetail WHERE CheckNumber = """ & check & """"
rsDailyCheckDetail.Open
If rsDailyCheckDetail.RecordCount = 0 Then
    CloseDbs
    Exit Function
End If
pIsValidCheck = True
CloseDbs
End Function

'==========================================================================


Public Sub SendCurrentItems(check As String)
pSendCurrentItems check
End Sub

Private Sub pSendCurrentItems(check As String)
OpenDbs
rsDailyCheckDetail.Source = "SELECT * From DailyCheckDetail WHERE CheckNumber = """ & currentcheck & """ ORDER BY Seat ASC"
rsDailyCheckDetail.Open
Do Until rsDailyCheckDetail.EOF
    rsDailyCheckDetail.Fields("Sent").value = True
    rsDailyCheckDetail.MoveNext
Loop
CloseDbs
End Sub


Public Sub RemoveItem(check As String, EntityGroup As Integer)
pRemoveItem check, EntityGroup
End Sub
Private Sub pRemoveItem(check As String, EntityGroup As Integer)
DeleteMatch Wrap(GetNewMatchObj(, check, "EntityGroup", EntityGroup))
End Sub


'==========================================================================


Public Sub CopyToTemp(check As String)
pCopyToTemp check
End Sub
Private Sub pCopyToTemp(check As String)
ClearTemp
AppendToTemp check
End Sub

Public Sub AppendToTemp(check As String)
pAppendToTemp check
End Sub
Public Sub pAppendToTemp(check As String)
Dim CopyToTemp As String
CopyToTemp = "INSERT INTO TempCheck IN 'C:\Jared\POS\Access_Int\CheckDB.accdb' SELECT * FROM DailyCheckDetail WHERE DailyCheckDetail.CheckNumber = """ & check & """"
OpenDbs
dbDailyCheckDetail.Execute CopyToTemp
dbDailyCheckDetail.Close
End Sub
Private Sub ClearTemp()
OpenDbs
rsDailyCheckDetail.Source = "TempCheck"
rsDailyCheckDetail.Open
Do Until rsDailyCheckDetail.EOF
    rsDailyCheckDetail.Delete
    rsDailyCheckDetail.MoveNext
Loop
CloseDbs
End Sub

Public Sub SubmitChanges(check As String)
DeleteMatch Wrap(GetNewMatchObj(, check))

Dim SubmitChanges As String
SubmitChanges = "INSERT INTO DailyCheckDetail IN 'C:\Jared\POS\Access_Int\CheckDB.accdb' SELECT * FROM TempCheck WHERE TempCheck.CheckNumber = """ & check & """"
OpenDbs
dbDailyCheckDetail.Execute SubmitChanges
dbDailyCheckDetail.Close
End Sub

'==========================================================================



