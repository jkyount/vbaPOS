VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Public dbMenu As New ADODB.Connection
Public rsMenu As New ADODB.RecordSet

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
GetDb = "AllItems"
End Function

Public Function GetDbFile() As String
GetDbFile = "Menu"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbMenu
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsMenu
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

Public Function GetItemClassCode(ItemID As Integer) As Integer
GetItemClassCode = pGetItemClassCode(ItemID)
End Function

Private Function pGetItemClassCode(ItemID As Integer) As Integer
OpenDbs
rsMenu.Source = "SELECT * FROM AllItems WHERE ID = " & ItemID & ""
rsMenu.Open
pGetItemClassCode = rsMenu.Fields("ClassCode").value
rsMenu.Close
dbMenu.Close
End Function

Public Function GetNextItemID(Family As String) As Integer
'STILL STRING
GetNextItemID = pGetNextItemID(Family)
End Function

Private Function pGetNextItemID(Optional Family As String) As Integer
'STILL STRING
Dim FamilyClause As String
FamilyClause = ""
If Not Family = "" Then
    FamilyClause = "Family = """ & Family & """ AND "
End If
OpenDbs
rsMenu.Source = "SELECT * FROM AllItems WHERE " & FamilyClause & "ItemName = """" ORDER BY ID ASC"
rsMenu.Open
If rsMenu.EOF Then
    rsMenu.AddNew
    rsMenu.Fields("ItemID").value = "Item" & rsMenu.Fields("ID").value
    rsMenu.Update
    pGetNextItemID = rsMenu.Fields("ID").value
    rsMenu.Close
    dbMenu.Close
    Exit Function
End If
rsMenu.MoveFirst
pGetNextItemID = rsMenu.Fields("ID").value
rsMenu.Close
dbMenu.Close
End Function

'Private Function pGetNextItemID(Family As String) As Integer
''STILL STRING
'OpenDbs
'rsMenu.Source = "SELECT * FROM AllItems WHERE Family = """ & Family & """ AND ItemName = """""
'rsMenu.Open
'rsMenu.MoveFirst
'Dim pItemID As String
'pItemID = rsMenu.Fields("ID").value
'pGetNextItemID = pItemID
'rsMenu.Close
'dbMenu.Close
'End Function

Public Sub CreateNewSpecialItem(ItemID As Integer, ItemName As String, Price As Currency)
pCreateNewSpecialItem ItemID, ItemName, Price
End Sub
Private Sub pCreateNewSpecialItem(ItemID As Integer, ItemName As String, Price As Currency)
OpenDbs
rsMenu.Source = "SELECT * FROM AllItems WHERE ID = " & ItemID & ""
rsMenu.Open
rsMenu.Fields("ItemName").value = ItemName
rsMenu.Fields("Price").value = Price
rsMenu.Fields("Category").value = "Food"
rsMenu.Update
rsMenu.Close
dbMenu.Close
End Sub

Public Sub ClearCustomItems()
ClearNameAndPrice "CustomItem"
ClearNameAndPrice "SpclInstruction"
End Sub

Private Sub ClearNameAndPrice(Family As String)
Dim iMenu As New zclsMenu
Dim iDataObject As New aclsDataObject
Set iDataObject = Wrap(GetNewMatchObj("Family", Family))
Set rsMenu = GetRecordsetMatch(iDataObject, ConstructMatchQuery(iDataObject))
If Not rsMenu.EOF Then
    Do Until rsMenu.EOF
    rsMenu.Fields("ItemName").value = ""
    rsMenu.Fields("Price").value = 0
    rsMenu.MoveNext
    Loop
    rsMenu.UpdateBatch
End If
CloseDbs
End Sub

Public Function GetItemsInFamily(Family As String) As Variant
GetItemsInFamily = pGetItemsInFamily(Family)
End Function

Private Function pGetItemsInFamily(Family As String) As Variant
Dim iDataObject As New aclsDataObject
Set iDataObject = Wrap(iDataObject.GetNewMatchObj("Family", Family, "NOT ItemName", ""))
Dim arr As Variant
arr = GetMatch(iDataObject)
If IsEmpty(arr) = True Then
    ReDim arr(1 To 1)
    Dim temparr(0 To 0, 0 To 0) As Variant
    temparr(0, 0) = ""
    arr(1) = temparr
    pGetItemsInFamily = arr
    Set iDataObject = Nothing
    Exit Function
End If
pGetItemsInFamily = arr
Set iDataObject = Nothing
End Function

Public Function PairItemNameAndID(arr As Variant) As Variant
If UBound(arr(1)) < 2 Then
    PairItemNameAndID = arr
    Exit Function
End If
Dim i As Integer
Dim TempArray As Variant
TempArray = arr
For i = 1 To UBound(TempArray)
    TempArray(i)(0, 0) = TempArray(i)(0, 0) & "  -  " & TempArray(i)(2, 0)
Next i
PairItemNameAndID = TempArray
End Function

Public Sub Remove(ItemID As Integer)
pRemove ItemID
End Sub

Private Sub pRemove(ItemID As Integer)
'Update Wrap(GetNewUpdateObj("ID", ItemID, "ItemName", ""))
'Update Wrap(GetNewUpdateObj("ID", ItemID, "Price", "0"))
'Update Wrap(GetNewUpdateObj("ID", ItemID, "Req1", ""))
'Update Wrap(GetNewUpdateObj("ID", ItemID, "Req2", ""))
'Update Wrap(GetNewUpdateObj("ID", ItemID, "Family", ""))
'Update Wrap(GetNewUpdateObj("ID", ItemID, "ClassCode", "0"))
Dim iMenu As New zclsMenu
Dim rs As ADODB.RecordSet
Set rs = GetRecordsetMatch(Wrap(GetNewMatchObj("ID", ItemID)))
Dim fld As Variant
For Each fld In rs.Fields
    If Not fld.name = "ID" Then
        If Not fld.name = "ItemID" Then
            fld.value = False
        End If
    End If
    If fld.Type = adVarWChar Then
        If Not fld.name = "ItemID" Then
            fld.value = ""
        End If
    End If
Next fld
rs.Update
CloseDbs
Set iMenu = Nothing
End Sub

'Public Sub AddNewFamily(FamilyDict As Dictionary)
'Dim rs As ADODB.RecordSet
'Set rs = Me.GetRs
'OpenDbs
'rs.Source = "AllItems"
'rs.Open
'Dim i As Integer
'For i = 1 To 24
'    rs.AddNew
'    rs.Fields("ItemID").value = "Item" & CStr(rs.Fields("ID").value)
'    rs.Fields("Family").value = FamilyDict("Family")
'    rs.Fields("ClassCode").value = FamilyDict("ClassCode")
'
'Next i
'rs.UpdateBatch
'CloseDbs
'Set rs = Nothing
'End Sub
