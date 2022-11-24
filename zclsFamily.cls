VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public dbFamily As New ADODB.Connection
Public rsFamily As New ADODB.RecordSet
Private pStyleDict As Dictionary
Private pCount As Integer
Private pFamily As String
Private pBtnLocation As Worksheet
Private pMembers() As Variant
Private pMenuStyle As String
Private pFamilyGroup As String

Public Property Let StyleDict(value As Dictionary)
Set pStyleDict = value
End Property

Public Property Get StyleDict() As Dictionary
Set StyleDict = GetStyleDict(Me)
End Property


Public Property Let Count(value As Integer)
pCount = value
End Property
Public Property Get Count() As Integer
Count = GetCount(Me)
End Property

Public Property Let Family(value As String)
pFamily = value
End Property
Public Property Get Family() As String
Family = pFamily
End Property

Public Property Let FamilyGroup(value As String)
pFamilyGroup = value
End Property
Public Property Get FamilyGroup() As String
FamilyGroup = GetFamilyGroup
End Property

Public Property Let Members(value As Variant)
pMembers = value
End Property
Public Property Get Members() As Variant
Members = GetMembers
End Property

Public Property Let BtnLocation(value As Worksheet)
Set pBtnLocation = value
End Property
Public Property Get BtnLocation() As Worksheet
Set BtnLocation = pBtnLocation
End Property

Public Property Let MenuStyle(value As String)
pMenuStyle = value
End Property
Public Property Get MenuStyle() As String
MenuStyle = GetMenuStyle(Me)
End Property

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
GetDb = "Family"
End Function

Public Function GetDbFile() As String
GetDbFile = "Menu"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbFamily
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsFamily
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

'Public Function GetFamilyButtonObj(Family As String) As zclsFamily
'
'Set GetFamilyButtonObj = BtnParams(Family)
'End Function

Private Function BtnParams(Family As String) As zclsFamily

Dim x As New zclsFamily
x.Family = Family
Dim iMenu As New zclsMenu
Dim iMenuStyle As New aclsMenuStyle
Set iMenuStyle = iMenuStyle.GetNewMenuStyleObj(x.MenuStyle)
x.StyleDict = iMenuStyle.StyleDict
x.BtnLocation = GetBtnLocation(x.StyleDict("BtnLocation"))
x.Count = CountMatch(iMenu.Wrap(GetNewMatchObj("Family", Family, "NOT ItemName", "")))
Set BtnParams = x
Set iMenu = Nothing
Set x = Nothing
Set iMenuStyle = Nothing
End Function


Private Function GetBtnLocation(Location As String) As Worksheet
Dim wksht As Worksheet
For Each wksht In ActiveWorkbook.Worksheets
    If wksht.CodeName = Location Then
        Set GetBtnLocation = wksht
        Exit Function
    End If
Next wksht
Set wksht = Nothing
End Function

Private Function GetMembers() As Variant
Dim qry As String
Dim iMenu As New zclsMenu
Dim iDataObj As aclsDataObject
If ValueMatch(Wrap(GetNewMatchObj("Family", Me.Family)), "MultiMenu") = True Then
    Set iDataObj = Wrap(GetNewMatchObj)
    qry = "SELECT Family FROM Family WHERE MultiMenuParent = """ & Me.Family & """"
    GetMembers = RsToArray(GetRecordsetMatch(iDataObj, qry))
    iDataObj.CloseDbs iDataObj
    Set iMenu = Nothing
    Set iDataObj = Nothing
    Exit Function
End If
Set iDataObj = iMenu.Wrap(GetNewMatchObj)
qry = "SELECT ID FROM AllItems WHERE Family = """ & Me.Family & """ AND NOT ItemName = """" ORDER BY LEN(ID), ID ASC"
GetMembers = RsToArray(GetRecordsetMatch(iDataObj, qry))
iDataObj.CloseDbs iDataObj
Set iMenu = Nothing
End Function

Public Function GetMembersColl() As Collection
Dim qry As String
Dim iMenu As New zclsMenu
qry = "SELECT ID, ItemName FROM AllItems WHERE Family = """ & Me.Family & """ AND NOT ItemName = """" ORDER BY LEN(ID), ID ASC"
Dim rs As ADODB.RecordSet
Set rs = GetRecordsetMatch(iMenu.Wrap(GetNewMatchObj), qry)
Dim coll As New Collection
Dim dict As Dictionary
Dim i As Integer
Dim cyclecount As Integer
cyclecount = 1
i = 1
Dim BtnsPerPage As Integer
BtnsPerPage = Me.StyleDict("BtnsPerPage")
Set dict = New Dictionary
For i = 1 To rs.RecordCount
    If i = (cyclecount * BtnsPerPage) + 1 Then
        coll.Add dict
        Set dict = New Dictionary
        cyclecount = cyclecount + 1
    End If
    dict.Add rs.Fields("ID").value, rs.Fields("ItemName").value
    rs.MoveNext
Next i
coll.Add dict
Set GetMembersColl = coll
iMenu.CloseDbs
Set iMenu = Nothing
Set coll = Nothing
Set dict = Nothing
Set rs = Nothing
Set iMenu = Nothing
End Function

Private Function GetFamilyGroup() As String
GetFamilyGroup = ValueMatch(Wrap(GetNewMatchObj("Family", Me.Family)), "FamilyGroup")
End Function

Private Function GetMenuStyle(iFamily As zclsFamily) As String
GetMenuStyle = ValueMatch(Wrap(GetNewMatchObj("Family", iFamily.Family)), "MenuStyle")
End Function

Public Function GetCount(iFamily As zclsFamily) As Integer
If iFamily.Family = "" Then
    GetCount = 0
End If
Dim iMenu As New zclsMenu
GetCount = CountMatch(iMenu.Wrap(GetNewMatchObj("Family", iFamily.Family, "NOT ItemName", "")))
Set iMenu = Nothing
End Function

Public Function GetStyleDict(iFamily As zclsFamily) As Dictionary
Dim iMenuStyle As New aclsMenuStyle
Set iMenuStyle = iMenuStyle.GetNewMenuStyleObj(iFamily.MenuStyle)
Set GetStyleDict = iMenuStyle.StyleDict
Set iMenuStyle = Nothing
End Function

Public Sub AddNew(FamilyDict As Dictionary)
Dim NewFamily As String
NewFamily = FamilyDict("Family")
Dim iDataObj As New aclsDataObject
Set iDataObj = Wrap(GetNewMatchObj("Family", NewFamily))
iDataObj.OpenDbs iDataObj
iDataObj.rs.Source = iDataObj.Db
iDataObj.rs.Open
iDataObj.rs.AddNew Array("Family"), Array(NewFamily)
iDataObj.rs.Update
iDataObj.CloseDbs iDataObj
UpdateFromDict iDataObj, FamilyDict
End Sub

Public Sub Remove(Family As String)
'Dim iBtn As New aclsItemButton
Me.Family = Family
'iBtn.RemoveArray Family, Me.Members
DeleteMatch Wrap(GetNewMatchObj("Family", Family))
Dim iMenu As New zclsMenu
Update iMenu.Wrap(GetNewUpdateObj("Family", Family, "Family", "Unassigned"))
Set iMenu = Nothing
'Set iBtn = Nothing
End Sub

Public Sub RemoveFromGUI(iFamily As zclsFamily)
Update Wrap(GetNewUpdateObj("Family", iFamily.Family, "Active", False))
'iFamily.BtnLocation.Shapes(Family & "Blank").Delete
End Sub

Public Sub Activate()
Update Wrap(GetNewUpdateObj("Family", Me.Family, "Active", True))
End Sub

Public Sub Deactivate()
Update Wrap(GetNewUpdateObj("Family", Me.Family, "Active", False))
End Sub



