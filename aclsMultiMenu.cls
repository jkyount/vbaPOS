VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsMultiMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public dbMultiMenu As New ADODB.Connection
Public rsMultiMenu As New ADODB.RecordSet

Private pMembers As Collection


Public Property Get Members(MultiMenuName As String) As Collection
Set Members = GetMembers(MultiMenuName)
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
Set GetConn = dbMultiMenu
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsMultiMenu
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

Private Function GetMembers(MultiMenuName As String) As Collection
Dim iDataObj As aclsDataObject
Set iDataObj = Wrap(GetNewMatchObj)
Dim qry As String
qry = "SELECT Family FROM Family WHERE MultiMenuParent = """ & MultiMenuName & """ ORDER BY ID ASC"
Dim rs As New ADODB.RecordSet
Set rs = GetRecordsetMatch(iDataObj, qry)
Dim coll As New Collection
Do Until rs.EOF
    coll.Add rs.Fields("Family").value, rs.Fields("Family").value
    rs.MoveNext
Loop
iDataObj.CloseDbs iDataObj
Set GetMembers = coll
Set iDataObj = Nothing
Set rs = Nothing
Set coll = Nothing
End Function

Public Function GetAvailableFamilies(MultiMenuName As String) As Variant
GetAvailableFamilies = FilteredMatch(Wrap(GetNewMatchObj("MultiMenu = FALSE AND Active", True, "NOT MultiMenuParent", MultiMenuName)), "Family")
End Function

Public Function GetCurrentFamilies(MultiMenuName As String) As Variant
GetCurrentFamilies = FilteredMatch(Wrap(GetNewMatchObj("MultiMenuParent", MultiMenuName)), "Family")
End Function

Public Sub AddMember(MultiMenuName As String, NewMemberName As String)
'5/17/22
'START HERE
Dim iFamily As New zclsFamily
iFamily.Family = MultiMenuName
If GetMembers(MultiMenuName).Count = 0 Then
    iFamily.Activate
End If
AssignMember MultiMenuName, NewMemberName
'PositionNewMemberButtons NewMemberName
End Sub

Public Sub RemoveMember(MultiMenuName As String, FormerMemberName As String)
'5/17/22
Dim iFamily As New zclsFamily
iFamily.Family = MultiMenuName
UnassignMember FormerMemberName
If GetMembers(MultiMenuName).Count = 0 Then
    iFamily.Deactivate
End If
End Sub

Private Sub AssignMember(MultiMenuName As String, NewMemberName As String)
'5/17/22
Update Wrap(GetNewUpdateObj("Family", NewMemberName, "MultiMenuParent", MultiMenuName))
Update Wrap(GetNewUpdateObj("Family", NewMemberName, "MultiMenuMember", True))
Update Wrap(GetNewUpdateObj("Family", NewMemberName, "MenuStyle", "Expanded"))
Update Wrap(GetNewUpdateObj("Family", NewMemberName, "Active", "False"))
End Sub

Private Sub UnassignMember(FormerMemberName As String)
'5/17/22
Update Wrap(GetNewUpdateObj("Family", FormerMemberName, "MultiMenuParent", ""))
Update Wrap(GetNewUpdateObj("Family", FormerMemberName, "MultiMenuMember", False))
'6/10/22
'Update Wrap(GetNewUpdateObj("Family", FormerMemberName, "Active", True))
End Sub


'Public Sub PositionNewMemberButtons(NewMemberName As String)
''5/17/22
'Dim NewFamily As New zclsFamily
'Set NewFamily = NewFamily.GetFamilyButtonObj(NewMemberName)
'Dim iBtn As New aclsItemButton
''iBtn.PositionAll NewFamily
'Set NewFamily = Nothing
'Set iBtn = Nothing
'End Sub

Public Sub Delete(MultiMenuName As String)
DeleteMatch Me.Wrap(GetNewMatchObj("Family", MultiMenuName))
End Sub


