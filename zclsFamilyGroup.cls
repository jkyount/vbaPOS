VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsFamilyGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public dbFamilyGroup As New ADODB.Connection
Public rsFamilyGroup As New ADODB.RecordSet

'Private pBtnTopRef As Double
'Private pBtnLeftRef As Double
'Private pBtnRowOffset As Double
'Private pBtnColumnOffset As Double
'Private pBtnsPerColumn As Integer

Private pCount As Integer
Private pFamilyGroup As String

'Private pBtnLocation As Worksheet
Private pMembers() As Variant


'Public Property Let BtnsPerColumn(value As Integer)
'pBtnsPerColumn = value
'End Property
'Public Property Get BtnsPerColumn() As Integer
'BtnsPerColumn = pBtnsPerColumn
'End Property
'Public Property Let BtnColumnOffset(value As Double)
'pBtnColumnOffset = value
'End Property
'Public Property Get BtnColumnOffset() As Double
'BtnColumnOffset = pBtnColumnOffset
'End Property
'Public Property Let BtnRowOffset(value As Double)
'pBtnRowOffset = value
'End Property
'Public Property Get BtnRowOffset() As Double
'BtnRowOffset = pBtnRowOffset
'End Property
'Public Property Let BtnLeftRef(value As Double)
'pBtnLeftRef = value
'End Property
'Public Property Get BtnLeftRef() As Double
'BtnLeftRef = pBtnLeftRef
'End Property
'Public Property Let BtnTopRef(value As Double)
'pBtnTopRef = value
'End Property
'Public Property Get BtnTopRef() As Double
'BtnTopRef = pBtnTopRef
'End Property
Public Property Let Count(value As Integer)
pCount = value
End Property
Public Property Get Count() As Integer
Count = pCount
End Property

Public Property Let FamilyGroup(value As String)
pFamilyGroup = value
End Property
Public Property Get FamilyGroup() As String
FamilyGroup = pFamilyGroup
End Property


Public Property Let Members(value As Variant)
pMembers = value
End Property
Public Property Get Members() As Variant
Members = GetMembers
End Property
'
'Public Property Let BtnLocation(value As Worksheet)
'Set pBtnLocation = value
'End Property
'Public Property Get BtnLocation() As Worksheet
'Set BtnLocation = pBtnLocation
'End Property

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
GetDb = "FamilyGroup"
End Function

Public Function GetDbFile() As String
GetDbFile = "Menu"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbFamilyGroup
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsFamilyGroup
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

'Public Function GetFamilyGroupButtonObj(FamilyGroup As String) As zclsFamilyGroup
'Dim iFamilyGroup As New zclsFamilyGroup
'Set iFamilyGroup = BtnParams(FamilyGroup)
'Set GetFamilyGroupButtonObj = iFamilyGroup
'Set iFamilyGroup = Nothing
'End Function

'Private Function BtnParams(FamilyGroup As String) As zclsFamilyGroup
'Dim x As New zclsFamilyGroup
'Dim dict As Dictionary
'Set dict = GetValueDict(Wrap(GetNewMatchObj("FamilyGroup", FamilyGroup)))(1)
'x.BtnsPerColumn = dict("BtnsPerColumn")
'x.BtnTopRef = dict("BtnTopRef")
'x.BtnLeftRef = dict("BtnLeftRef")
'x.BtnRowOffset = dict("BtnRowOffset")
'x.BtnColumnOffset = dict("BtnColumnOffset")
'x.FamilyGroup = FamilyGroup
'x.BtnLocation = GetBtnLocation(dict("BtnLocation"))
'Dim iFamily As New zclsFamily
'x.Count = CountMatch(iFamily.Wrap(GetNewMatchObj("FamilyGroup", FamilyGroup)))
'Set BtnParams = x
'Set iFamily = Nothing
'Set dict = Nothing
'Set x = Nothing
'End Function

'Private Function GetBtnLocation(Location As String) As Worksheet
'Dim wksht As Worksheet
'For Each wksht In ActiveWorkbook.Worksheets
'    If wksht.CodeName = Location Then
'        Set GetBtnLocation = wksht
'        Exit Function
'    End If
'Next wksht
'Set wksht = Nothing
'End Function

Public Function GetMembers(Optional FamilyGroup As String) As Variant
If FamilyGroup = "" Then FamilyGroup = Me.FamilyGroup
Dim iFamily As New zclsFamily
Dim iDataObj As New aclsDataObject
Set iDataObj = iFamily.Wrap(GetNewMatchObj("FamilyGroup", FamilyGroup, "NOT MultiMenu", True))

GetMembers = FilteredMatch(iDataObj, "Family")
Set iFamily = Nothing

End Function

Public Function GetFamilyGroups() As Variant

Dim iDataObj As New aclsDataObject
Set iDataObj = Me.Wrap(GetNewMatchObj)
Dim qry As String
qry = "SELECT FamilyGroup FROM FamilyGroup"
GetFamilyGroups = RsToArray(GetRecordsetMatch(iDataObj, qry))
iDataObj.CloseDbs iDataObj


End Function

