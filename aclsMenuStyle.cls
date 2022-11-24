VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsMenuStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public dbMenuStyle As New ADODB.Connection
Public rsMenuStyle As New ADODB.RecordSet

Private pStyleDict As Dictionary

Public Property Let StyleDict(value As Dictionary)
Set pStyleDict = value
End Property

Public Property Get StyleDict() As Dictionary
Set StyleDict = pStyleDict
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
GetDb = "MenuStyle"
End Function

Public Function GetDbFile() As String
GetDbFile = "Menu"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbMenuStyle
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsMenuStyle
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


Public Function GetNewMenuStyleObj(MenuStyle As String) As aclsMenuStyle
Set GetNewMenuStyleObj = MenuStyleParams(MenuStyle)
End Function

Private Function MenuStyleParams(MenuStyle As String) As aclsMenuStyle
Dim x As New aclsMenuStyle
Dim iDataObj As aclsDataObject
Set iDataObj = Wrap(GetNewMatchObj("MenuStyle", MenuStyle))

x.StyleDict = CDictCollection(GetRecordsetMatch(iDataObj, ConstructMatchQuery(iDataObj)))(1)
Set MenuStyleParams = x


End Function

Public Function GetMenuStyles() As Variant

Dim iDataObj As New aclsDataObject
Set iDataObj = Me.Wrap(GetNewMatchObj)
Dim qry As String
qry = "SELECT MenuStyle FROM MenuStyle"
GetMenuStyles = RsToArray(GetRecordsetMatch(iDataObj, qry))
iDataObj.CloseDbs iDataObj

Set iDataObj = Nothing
End Function
