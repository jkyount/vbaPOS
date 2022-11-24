VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsItemClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public dbItemClass As New ADODB.Connection
Public rsItemClass As New ADODB.RecordSet
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
GetDb = "ItemClass"
End Function

Public Function GetDbFile() As String
GetDbFile = "ItemClass"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbItemClass
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsItemClass
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
'Public Sub OpenDbs()
'dbItemClass.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Jared\POS\Access_Int\ItemClass.accdb;Persist Security Info=False;"
'dbItemClass.Open
'rsItemClass.CursorType = adUseClient
'rsItemClass.LockType = adLockOptimistic
'rsItemClass.ActiveConnection = dbItemClass
'End Sub

Public Function GetItemType(ClassCode As String) As Variant
Dim coll As New Collection
Set coll = GetClassColl
Set GetItemType = coll(ClassCode)
If ClassCode = 1 Then
    If CItem.Count > 0 Then
        Set GetItemType = coll(2)
    End If
End If
End Function

Public Function GetClassColl() As Collection
Dim ClassColl As New Collection
Dim x As Variant
Set x = New zclsPrimary
ClassColl.Add x, GetClassCode(TypeName(x))
Set x = New zclsNotPrimary
ClassColl.Add x, GetClassCode(TypeName(x))
Set x = New zclsCustomItem
ClassColl.Add x, GetClassCode(TypeName(x))
Set x = New zclsSpecialInstruction
ClassColl.Add x, GetClassCode(TypeName(x))
Set x = New zclsDiscount
ClassColl.Add x, GetClassCode(TypeName(x))
Set x = New zclsPizza
ClassColl.Add x, GetClassCode(TypeName(x))
Set x = New zclsPizzaTopping
ClassColl.Add x, GetClassCode(TypeName(x))
Set x = New zclsMod
ClassColl.Add x, GetClassCode(TypeName(x))
Set x = Nothing
Set GetClassColl = ClassColl
End Function

Public Function GetClassCode(ClassType As String) As String
OpenDbs
rsItemClass.Source = "SELECT * FROM ItemClass WHERE ClassType = """ & ClassType & """"
rsItemClass.Open
GetClassCode = CStr(rsItemClass.Fields("ClassCode").value)
rsItemClass.Close
dbItemClass.Close
End Function

Public Function GetAssignableClassNames() As Variant
Dim ClassNameArray As Variant
Dim ClassCodeArray As Variant

ClassNameArray = FilteredMatch(Wrap(GetNewMatchObj("Assignable", True)), "ClassName")
ClassCodeArray = FilteredMatch(Wrap(GetNewMatchObj("Assignable", True)), "ClassCode")

Dim i As Integer
For i = 1 To UBound(ClassNameArray)
    ClassCodeArray(i)(0, 0) = ClassCodeArray(i)(0, 0) & " - " & ClassNameArray(i)(0, 0)
Next i
GetAssignableClassNames = ClassCodeArray
End Function
