Attribute VB_Name = "Functions_Shared"
Option Explicit
Dim rs As ADODB.RecordSet


Public Function GetNewMatchObj(Optional Where As String = "CheckNumber", Optional Equals As Variant, Optional AndWhere As String = "", Optional Equals2 As Variant) As aclsDataObject
Dim iDataObject As New aclsDataObject

Set GetNewMatchObj = iDataObject.GetNewMatchObj(Where, Equals, AndWhere, Equals2)
Set iDataObject = Nothing
End Function

Public Function GetNewUpdateObj(Optional Where As String = "CheckNumber", Optional Equals As Variant, Optional UpdateField As String = "", Optional UpdateValue As Variant) As aclsDataObject
Dim iDataObject As New aclsDataObject

Set GetNewUpdateObj = iDataObject.GetNewUpdateObj(Where, Equals, UpdateField, UpdateValue)
Set iDataObject = Nothing
End Function

Public Function GetNewArchiveObj(obj As Variant, iDataObj As aclsDataObject) As aclsDataObject

Set GetNewArchiveObj = iDataObj.GetNewArchiveObj(obj, iDataObj)

End Function

Public Function Match(DataObject As aclsDataObject) As Boolean

Set rs = DataObject.GetMatch(DataObject, ConstructMatchQuery(DataObject))
If Not rs.EOF = True Then
    Match = True
End If
DataObject.CloseDbs DataObject
Set rs = Nothing
End Function

Public Function GetMatch(DataObject As aclsDataObject) As Variant
'5/21
If Match(DataObject) = False Then
    MsgBox "Invalid match parameters.  No match found."
    Exit Function
End If
Set rs = DataObject.GetMatch(DataObject, ConstructMatchQuery(DataObject))
GetMatch = GetResult(rs)
'Dim ResultArray() As Variant
'If rs.RecordCount = 0 Then
'    GetMatch = Empty
'    DataObject.CloseDbs DataObject
'    Exit Function
'End If
'ReDim ResultArray(1 To rs.RecordCount)
'Dim i As Integer
'For i = 1 To rs.RecordCount
'    ResultArray(i) = rs.GetRows(1)
'Next i
'GetMatch = ResultArray
DataObject.CloseDbs DataObject
Set rs = Nothing
End Function

Public Function GetRecordsetMatch(obj As aclsDataObject, Optional Query As String = "") As ADODB.RecordSet
If Query = "" Then
    Query = ConstructMatchQuery(obj)
End If
Dim iDataObject As New aclsDataObject
Set rs = iDataObject.GetMatch(obj, Query)
Set GetRecordsetMatch = rs
Set rs = Nothing
End Function

Public Function CDictCollection(rs As ADODB.RecordSet) As Collection
Dim coll As New Collection
Dim dict As Dictionary

If rs.RecordCount = 0 Then
    coll.Add dict
    Set CDictCollection = coll
    Exit Function
End If
 
Dim fld As Object
rs.MoveFirst
Do Until rs.EOF
    Set dict = New Dictionary
    For Each fld In rs.Fields
        dict.Add fld.name, fld.value
    Next fld
    coll.Add dict
    rs.MoveNext
Loop
Set CDictCollection = coll
            
End Function

Public Function CountMatch(obj As aclsDataObject) As Integer
Set rs = GetRecordsetMatch(obj, ConstructMatchQuery(obj))
CountMatch = rs.RecordCount
obj.CloseDbs obj
Set rs = Nothing
End Function

Public Function RsToArray(RecordSet As ADODB.RecordSet) As Variant
If Not RecordSet.RecordCount = 0 Then
RsToArray = RecordSet.GetRows
End If
End Function

Public Sub DeleteMatch(obj As aclsDataObject)
Set rs = GetRecordsetMatch(obj, ConstructMatchQuery(obj))
If Not rs.RecordCount = 0 Then
    Do Until rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
End If
obj.CloseDbs obj
Set rs = Nothing
End Sub

Public Function SumMatch(obj As aclsDataObject, Field As String) As Double
Set rs = GetRecordsetMatch(obj, ConstructMatchQuery(obj))
If Not rs.RecordCount = 0 Then
    Do Until rs.EOF
        SumMatch = SumMatch + rs.Fields(Field).value
        rs.MoveNext
    Loop
    obj.CloseDbs obj
    Set rs = Nothing
    Exit Function
End If
SumMatch = 0
obj.CloseDbs obj
Set rs = Nothing
End Function

Public Function ValueMatch(obj As aclsDataObject, Field) As Variant
Set rs = GetRecordsetMatch(obj, ConstructMatchQuery(obj))
If Not rs.RecordCount = 0 Then
    ValueMatch = rs.Fields(Field).value
End If
obj.CloseDbs obj
Set rs = Nothing
End Function

Public Sub AddToMatch(obj As aclsDataObject)
Dim i As Integer
Dim UpdateField As String
Dim Amount As String
Amount = obj.Value2
UpdateField = obj.Field2

Set rs = GetRecordsetMatch(obj, ConstructUpdateQuery(obj))
If Not rs.RecordCount = 0 Then
    For i = 1 To rs.RecordCount
    rs.Fields(UpdateField).value = rs.Fields(UpdateField).value + Amount
    rs.Update
    rs.MoveNext
    Next i
End If
obj.CloseDbs obj
Set rs = Nothing
End Sub

Public Function FilteredMatch(obj As aclsDataObject, ParamArray SelectedFields() As Variant) As Variant
'5/21
Dim arr As Variant
ReDim arr(0 To UBound(SelectedFields)) As Variant
arr = SelectedFields
Set rs = obj.GetMatch(obj, ConstructFilteredQuery(obj, arr))
FilteredMatch = GetResult(rs)
'If rs.RecordCount = 0 Then
'    FilteredMatch = False
'    obj.CloseDbs obj
'    Set rs = Nothing
'    Exit Function
'End If
'Dim ResultArray() As Variant
'ReDim ResultArray(1 To rs.RecordCount)
'Dim i As Integer
'For i = 1 To rs.RecordCount
'    ResultArray(i) = rs.GetRows(1)
'Next i
'FilteredMatch = ResultArray
obj.CloseDbs obj
Set rs = Nothing
End Function

Public Function FltrdOrdrdMtch(obj As aclsDataObject, OrderBy As String, ParamArray SelectedFields() As Variant) As Variant
'5/21
Dim arr As Variant
ReDim arr(0 To UBound(SelectedFields)) As Variant
arr = SelectedFields
Set rs = obj.GetMatch(obj, ConstructOrdrdFltrdQry(ConstructFilteredQuery(obj, arr), OrderBy))
FltrdOrdrdMtch = GetResult(rs)
obj.CloseDbs obj
Set rs = Nothing
End Function

Public Function GetResult(rs As ADODB.RecordSet) As Variant
If rs.RecordCount = 0 Then
    GetResult = False
    Exit Function
End If
Dim ResultArray() As Variant
ReDim ResultArray(1 To rs.RecordCount)
Dim i As Integer
For i = 1 To rs.RecordCount
    ResultArray(i) = rs.GetRows(1)
Next i
GetResult = ResultArray
End Function

Public Function ConstructMatchQuery(obj As aclsDataObject, Optional Filter As String = "*") As String
Dim str As String
If obj.Field1 = "" Then
    str = "SELECT " & Filter & " FROM " & obj.Db & " WHERE False"
    ConstructMatchQuery = str
    Exit Function
End If
If VarType(obj.Value1) = vbString Then
    str = "SELECT " & Filter & " FROM " & obj.Db & " WHERE " & obj.Field1 & " = """ & obj.Value1 & """"
End If

If Not VarType(obj.Value1) = vbString Then
    str = "SELECT " & Filter & " FROM " & obj.Db & " WHERE " & obj.Field1 & " = " & obj.Value1 & ""
End If
If Not obj.Field2 = "" Then
    If VarType(obj.Value2) = vbString Then
        str = str & " AND " & obj.Field2 & " = """ & obj.Value2 & """"
    End If
    If Not VarType(obj.Value2) = vbString Then
        str = str & " AND " & obj.Field2 & " = " & obj.Value2 & ""
    End If
End If
ConstructMatchQuery = str
End Function

Public Function ConstructFilteredQuery(obj As aclsDataObject, SelectedFields As Variant) As String
Dim str As String
str = Join(SelectedFields, ", ")
ConstructFilteredQuery = ConstructMatchQuery(obj, str)
End Function

Public Function ConstructOrderedQuery(obj As aclsDataObject, OrderBy As String) As String

ConstructOrderedQuery = ConstructMatchQuery(obj) & OrderBy
End Function

Public Function ConstructOrdrdFltrdQry(qry As String, OrderBy As String) As String

ConstructOrdrdFltrdQry = qry & " ORDER BY " & OrderBy
End Function
'==========================================================================

Public Sub Update(obj As aclsDataObject)
Dim i As Integer
Set rs = GetRecordsetMatch(obj, ConstructUpdateQuery(obj))
If Not rs.RecordCount = 0 Then
    For i = 1 To rs.RecordCount
    rs.Fields(obj.Field2).value = obj.Value2
    rs.Update
    rs.MoveNext
    Next i
End If
obj.CloseDbs obj
Set rs = Nothing
End Sub

Public Sub UpdateFromDict(obj As aclsDataObject, dict As Dictionary)
Set rs = GetRecordsetMatch(obj, ConstructMatchQuery(obj))
If Not rs.RecordCount = 0 Then
    Dim fld As Object
    For Each fld In rs.Fields
        If Not fld.name = "ItemID" Then
            If dict.Exists(fld.name) Then
            
                fld.value = dict(fld.name)
            End If
        End If
    Next fld
    
    
    rs.Update
 
End If
obj.CloseDbs obj
Set rs = Nothing
End Sub

Public Function ConstructUpdateQuery(obj As aclsDataObject) As String
Dim str As String
If VarType(obj.Value1) = vbString Then
    str = "SELECT * FROM " & obj.Db & " WHERE " & obj.Field1 & " = """ & obj.Value1 & """"
End If
If Not VarType(obj.Value1) = vbString Then
    str = "SELECT * FROM " & obj.Db & " WHERE " & obj.Field1 & " = " & obj.Value1 & ""
End If
ConstructUpdateQuery = str
End Function


'==========================================================================




Public Function GetValueDict(obj As aclsDataObject, Optional qry As String = "") As Collection
Dim coll As New Collection
Dim dict As Dictionary
If Match(obj) = False Then
    
    Set dict = New Dictionary
    coll.Add dict
    Set GetValueDict = coll
    Set dict = Nothing
    Set coll = Nothing
    Exit Function
End If
If qry = "" Then
    qry = ConstructMatchQuery(obj)
End If
Set rs = GetRecordsetMatch(obj, qry)
Do Until rs.EOF
    Dim fld As Object
    Set dict = New Dictionary
    For Each fld In rs.Fields
        dict.Add fld.name, fld.value
    Next fld
    coll.Add dict
    rs.MoveNext
Loop
Set GetValueDict = coll
obj.CloseDbs obj
Set coll = Nothing
Set fld = Nothing
Set dict = Nothing
Set rs = Nothing
End Function



'==========================================================================

Public Function AddNewRecord(obj As Variant, dict As Dictionary)
Dim key As Variant
Dim iDataObject As New aclsDataObject
Set iDataObject = obj.Wrap(iDataObject)
Set rs = GetRecordsetMatch(iDataObject, iDataObject.Db)

rs.AddNew
On Error Resume Next
For Each key In dict.Keys
    rs.Fields(key) = dict(key)
    rs.Update
Next key
On Error GoTo 0

iDataObject.CloseDbs iDataObject

End Function

'==========================================================================

Public Sub ArchiveData(obj As aclsDataObject)
Dim iDataObject As New aclsDataObject
iDataObject.ArchiveData obj, ConstructArchiveCmd(obj)
Set iDataObject = Nothing
End Sub

Public Function ConstructArchiveCmd(obj As aclsDataObject) As String
Dim str As String

str = "INSERT INTO " & obj.Archive & " IN 'C:\Jared\POS\Access_Int\ReportsDB.accdb' SELECT * FROM " & obj.Db & ""

ConstructArchiveCmd = str
End Function





