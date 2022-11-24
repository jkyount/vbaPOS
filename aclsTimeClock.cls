VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsTimeClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public dbTimeClock As New ADODB.Connection
Public rsTimeClock As New ADODB.RecordSet

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
GetDb = "TimeClock"
End Function

Public Function GetDbFile() As String
GetDbFile = "Employee"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbTimeClock
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsTimeClock
End Function

Public Function GetArchive() As String
GetArchive = "TimeClock"
End Function
Public Function GetArchiveDbFile() As String
GetArchiveDbFile = "Employee"
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



Public Sub ClockIn(ID As Long)
Dim k As Date
k = Format(Now, "mm/dd/yy HH:MM")
AddNewRecord Me, GetTimeClockDict(ID, k)
MsgBox ("Clocked in at" & k)

End Sub

Private Function GetTimeClockDict(EmployeeID As Long, k As Date) As Dictionary
Dim dict As New Dictionary
With dict
    .Add "IdNumber", EmployeeID
    .Add "ClockDate", Format(Now, "mmddyy")
    .Add "FirstName", ThisEmployee.FirstName
    .Add "LastName", ThisEmployee.LastName
    .Add "ServerNum", ThisEmployee.ServerNum
    .Add "In", Format(k, "mm/dd/yy HH:MM")
End With
Set GetTimeClockDict = dict
Set dict = Nothing

End Function

Public Sub ClockOut(EmployeeID As Long)
Dim ExtendQuery As String
Dim OutTime As Date, OutDate As Date, NullDate As Date
OutTime = Format(Now, "mm/dd/yy HH:MM")
OutDate = Format(Date, "mm/dd/yy")
NullDate = CDate("12:00:00 AM")
ExtendQuery = "Out AND IDNumber"
Dim RecordIndex As Integer
RecordIndex = ValueMatch(Wrap(GetNewMatchObj("Out", NullDate, "IDNumber", EmployeeID)), "ID")

Dim InDate As Date
InDate = Format(CDate(ValueMatch(Wrap(GetNewMatchObj("Out", NullDate, "IDNumber", EmployeeID)), "In")), "mm/dd/yy")
Dim EndOfDayClockOut As Date
EndOfDayClockOut = CDate(" 11:59 PM" & InDate)
Dim StartOfDayClockIn As Date
StartOfDayClockIn = CDate("12:01 AM" & OutDate)
If Not OutDate = InDate Then
    Update Wrap(GetNewUpdateObj("ID", RecordIndex, "Out", EndOfDayClockOut))
    Update Wrap(GetNewUpdateObj("ID", RecordIndex, "OutDate", Format(InDate, "mmddyy")))
    AddNewRecord Me, GetTimeClockDict(EmployeeID, StartOfDayClockIn)
    RecordIndex = ValueMatch(Wrap(GetNewMatchObj("Out", NullDate, "IDNumber", EmployeeID)), "ID")
End If
Update Wrap(GetNewUpdateObj("ID", RecordIndex, "Out", OutTime))
Update Wrap(GetNewUpdateObj("ID", RecordIndex, "OutDate", Format(OutDate, "mmddyy")))




MsgBox ("Clocked out at " & OutTime)



End Sub

Public Sub ClockOutAll()

Dim coll As New Collection
Dim rs As New ADODB.RecordSet

Dim ExtendQuery As String
Dim k As Date, NullDate As Date
NullDate = CDate("12:00:00 AM")
k = Format(Now, "HH:MM")
ExtendQuery = "IsNull(Out) AND NOT IDNumber"
Dim iDataObj As New aclsDataObject
Set iDataObj = Wrap(GetNewMatchObj("Out", NullDate, "NOT IDNumber", 0))
Set rs = GetRecordsetMatch(iDataObj, ConstructMatchQuery(iDataObj))
ClearCollection coll
Do
    coll.Add rs.Fields("IDNumber").value
    rs.MoveNext
Loop Until rs.EOF

Me.CloseDbs

Dim ID As Integer
For ID = 1 To coll.Count
    Me.ClockOut coll(ID)
Next ID
MsgBox ("All employees have been clocked out at " & k & ".")

End Sub
