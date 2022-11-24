VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private pFirstName As String
Private pLastName As String
Private pIDNumber As Long
Private pClockedIn As Boolean
Private pServerNum As Integer
Private pAccentColor As XlRgbColor

Public dbEmployee As New ADODB.Connection
Public rsEmployee As New ADODB.RecordSet

Private Enum JobVals
    EmptyJob = 7
End Enum



Public Property Get Db() As String
Db = GetDb
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
GetDb = "Employee"
End Function

Public Function GetDbFile() As String
GetDbFile = "Employee"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbEmployee
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsEmployee
End Function


Public Property Let AccentColor(value As XlRgbColor)
pAccentColor = value
End Property
Public Property Get AccentColor() As XlRgbColor
AccentColor = Me.GetEmployeeParam("AccentColor")
End Property

Public Property Let FirstName(value As String)
pFirstName = value
End Property
Public Property Get FirstName() As String
FirstName = Me.GetEmployeeParam("FirstName")
End Property

Public Property Let ClockedIn(value As Boolean)
pClockedIn = value
End Property
Public Property Get ClockedIn() As Boolean
ClockedIn = Me.GetEmployeeParam("ClockedIn")
End Property

Public Property Let LastName(value As String)
pLastName = value
End Property
Public Property Get LastName() As String
LastName = Me.GetEmployeeParam("LastName")
End Property

Public Property Let IDNumber(value As Long)
pIDNumber = value
End Property
Public Property Get IDNumber() As Long
IDNumber = pIDNumber
End Property
Public Property Let ServerNum(value As Long)
pServerNum = value
End Property
Public Property Get ServerNum() As Long
ServerNum = Me.GetEmployeeParam("ServerNum")
End Property



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









'==========================================================================



Public Function GetAccentColor() As XlRgbColor
OpenDbs
rsEmployee.Source = "SELECT AccentColor FROM Employee WHERE ServerNum = " & Me.ServerNum & ""
rsEmployee.Open
If IsNull(rsEmployee.Fields("AccentColor").value) Then
    GetAccentColor = RGB(255, 195, 87)
    rsEmployee.Close
    dbEmployee.Close
    Exit Function
End If
GetAccentColor = rsEmployee.Fields("AccentColor").value
rsEmployee.Close
dbEmployee.Close
End Function

Public Sub SetAccentColor(color As XlRgbColor)
'OpenDbs
'rsEmployee.Source = "SELECT AccentColor FROM Employee WHERE ServerNum = " & Me.ServerNum & ""
'rsEmployee.Open
'rsEmployee.Fields("AccentColor").value = color
'rsEmployee.Update
'rsEmployee.Close
'dbEmployee.Close
Update Wrap(GetNewUpdateObj("ServerNum", Me.ServerNum, "AccentColor", color))

End Sub


Public Function GetEmployeeParam(param As String) As Variant
OpenDbs
rsEmployee.Source = "SELECT * FROM Employee WHERE IDNumber = " & Me.IDNumber & ""
rsEmployee.Open
GetEmployeeParam = rsEmployee.Fields(param).value
rsEmployee.Close
dbEmployee.Close
End Function

Public Function IsLoginValid(IDNumber As Long) As Boolean
IsLoginValid = pIsLoginValid(IDNumber)
End Function
Private Function pIsLoginValid(IDNumber As Long) As Boolean
OpenDbs
rsEmployee.Source = "Employee"
rsEmployee.Open
Do
    If rsEmployee.Fields("IDNumber") = IDNumber Then
        pIsLoginValid = True
        rsEmployee.Close
        dbEmployee.Close
        Exit Function
    End If
    rsEmployee.MoveNext
Loop Until rsEmployee.EOF
pIsLoginValid = False
rsEmployee.Close
dbEmployee.Close
End Function

Public Function Reset()
Me.FirstName = ""
Me.IDNumber = 0
Me.LastName = ""
Me.ServerNum = 0
End Function

Public Sub ClockIn(ID As Long)
Dim iEmployee As New zclsEmployee

Update iEmployee.Wrap(GetNewUpdateObj("IDNumber", ID, "ClockedIn", True))
End Sub

Public Sub ClockOut(ID As Long)
Dim iEmployee As New zclsEmployee
Update iEmployee.Wrap(GetNewUpdateObj("IDNumber", ID, "ClockedIn", False))
End Sub

Public Sub ClockOutAll()
Dim iEmployee As New zclsEmployee
Update iEmployee.Wrap(GetNewUpdateObj("ClockedIn", True, "ClockedIn", False))
End Sub

Public Function GetEmployees() As Variant
Dim arr As Variant
arr = FilteredMatch(Wrap(GetNewMatchObj("NOT ServerNum", 0)), "ServerNum", "FirstName", "LastName")

Dim i As Integer
For i = 1 To UBound(arr)
    arr(i)(0, 0) = arr(i)(0, 0) & " - " & arr(i)(1, 0) & " " & arr(i)(2, 0)
Next i
GetEmployees = arr

End Function

Public Function GetAllJobs() As Variant
Dim iDataObj As New aclsDataObject
Set iDataObj = Wrap(GetNewMatchObj("NOT ID", 0))
iDataObj.Db = "Job"

GetAllJobs = FilteredMatch(iDataObj, "Job")
End Function

Public Function GetJobs(ServerNum As Integer) As Variant

GetJobs = FilteredMatch(Wrap(GetNewMatchObj("ServerNum", ServerNum)), "Job1", "Job2", "Job3")

End Function

Public Function JobCodeToName(JobCode As Variant) As String
Dim iDataObj As New aclsDataObject
Set iDataObj = Wrap(GetNewMatchObj("ID", JobCode))
iDataObj.Db = "Job"
JobCodeToName = ValueMatch(iDataObj, "Job")
Set iDataObj = Nothing
End Function

Public Function FormatGetJobs(Jobs As Variant) As Variant
Dim arr As Variant
ReDim arr(1 To UBound(Jobs(1)) + 1)
Dim arr2(0, 0) As Variant
Dim i As Integer
For i = 1 To UBound(Jobs(1)) + 1
    arr2(0, 0) = Jobs(1)(i - 1, 0)
    arr(i) = arr2
Next i
FormatGetJobs = arr
End Function

Public Function JobNameToCode(Job As String) As Integer
Dim iDataObj As New aclsDataObject
Set iDataObj = Wrap(Wrap(GetNewMatchObj("Job", Job)))
iDataObj.Db = "Job"
JobNameToCode = ValueMatch(iDataObj, "ID")
Set iDataObj = Nothing
End Function

Public Sub AddNewJob(JobCode As Integer, PayRate As Double, ServerNum As Integer)
Dim NewJobIndex As String
NewJobIndex = GetJobIndex(ServerNum, EmptyJob)
Update Wrap(GetNewUpdateObj("ServerNum", ServerNum, "Job" & NewJobIndex, JobCode))
Update Wrap(GetNewUpdateObj("ServerNum", ServerNum, "Payrate" & NewJobIndex, PayRate))
End Sub

Public Sub RemoveJob(JobCode As Integer, ServerNum As Integer)
Dim JobIndex As String
JobIndex = GetJobIndex(ServerNum, JobCode)
Update Wrap(GetNewUpdateObj("ServerNum", ServerNum, "Job" & JobIndex, EmptyJob))
Update Wrap(GetNewUpdateObj("ServerNum", ServerNum, "Payrate" & JobIndex, 0))
End Sub

Public Function GetJobDict(ServerNum As Integer) As Dictionary
Dim dict As New Dictionary
Set dict = GetValueDict(Wrap(GetNewMatchObj("ServerNum", ServerNum)))(1)
Dim JobDict As New Dictionary
Dim JunkDict As Dictionary
Dim i As Integer
For i = 1 To 3
    Set JunkDict = New Dictionary
    JunkDict.Add "Job", dict("Job" & i)
    JunkDict.Add "Payrate", dict("Payrate" & i)
    JobDict.Add "Job" & i, JunkDict
Next i
Set GetJobDict = JobDict
Set dict = Nothing
Set JobDict = Nothing
Set JunkDict = Nothing
End Function

Private Function GetJobIndex(ServerNum As Integer, TargetJobCode As Integer) As String
Dim JobDict As New Dictionary
Set JobDict = GetJobDict(ServerNum)
Dim key As Variant
For Each key In JobDict
    If JobDict(key)("Job") = TargetJobCode Then
        GetJobIndex = right(key, 1)
        Exit For
        Exit Function
    End If
Next key
End Function




