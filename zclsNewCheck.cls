VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsNewCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit


Public dbNewCheck As New ADODB.Connection
Public rsNewCheck As New ADODB.RecordSet

Public Function GetDb() As String
GetDb = "DailyCheckIndex"
End Function
Public Function GetDbFile() As String
GetDbFile = "CheckDb"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbNewCheck
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsNewCheck
End Function

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

Public Function GetNextCheck() As String
GetNextCheck = pGetNextCheck
End Function

Public Function pGetNextCheck() As String
OpenDbs
rsNewCheck.Source = "CheckNumbers"
rsNewCheck.Open
Do Until rsNewCheck.Fields("Used").value = False
    rsNewCheck.MoveNext
    If rsNewCheck.EOF Then
        
        rsNewCheck.Close
        dbNewCheck.Close
        pGetNextCheck = ResetChecks
        Exit Function
    End If
Loop
pGetNextCheck = "ch" & rsNewCheck.Fields("CheckNumber").value
rsNewCheck.Fields("Used").value = True
rsNewCheck.Update
rsNewCheck.Close
dbNewCheck.Close
End Function
Private Function ResetChecks() As String
OpenDbs
rsNewCheck.Source = "CheckNumbers"
rsNewCheck.Open
rsNewCheck.MoveFirst
Do Until rsNewCheck.EOF
    rsNewCheck.Fields("Used").value = False
    rsNewCheck.MoveNext
Loop
rsNewCheck.UpdateBatch
rsNewCheck.Close
dbNewCheck.Close
ResetChecks = pGetNextCheck


End Function

Public Function PeekNextCheck() As String
PeekNextCheck = pPeekNextCheck
End Function

Public Function pPeekNextCheck() As String
OpenDbs
rsNewCheck.Source = "CheckNumbers"
rsNewCheck.Open
Do Until rsNewCheck.Fields("Used").value = False
    rsNewCheck.MoveNext
Loop
pPeekNextCheck = "ch" & rsNewCheck.Fields("CheckNumber").value
rsNewCheck.Close
dbNewCheck.Close
End Function

Public Function SetCheckInUse(check) As String
SetCheckInUse = pSetCheckInUse(check)
End Function
Private Function pSetCheckInUse(check) As String

OpenDbs
rsNewCheck.Source = "SELECT * FROM CheckNumbers WHERE CheckNumber = """ & check & """"
rsNewCheck.Open
rsNewCheck.Fields("Used").value = True
rsNewCheck.Update
rsNewCheck.Close
dbNewCheck.Close
End Function

Public Sub SetCheckUnused(check As String)

pSetCheckUnused check
End Sub

Private Sub pSetCheckUnused(check As String)
Dim CheckInt As Integer
CheckInt = CInt(right(check, (Len(check) - 2)))
OpenDbs
rsNewCheck.Source = "SELECT * FROM CheckNumbers WHERE CheckNumber = " & CheckInt & ""
rsNewCheck.Open
rsNewCheck.Fields("Used").value = False
rsNewCheck.Update
rsNewCheck.Close
dbNewCheck.Close
End Sub

Public Sub InitializeCheck(check As String)
pInitializeCheck check
End Sub
Private Sub pInitializeCheck(check As String)
OpenDbs
rsNewCheck.Source = "DailyCheckIndex"
rsNewCheck.Open
rsNewCheck.AddNew
rsNewCheck.Fields("CheckNumber").value = check
rsNewCheck.Update
rsNewCheck.Close
dbNewCheck.Close
End Sub

'Public Sub InitializeCheck(iOrder As aclsOrder)
'pInitializeCheck iOrder
'
'End Sub
'Private Sub pInitializeCheck(iOrder As aclsOrder)
'OpenDbs
'rsNewCheck.Source = "DailyCheckIndex"
'rsNewCheck.Open
'rsNewCheck.AddNew
'rsNewCheck.Fields("CheckNumber").value = iOrder.Check
'rsNewCheck.Update
'rsNewCheck.Close
'dbNewCheck.Close
'End Sub


