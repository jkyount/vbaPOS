VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rclsSalesSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pRequiredParams As Dictionary
Private pDataSource As Variant

Public Property Let RequiredParams(value As Dictionary)
Set pRequiredParams = value
End Property
Public Property Get RequiredParams() As Dictionary
Set RequiredParams = pRequiredParams
End Property

Public Property Let DataSource(value As Variant)
Set pDataSource = value
End Property
Public Property Get DataSource() As Variant
Set DataSource = pDataSource
End Property


Private Sub Class_Initialize()
Dim dict As New Dictionary
Me.RequiredParams = dict
End Sub


Public Function GetHeaderLines(iReport As aclsReport) As Collection
Dim coll As New Collection
Dim arr As Variant

'-----------------------------------------------------------------------'
arr = Array("S A L E S  S U M M A R Y", "", "", "", "")
coll.Add arr, "1"
PopFormatDict GetNewFormatObj(, 12, xlCenterAcrossSelection), CStr(coll.Count)
'--------------------------------------------------------'
arr = Array("Rocky River  *=-?-=* (440)-331-5157", "", "", "", "")
coll.Add arr, "2"
PopFormatDict GetNewFormatObj(, 10, xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)
'--------------------------------------------------------'
arr = Array("", "", "", "", "")
coll.Add arr, "3"
PopFormatDict GetNewFormatObj(True, 12, xlCenterAcrossSelection), CStr(coll.Count)
'---------------------------------------------------------'
arr = Array("From: " & Format(iReport.StartDate, "dddd mm/dd/yyyy"), "", "", "", "")
coll.Add arr, "4"
PopFormatDict GetNewFormatObj(True, 12, xlCenterAcrossSelection), CStr(coll.Count)
'-----------------------------------------------------------'
arr = Array("To: " & Format(iReport.EndDate, "dddd mm/dd/yyyy"), "", "", "", "")
coll.Add arr, "5"
PopFormatDict GetNewFormatObj(True, 12, xlCenterAcrossSelection), CStr(coll.Count)
'------------------------------------------------------------'
arr = Array("Generated: " & Format(Now, "mm/dd/yy ttttt"), "", "", "", "")
coll.Add arr, "6"
PopFormatDict GetNewFormatObj(True, 11, xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)
'------------------------------------------------------------'
Set GetHeaderLines = coll
Set coll = Nothing
End Function


Public Function GetClosedChecks(iReport As aclsReport) As Collection
   Set GetClosedChecks = New Collection
    Dim dict As New Dictionary
        dict.Add "CheckNumber", "No Closed Chks"
        dict.Add "Total", "0"
        dict.Add "Cash", 0
        dict.Add "Charge", 0
        dict.Add "GiftCert", 0
        dict.Add "ChargeTip", 0
    GetClosedChecks.Add dict
End Function



Public Function GetOpenChecks(iReport As aclsReport) As Collection
    Set GetOpenChecks = New Collection
    Dim dict As New Dictionary
        dict.Add "CheckNumber", "No Open Chks"
        dict.Add "Total", "0"
        dict.Add "Cash", 0
        dict.Add "Charge", 0
        dict.Add "GiftCert", 0
        dict.Add "ChargeTip", 0
    GetOpenChecks.Add dict
End Function


Public Function GetStatPair(stat As String, iReport As aclsReport) As Variant
Dim iIndex As New zclsDailyCheckIndex
Dim iDataObj As New aclsDataObject
Set iDataObj = Me.DataSource.CfgDataSource(iIndex.Wrap(GetNewMatchObj))
Dim qry As String
Dim rs As New ADODB.RecordSet
Set rs = iIndex.GetRs
iDataObj.OpenDbs iDataObj
qry = "SELECT Sum(" & stat & "), Count(" & stat & ") FROM " & iDataObj.Db & " WHERE CheckDate " & Me.DataSource.DateClause(iReport) & "  AND NOT " & stat & " = 0"
rs.Source = qry
On Error GoTo EH
    rs.Open
On Error GoTo 0
'=========================================='
GetStatPair = NullToZero(Array(rs.Fields(0).value, rs.Fields(1).value))
'============================================='
iDataObj.CloseDbs iDataObj
Set iIndex = Nothing
Set iDataObj = Nothing
Set rs = Nothing
Exit Function
EH:
iIndex.GetConn.Close
Set iIndex = Nothing
Set iDataObj = Nothing
Set rs = Nothing
End Function



