VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rclsCurrentClockIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPrintRange As range
Private pReportCols As Integer
Private pColWidths() As Variant
Private pRowHeight As Integer
Private pHeaderRows As Integer
Private pHeaderLines As Collection
Private pColFormats As Variant
Private pColAlignments As Variant
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

Public Property Let ReportCols(value As Integer)
pReportCols = value
End Property
Public Property Get ReportCols() As Integer
ReportCols = pReportCols
End Property

Public Property Let ColWidths(value As Variant)
pColWidths = value
End Property
Public Property Get ColWidths() As Variant
ColWidths = pColWidths
End Property

Public Property Let RowHeight(value As Integer)
pRowHeight = value
End Property
Public Property Get RowHeight() As Integer
RowHeight = pRowHeight
End Property

Public Property Let HeaderRows(value As Integer)
pHeaderRows = value
End Property
Public Property Get HeaderRows() As Integer
HeaderRows = pHeaderRows
End Property

Public Property Let ColFormats(value As Variant)
pColFormats = value
End Property
Public Property Get ColFormats() As Variant
ColFormats = pColFormats
End Property

Public Property Let ColAlignments(value As Variant)
pColAlignments = value
End Property
Public Property Get ColAlignments() As Variant
ColAlignments = pColAlignments
End Property


Private Sub Class_Initialize()
Me.ColWidths = Array(3.7, 18, 22)
Me.ColFormats = Array("0", xlGeneralFormat, "mm/dd/yyyy hh:mm")
Me.ColAlignments = Array(xlLeft, xlLeft, xlLeft)
Me.HeaderRows = 3
Me.ReportCols = 3
Me.RowHeight = 24
Dim dict As New Dictionary
Me.RequiredParams = dict
Set dict = Nothing
End Sub


Public Function GetHeaderLines(iReport As aclsReport) As Collection
Dim coll As New Collection
Dim arr As Variant

arr = Array("C u r r e n t  C l o c k - I n", "", "")
coll.Add arr, "1"
PopFormatDict GetNewFormatObj(, 12, xlCenterAcrossSelection), CStr(coll.Count)

arr = Array("Rocky River  *=-?-=* (440)-331-5157", "", "")
coll.Add arr, "2"
PopFormatDict GetNewFormatObj(, 10, xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)

arr = Array("Generated: " & Format(Now, "mm/dd/yy ttttt"), "", "")
coll.Add arr, "3"
PopFormatDict GetNewFormatObj(True, 11, xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)

Set GetHeaderLines = coll
Set coll = Nothing
End Function

Public Function GetReportLines(iReport As aclsReport) As Collection
Dim rs As New ADODB.RecordSet
Dim iEmployee As New zclsEmployee
Dim qry As String
qry = "SELECT TimeClock.ServerNum, TimeClock.FirstName, TimeClock.LastName, TimeClock.In, TimeClock.ClockDate FROM TimeClock " _
        & "INNER JOIN Employee " _
        & "ON TimeClock.ServerNum = Employee.ServerNum " _
        & "WHERE Employee.ClockedIn = True AND IsNull(TimeClock.Out)"
Set rs = GetRecordsetMatch(iEmployee.Wrap(GetNewMatchObj), qry)
If rs.RecordCount = 0 Then
    Dim NoRecordsColl As New Collection
    NoRecordsColl.Add Array("ID", "Name", "Clocked In")
    PopFormatDict GetNewFormatObj(, 11, xlCenter, , xlEdgeBottom), CStr(NoRecordsColl.Count)
    NoRecordsColl.Add Array("No", "Current", "Clock Ins")
    Set GetReportLines = NoRecordsColl
    Exit Function
End If
Dim arr As Variant
arr = rs.GetRows
iEmployee.CloseDbs
Set iEmployee = Nothing
Set rs = Nothing
Dim coll As New Collection
Dim i As Integer
coll.Add Array("ID", "Name", "Clocked In")
PopFormatDict GetNewFormatObj(, 11, xlCenter, , xlEdgeBottom), CStr(coll.Count)
Do Until i > UBound(arr, 2)
    coll.Add Array(arr(0, i), arr(1, i) & " " & arr(2, i), arr(3, i))
i = i + 1
Loop
Set GetReportLines = coll
Set coll = Nothing
End Function

