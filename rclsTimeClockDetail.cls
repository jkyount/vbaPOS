VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rclsTimeClockDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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




Public Property Let PrintRange(value As range)
Set pPrintRange = value
End Property
Public Property Get PrintRange() As range
Set PrintRange = GetPrintRange
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
Me.ColWidths = Array(4.7, 15.86, 12, 12, 12)
Me.ColFormats = Array("0", xlGeneralFormat, "hh:mm:ssss", "hh:mm:ssss", "0.00")
Me.ColAlignments = Array(xlLeft, xlLeft, xlLeft, xlLeft, xlLeft)
Me.HeaderRows = 5
Me.ReportCols = 5
Me.RowHeight = 24
Dim dict As New Dictionary
Me.RequiredParams = dict
Set dict = Nothing
End Sub


Public Function GetHeaderLines(iReport As aclsReport) As Collection
Dim coll As New Collection
Dim arr As Variant


arr = Array("T i m e  C l o c k  D e t a i l", "", "", "", "")
coll.Add arr, "1"
PopFormatDict GetNewFormatObj(, 12, xlCenterAcrossSelection), CStr(coll.Count)

arr = Array("Rocky River  *=-?-=* (440)-331-5157", "", "", "", "")
coll.Add arr, "2"
PopFormatDict GetNewFormatObj(, 10, xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)

arr = Array("From: " & Format(iReport.StartDate, "mm/dd/yyyy"), "", "", "", "")
coll.Add arr, "3"
PopFormatDict GetNewFormatObj(, 10, xlCenterAcrossSelection), CStr(coll.Count)

arr = Array("To: " & Format(iReport.EndDate, "mm/dd/yyyy"), "", "", "", "")
coll.Add arr, "4"
PopFormatDict GetNewFormatObj(, 10, xlCenterAcrossSelection), CStr(coll.Count)

arr = Array("Generated: " & Format(Now, "mm/dd/yy ttttt"), "", "", "", "")
coll.Add arr, "5"
PopFormatDict GetNewFormatObj(True, 11, xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)

Set GetHeaderLines = coll
Set coll = Nothing

End Function


Public Function GetReportLines(iReport As aclsReport) As Collection
Dim rs As New ADODB.RecordSet
Dim iEmployee As New zclsEmployee


Dim qry As String

'qry = "SELECT " & DetailDataObj.Db & ".ItemID, " & DetailDataObj.Db & ".ItemName, " & DetailDataObj.Db & ".Family, Count(" & DetailDataObj.Db & ".ItemID), Sum(" & DetailDataObj.Db & ".Price) FROM " & DetailDataObj.Db & " " _
'        & "INNER JOIN " & IndexDataObj.Db & " " _
'        & "ON " & DetailDataObj.Db & ".CheckNumber = " & IndexDataObj.Db & ".CheckNumber " _
'        & "WHERE " & IndexDataObj.Db & ".CheckDate between " & Format(iReport.StartDate, "mmddyy") & " and " & Format(iReport.EndDate, "mmddyy") & " " _
'        & "GROUP BY " & DetailDataObj.Db & ".Family, " & DetailDataObj.Db & ".ItemID, " & DetailDataObj.Db & ".ItemName"


qry = "SELECT TimeClock.ServerNum, TimeClock.FirstName, TimeClock.LastName, TimeClock.In, TimeClock.Out, TimeClock.ClockDate FROM TimeClock " _
        & "INNER JOIN Employee " _
        & "ON TimeClock.ServerNum = Employee.ServerNum " _
        & "WHERE TimeClock.ClockDate BETWEEN " & Format(iReport.StartDate, "mmddyy") & " AND " & Format(iReport.EndDate, "mmddyy") & " " _
        & "ORDER BY TimeClock.ClockDate ASC"
Set rs = GetRecordsetMatch(iEmployee.Wrap(GetNewMatchObj), qry)




Dim coll As New Collection

If rs.RecordCount = 0 Then
    coll.Add Array("ID", "Name", "In", "Out", "Hrs")
    PopFormatDict GetNewFormatObj(, 11, xlCenter, , xlEdgeBottom), CStr(coll.Count)
    coll.Add Array("No", "Records", "Available", "", "")
    Set GetReportLines = coll
    Exit Function
End If
Dim arr As Variant
arr = rs.GetRows
iEmployee.CloseDbs
Set iEmployee = Nothing
Set rs = Nothing

Dim i As Integer
coll.Add Array("ID", "Name", "In", "Out", "Hrs")
PopFormatDict GetNewFormatObj(, 11, xlCenter, , xlEdgeBottom), CStr(coll.Count)


Do Until i > UBound(arr, 2)
    coll.Add Array(Format(arr(3, i), "dddd") & " " & Format(arr(3, i), "mm/dd/yyyy"), "", "", "", "")
    PopFormatDict GetNewFormatObj(True), CStr(coll.Count)
    Dim ttl As Double
    Dim ClockInTime As Variant
    
    ttl = 0
    Do
        If Not IsNull(arr(4, i)) Then ClockInTime = (DateDiff("n", arr(3, i), arr(4, i)) / 60)
        If IsNull(arr(4, i)) Then ClockInTime = 0
        coll.Add Array(arr(0, i), arr(1, i) & " " & arr(2, i), arr(3, i), arr(4, i), ClockInTime)
        ttl = ttl + ClockInTime
        i = i + 1
        If i > UBound(arr, 2) Then Exit Do
    Loop Until Format(arr(3, i), "mmddyy") <> Format(arr(3, i - 1), "mmddyy")
    coll.Add Array("", "", "", "Ttl:", CCur(ttl))
    PopFormatDict GetNewFormatObj(True, , , , xlEdgeBottom), CStr(coll.Count)
    Set GetReportLines = coll
Loop
Set coll = Nothing




'Do Until i > UBound(arr, 2)
'    coll.Add Array(arr(0, i), arr(1, i) & " " & arr(2, i), arr(3, i))
'i = i + 1
'Loop
'Set GetReportLines = coll
'Set coll = Nothing
End Function

