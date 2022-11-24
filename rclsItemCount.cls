VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rclsItemCount"
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
Me.ColWidths = Array(8.7, 18.5, 2.86, 3.71, 9.86)
Me.ColFormats = Array("0", xlGeneralFormat, "0", xlGeneralFormat, "0.00")
Me.ColAlignments = Array(xlLeft, xlLeft, xlRight, xlLeft, xlCenter)
Me.HeaderRows = 6
Me.ReportCols = 5
Me.RowHeight = 22
Dim dict As New Dictionary
Me.RequiredParams = dict
Set dict = Nothing
End Sub


Public Function GetHeaderLines(iReport As aclsReport) As Collection
Dim coll As New Collection
Dim arr As Variant


arr = Array("I t e m  C o u n t", "", "", "", "")
coll.Add arr, "1"
PopFormatDict GetNewFormatObj(, 12, xlCenterAcrossSelection), CStr(coll.Count)

arr = Array("Rocky River  *=-?-=* (440)-331-5157", "", "", "", "")
coll.Add arr, "2"
PopFormatDict GetNewFormatObj(, 10, xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)

arr = Array("From: " & Format(iReport.StartDate, "mm/dd/yyyy"), "", "", "", "")
coll.Add arr, "3"
PopFormatDict GetNewFormatObj(True, 12, xlCenterAcrossSelection), CStr(coll.Count)

arr = Array("To: " & Format(iReport.EndDate, "mm/dd/yyyy"), "", "", "", "")
coll.Add arr, "4"
PopFormatDict GetNewFormatObj(True, 12, xlCenterAcrossSelection), CStr(coll.Count)

arr = Array("Generated: " & Format(Now, "mm/dd/yy ttttt"), "", "", "", "")
coll.Add arr, "5"
PopFormatDict GetNewFormatObj(True, 11, xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)

arr = Array("", "", "", "", "")
coll.Add arr, "6"
PopFormatDict GetNewFormatObj(True, 12, xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)
Set GetHeaderLines = coll
Set coll = Nothing

End Function


Public Function GetPrintRange() As range
Dim Rg1 As range
Dim rg2 As range
Set Rg1 = Sheet13.range("TopLeftReportCell")
Set rg2 = Rg1.Offset(700, (Me.ReportCols - 1)).End(xlUp)

Set GetPrintRange = Sheet13.range(Rg1, rg2)
End Function


Public Function GetReportLines(iReport As aclsReport) As Collection
Dim rs As New ADODB.RecordSet
Dim iDetail As New zclsDailyCheckDetail
Dim iIndex As New zclsDailyCheckIndex
Dim DetailDataObj As New aclsDataObject
Dim IndexDataObj As New aclsDataObject
Set DetailDataObj = Me.DataSource.CfgDataSource(iDetail.Wrap(GetNewMatchObj))
Set IndexDataObj = Me.DataSource.CfgDataSource(iIndex.Wrap(GetNewMatchObj))

Dim qry As String



qry = "SELECT " & DetailDataObj.Db & ".ItemID, " & DetailDataObj.Db & ".ItemName, " & DetailDataObj.Db & ".Family, Count(" & DetailDataObj.Db & ".ItemID), Sum(" & DetailDataObj.Db & ".Price) FROM " & DetailDataObj.Db & " " _
        & "INNER JOIN " & IndexDataObj.Db & " " _
        & "ON " & DetailDataObj.Db & ".CheckNumber = " & IndexDataObj.Db & ".CheckNumber " _
        & "WHERE " & IndexDataObj.Db & ".CheckDate " & Me.DataSource.DateClause(iReport) & " " _
        & "GROUP BY " & DetailDataObj.Db & ".Family, " & DetailDataObj.Db & ".ItemID, " & DetailDataObj.Db & ".ItemName"

Set rs = iDetail.GetRs
DetailDataObj.OpenDbs DetailDataObj
rs.Source = qry
rs.Open
Dim arr As Variant
arr = rs.GetRows
DetailDataObj.CloseDbs DetailDataObj
Set iDetail = Nothing
Set iIndex = Nothing
Set IndexDataObj = Nothing
Set DetailDataObj = Nothing
Set rs = Nothing
Dim coll As New Collection
Dim i As Integer
coll.Add Array("Item", "Name", "#", "/", "$")
PopFormatDict GetNewFormatObj(, 14, xlCenter, , xlEdgeBottom), CStr(coll.Count)
Do Until i > UBound(arr, 2)
    coll.Add Array(arr(2, i), "", "", "", "")
    PopFormatDict GetNewFormatObj(True), CStr(coll.Count)
    Dim ttl As Double
    ttl = 0
    Do
        coll.Add Array(arr(0, i), Trim(arr(1, i)), arr(3, i), "/", arr(4, i))
        ttl = ttl + arr(4, i)
        i = i + 1
        If i > UBound(arr, 2) Then Exit Do
    Loop Until arr(2, i) <> arr(2, i - 1)
    coll.Add Array("", "", "", "Ttl:", CCur(ttl))
    PopFormatDict GetNewFormatObj(True, , , , xlEdgeBottom), CStr(coll.Count)
    Set GetReportLines = coll
Loop
Set coll = Nothing
End Function




