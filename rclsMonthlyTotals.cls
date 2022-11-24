VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rclsMonthlyTotals"
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




'Public Property Let PrintRange(value As Range)
'Set pPrintRange = value
'End Property
'Public Property Get PrintRange() As Range
'Set PrintRange = GetPrintRange
'End Property

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
Me.ColWidths = Array(3.57, 14, 12, 12, 12, 12, 12, 12, 15)
Me.ColFormats = Array("0", xlGeneralFormat, "0.00", "0.00", "0.00", "0.00", "0.00", "0.00", "0.00")
Me.ColAlignments = Array(xlRight, xlLeft, xlCenter, xlCenter, xlCenter, xlCenter, xlCenter, xlCenter, xlCenter)
Me.HeaderRows = 2
Me.ReportCols = 9
Me.RowHeight = 22
Dim dict As New Dictionary
Me.RequiredParams = dict
Set dict = Nothing
End Sub


Public Function GetHeaderLines(iReport As aclsReport) As Collection
Dim coll As New Collection
Dim arr As Variant


arr = Array("M o n t h l y  T o t a l s", "", "", "", "", "", "", "", "")
coll.Add arr, "1"
PopFormatDict GetNewFormatObj(, 12, xlCenterAcrossSelection), CStr(coll.Count)

arr = Array(Format(iReport.StartDate, "mmmm"), "", "", "", "", "", "", "", "")
coll.Add arr, "2"
PopFormatDict GetNewFormatObj(, 16, xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)
Set GetHeaderLines = coll
Set coll = Nothing

End Function



Public Function GetReportLines(iReport As aclsReport) As Collection
Dim StartDay As Integer, EndDay As Integer
Dim StartMonth As Integer, EndMonth As Integer
Dim StartYear As Integer, EndYear As Integer
StartDay = Day(iReport.StartDate)
StartMonth = Month(iReport.StartDate)
StartYear = Year(iReport.StartDate)
EndDay = Day(iReport.EndDate)
EndMonth = Month(iReport.EndDate)
EndYear = Year(iReport.EndDate)
Dim dict As New Dictionary
Dim dt As Date
Dim ReportLines As New Collection

'--------------------------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------------------------'
ReportLines.Add Array("", "", "Carryout", "Food In", "Beer", "Wine", "Taxable", "Tax", "Total Sales")
PopFormatDict GetNewFormatObj(True, 12, xlCenter, , xlEdgeBottom), CStr(ReportLines.Count)
Do
    dt = DateSerial(StartYear, StartMonth, StartDay)
    If Not Format(dt, "dddd") = "Monday" Then
        Set dict = GetDailyValues(dt, iReport)
        ReportLines.Add Array(CStr(StartDay), Format(dt, "dddd"), dict("Carryout"), dict("FoodIn"), dict("Beer"), dict("Wine"), dict("Taxable"), dict("Tax"), dict("SubTotal") + dict("Tax"))
        PopFormatDict GetNewFormatObj(, 12, , , xlEdgeBottom), CStr(ReportLines.Count)
    End If
    If Format(dt, "dddd") = "Monday" Then
        ReportLines.Add Array(CStr(StartDay), Format(dt, "dddd"), "----", "----", "----", "----", "----", "----", "----")
        PopFormatDict GetNewFormatObj(, 12, , , xlEdgeBottom), CStr(ReportLines.Count)
    End If
StartDay = StartDay + 1
If StartDay > Day(DateSerial(StartYear, StartMonth + 1, 1) - 1) Then
    StartMonth = StartMonth + 1
        If StartMonth > Month(DateSerial(StartYear + 1, 1, 1) - 1) Then
            StartYear = StartYear + 1
            StartMonth = 1
        End If
    StartDay = 1
End If
Loop Until DateSerial(StartYear, StartMonth, StartDay) > DateSerial(EndYear, EndMonth, EndDay)
Debug.Print DateSerial(StartYear, StartMonth, StartDay)
Debug.Print DateSerial(EndYear, EndMonth, EndDay)
Set GetReportLines = ReportLines
'--------------------------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------------------------'

Set dict = Nothing
Set ReportLines = Nothing

SpclFormatting


End Function

Public Function GetDailyValues(vDate As Date, iReport As aclsReport) As Dictionary
Dim iIndex As New zclsDailyCheckIndex
Dim categories As New Collection
Set categories = GetCategories
Dim member As Variant
Dim val As Variant
Dim dict As New Dictionary
Dim qryDate As Long
'--------------------------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------------------------'
For Each member In categories
    qryDate = Format(vDate, "mmddyy")
    val = FilteredMatch(Me.DataSource.CfgDataSource(iIndex.Wrap(GetNewMatchObj("CheckDate", qryDate))), "Sum(" & member & ")")(1)(0, 0)
    If IsNull(val) Then val = 0
    dict.Add member, val
Next member
Set GetDailyValues = dict
'--------------------------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------------------------'
Set iIndex = Nothing
Set categories = Nothing
Set dict = Nothing
End Function

Private Function GetCategories() As Collection
Dim coll As New Collection
coll.Add "Carryout"
coll.Add "FoodIn"
coll.Add "Beer"
coll.Add "Wine"
coll.Add "Taxable"
coll.Add "Tax"
coll.Add "SubTotal"
Set GetCategories = coll
End Function

Public Function SpclFormatting()
Dim Rg1 As range
Set Rg1 = Sheet13.range("TopLeftReportCell").Offset(Me.HeaderRows - 1, 0)
Dim rg2 As range
Set rg2 = Sheet13.range("TopLeftReportCell").Offset(50, Me.ReportCols - 1)

Sheet13.range(Rg1, rg2).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Sheet13.range(Rg1, rg2).Borders(xlInsideVertical).LineStyle = xlContinuous

End Function
