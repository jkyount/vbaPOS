VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ReportRange As range
Dim ReportCols As Integer
Dim ColWidths() As Variant
Dim ColFormats() As Variant
Dim ColAlignments() As Variant
Dim RowHeight As Integer
Dim HeaderRows As Integer
Dim TopLeftReportCell As range

Public PrintReportClicked As Boolean

Private pStartDate As Date
Private pEndDate As Date
Private pServerNum As Integer
Private pCheck As String

Public Property Let ServerNum(value As Integer)
pServerNum = value
End Property
Public Property Get ServerNum() As Integer
ServerNum = pServerNum
End Property

Public Property Let check(value As String)
pCheck = value
End Property
Public Property Get check() As String
check = pCheck
End Property

Public Property Let EndDate(value As Date)
pEndDate = value
End Property
Public Property Get EndDate() As Date
EndDate = pEndDate
End Property
Public Property Let StartDate(value As Date)
pStartDate = value
End Property
Public Property Get StartDate() As Date
StartDate = pStartDate
End Property

Public Sub RunReport(ReportType As Variant)
ClearFormatDict
InitializeReport ReportType
FormatReportRange ReportType
WriteHeader ReportType
FormatHeader
ClearFormatDict

WriteLines ReportType
FormatBody

ViewReport ReportType
ClearFormatDict

End Sub



Public Sub InitializeReport(ReportType As Variant)
ColWidths() = ReportType.ColWidths
ReportCols = ReportType.ReportCols
HeaderRows = ReportType.HeaderRows
RowHeight = ReportType.RowHeight
ColFormats() = ReportType.ColFormats
ColAlignments() = ReportType.ColAlignments
End Sub

Public Sub FormatReportRange(ReportType As Variant)
Set TopLeftReportCell = Sheet13.range("TopLeftReportCell")
Set ReportRange = Sheet13.range(TopLeftReportCell, TopLeftReportCell.Offset(700, 10))
ReportRange.Clear
Set ReportRange = Sheet13.range(TopLeftReportCell, TopLeftReportCell.Offset(700, ReportCols - 1))
Dim i As Integer
With ReportRange
    .Rows.RowHeight = RowHeight
    .BorderAround xlSolid, 2
    .Font.name = "Lucida Sans Unicode"
    For i = 1 To ReportCols
        .Columns(i).ColumnWidth = ColWidths(i - 1)
        .Columns(i).NumberFormat = ColFormats(i - 1)
        .Columns(i).HorizontalAlignment = ColAlignments(i - 1)
        
    Next i

    End With
End Sub
Public Sub WriteHeader(ReportType As Variant)
Dim ReportHeader As Collection
Set ReportHeader = ReportType.GetHeaderLines(Me)
Dim i As Integer
For i = 1 To HeaderRows
    ReportRange.Rows(i).value = ReportHeader(CStr(i))
Next i
End Sub

Public Sub FormatHeader()
Dim dict As Dictionary
Dim key As Variant
Set dict = GetFormatDict
For Each key In dict.Keys
    ApplyFormat CInt(key), dict(key)
Next key
End Sub
Public Sub FormatBody()
Dim dict As Dictionary
Dim key As Variant
Set dict = GetFormatDict
For Each key In dict.Keys
    ApplyFormat (CInt(key) + HeaderRows), dict(key)
Next key
End Sub

Private Sub ApplyFormat(row As Integer, iFmt As aclsFormatObj)
ReportRange.Rows(row).Font.Bold = iFmt.FmtBold
If Not iFmt.FmtFontSize = 0 Then
    ReportRange.Rows(row).Font.Size = iFmt.FmtFontSize
End If
If Not iFmt.FmtAlign = "" Then
    ReportRange.Rows(row).HorizontalAlignment = iFmt.FmtAlign
End If
If Not iFmt.FmtBorderStyle = "" Then
    ReportRange.Rows(row).Borders(iFmt.FmtBorderStyle).weight = xlThin
    
End If
If Not iFmt.FmtBorderWeight = 0 Then
    ReportRange.Rows(row).Borders.weight = iFmt.FmtBorderWeight
End If

End Sub

Private Sub WriteLines(ReportType As Variant)
Dim coll As Collection
Set coll = ReportType.GetReportLines(Me)
Dim i As Integer
For i = 1 To coll.Count
    ReportRange.Rows(i + HeaderRows).value = coll(i)
Next i

ReportRange.Rows(coll.Count + HeaderRows + 2).Columns(ReportCols).value = ":::"

End Sub



Public Sub PrintReport(ReportType As Variant)
Dim PrintRange As range
Set PrintRange = GetPrintRange(ReportType)
PrintRange.Interior.color = rgbWhite
PrintRange.PrintOut ActivePrinter:="EPSON TM-T88V Receipt"
End Sub

Private Function GetPrintRange(ReportType As Variant) As range
Dim Rg1 As range
Dim rg2 As range
Set Rg1 = Sheet13.range("TopLeftReportCell")
Set rg2 = Rg1.Offset(700, ReportType.ReportCols - 1).End(xlUp)
Set GetPrintRange = range(Rg1, rg2)
End Function

Public Sub ViewReport(ReportType As Variant)
Dim ThisSheet As Worksheet
Set ThisSheet = ActiveSheet
Sheet13.Activate
Dim PrintRange As range
Set PrintRange = GetPrintRange(ReportType)
PrintRange.Copy
'Application.ScreenUpdating = True
range("AL1").Select
ActiveSheet.Pictures.Paste Link:=True
ActiveSheet.Pictures.Select
Application.CutCopyMode = False
Dim myChart As String, myPicture As String
Dim picWidth As Long, picHeight As Long
'Range("CC").Select
Application.ScreenUpdating = False
myPicture = Selection.name
picHeight = Selection.ShapeRange.Height
picWidth = Selection.ShapeRange.Width
Charts.Add
ActiveChart.Location xlLocationAsObject, "ReportSpace"
Selection.Border.LineStyle = 0
myChart = Selection.name & " " & Split(ActiveChart.name, " ")(2)
ActiveSheet.Shapes(myChart).Width = picWidth
ActiveSheet.Shapes(myChart).Height = picHeight
ActiveSheet.Shapes(myPicture).Copy
ActiveChart.ChartArea.Select
ActiveChart.Paste
ActiveSheet.ChartObjects(1).Chart.Export Filename:="C:\Jared\RepImg.jpg", FilterName:="jpg"

ActiveSheet.Shapes(myChart).Cut
'Application.ScreenUpdating = True
RepViewer.Rep.Width = picWidth + 30
RepViewer.Rep.Height = picHeight + 7
RepViewer.Frame1.Width = RepViewer.Rep.Width + 30
RepViewer.Frame1.Left = (ActiveWindow.Width / 2) - (RepViewer.Frame1.Width / 2)
RepViewer.Frame1.ScrollHeight = RepViewer.Rep.Height + 20
RepViewer.Width = ActiveWindow.Width
RepViewer.Height = ActiveWindow.Height
RepViewer.PrintReport.Left = (ActiveWindow.Width / 2) - (RepViewer.PrintReport.Width / 2)
RepViewer.Rep.Picture = LoadPicture("C:\Jared\RepImg.jpg")
ActiveSheet.Pictures.Delete
Application.ScreenUpdating = True
ThisSheet.Activate
RepViewer.ShowRepViewer ReportType


End Sub


