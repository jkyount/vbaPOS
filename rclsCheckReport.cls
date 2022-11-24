VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rclsCheckReport"
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
Me.ColWidths = Array(31.5, 12)
Me.ColFormats = Array(xlGeneralFormat, "$0.00")
Me.ColAlignments = Array(xlLeft, xlCenter)
Me.HeaderRows = 6
Me.ReportCols = 2
Me.RowHeight = 24
Dim dict As New Dictionary
dict.Add "Check", 1
Me.RequiredParams = dict
Set dict = Nothing
End Sub


Public Function GetHeaderLines(iReport As aclsReport) As Collection
Dim coll As New Collection
Dim CheckDetails As New Dictionary
Dim arr As Variant
Dim qry As String
Dim iDataObj As New aclsDataObject
Dim iIndex As New zclsDailyCheckIndex
Set iDataObj = Me.DataSource.CfgDataSource(iIndex.Wrap(GetNewMatchObj(, iReport.check)))
Set CheckDetails = GetValueDict(iDataObj, ConstructMatchQuery(iDataObj))(1)



arr = Array("Bucci's Italian Cuisine", "")
coll.Add arr, "1"
PopFormatDict GetNewFormatObj(, 12, xlCenterAcrossSelection), CStr(coll.Count)

arr = Array("Rocky River  *=-?-=* (440)-331-5157", "", "", "", "")
coll.Add arr, "2"
PopFormatDict GetNewFormatObj(, 10, xlCenterAcrossSelection, 2, xlEdgeBottom), CStr(coll.Count)

arr = Array("Check:", iReport.check)
coll.Add arr, "3"
PopFormatDict GetNewFormatObj(True, 12, , , xlEdgeBottom), CStr(coll.Count)

arr = Array(CheckDetails("OrderName"), "")
coll.Add arr, "4"
PopFormatDict GetNewFormatObj(True, 12), CStr(coll.Count)

arr = Array(CheckDetails("Phone"), CheckDetails("PickupTime"))
coll.Add arr, "5"
PopFormatDict GetNewFormatObj(True, 12, , , xlEdgeBottom), CStr(coll.Count)

arr = Array("Server", CheckDetails("ServerName"))
coll.Add arr, "6"
PopFormatDict GetNewFormatObj(True, 12, , , xlEdgeBottom), CStr(coll.Count)

Set GetHeaderLines = coll
Set coll = Nothing

End Function


Public Function GetReportLines(iReport As aclsReport) As Collection
Dim CheckItems As New Collection, coll As New Collection
Set CheckItems = GetReportData(iReport)

Dim i As Integer, seat As Integer
Dim SeatIndicator As String
Dim arr As Variant
'-----------IF CHECKITEMS.COUNT = 0 THEN------------------------'
seat = CheckItems("Line1").seat
SeatIndicator = "- - - - - - - - - - - Seat " & seat & "- - - - - - - - - - -"
coll.Add Array(SeatIndicator, "")
PopFormatDict GetNewFormatObj(, , xlCenterAcrossSelection), CStr(coll.Count)
For i = 1 To CheckItems.Count
    If Not CheckItems("Line" & i).seat = seat Then
        seat = CheckItems("Line" & i).seat
        SeatIndicator = "- - - - - - - - - - - Seat " & seat & "- - - - - - - - - - -"
        coll.Add Array(SeatIndicator, "")
        PopFormatDict GetNewFormatObj(, , xlCenterAcrossSelection), CStr(coll.Count)
    End If
    arr = CheckItems("Line" & i).Data
    If CheckItems("Line" & i).primary = True Then
        coll.Add Array(" > " & arr(1)(0, 0), arr(1)(1, 0))
    End If
    If CheckItems("Line" & i).primary = False Then
        coll.Add Array("        " & arr(1)(0, 0), arr(1)(1, 0))
    End If
Next i

Dim Totals As New Dictionary
Dim iDataObj As New aclsDataObject
Dim iIndex As New zclsDailyCheckIndex
Set iDataObj = Me.DataSource.CfgDataSource(iIndex.Wrap(GetNewMatchObj("CheckNumber", iReport.check)))
Set Totals = GetValueDict(iDataObj, ConstructFilteredQuery(iDataObj, Array("Subtotal", "Tax", "Total")))(1)

coll.Add Array("-------------", "")
PopFormatDict GetNewFormatObj(, , xlCenterAcrossSelection), CStr(coll.Count)

coll.Add Array("Subtotal:   ", Format(Totals("Subtotal"), "$0.00"))
PopFormatDict GetNewFormatObj(, , xlRight), CStr(coll.Count)

coll.Add Array("Tax:   ", Format(Totals("Tax"), "$0.00"))
PopFormatDict GetNewFormatObj(, , xlRight), CStr(coll.Count)

coll.Add Array("Total:   ", Format(Totals("Total"), "$0.00"))
PopFormatDict GetNewFormatObj(, , xlRight), CStr(coll.Count)


AppendPayments iReport, coll
coll.Add Array("", "")
coll.Add Array("Thank you!    Visit us again soon.", "")
PopFormatDict GetNewFormatObj(, , xlCenterAcrossSelection), CStr(coll.Count)



Set GetReportLines = coll
Set coll = Nothing




SpclFormatting
End Function

Private Function GetReportData(iReport As aclsReport) As Collection
Dim iDetail As New zclsDailyCheckDetail, iDataObj As New aclsDataObject
Dim OrderBy As String
Set iDataObj = Me.DataSource.CfgDataSource(iDetail.Wrap(GetNewMatchObj("CheckNumber", iReport.check)))
OrderBy = "ORDER BY Seat ASC, LocalGroup ASC"
Set GetReportData = RemoveZeroPrice(DefineWriteData(iReport, SortCheckLines(GetCheckLines(iReport.check, GetValueDict(iDataObj, ConstructOrderedQuery(iDataObj, OrderBy))))))
End Function

Private Function RemoveZeroPrice(CheckItems As Collection) As Collection
Dim NonZeroPrice As New Collection
Dim arr As Variant, row As Integer, LinePrice As Double
Dim line As zclsCheckLines
row = 1
For Each line In CheckItems
    arr = line.Data
    LinePrice = arr(1)(1, 0)
    If Not LinePrice = 0 Then
        NonZeroPrice.Add line, "Line" & row
        row = row + 1
    End If
Next line
Set RemoveZeroPrice = NonZeroPrice
End Function

Private Function DefineWriteData(iReport As aclsReport, CheckItems As Collection) As Collection
Dim iDetail As New zclsDailyCheckDetail
Dim iDataObj As New aclsDataObject

Dim line As zclsCheckLines
For Each line In CheckItems
    Set iDataObj = Me.DataSource.CfgDataSource(iDetail.Wrap(GetNewMatchObj("CheckNumber", iReport.check, "LocalGroup", line.LocalGroup)))
    line.Data = FilteredMatch(iDataObj, "ItemName", "Price")
Next line
Set DefineWriteData = CheckItems
End Function

Private Sub AppendPayments(iReport As aclsReport, coll As Collection)
Dim Payments As New Dictionary
Dim iDataObj As New aclsDataObject
Dim iIndex As New zclsDailyCheckIndex
Set iDataObj = Me.DataSource.CfgDataSource(iIndex.Wrap(GetNewMatchObj("CheckNumber", iReport.check)))
Set Payments = GetValueDict(iDataObj, ConstructFilteredQuery(iDataObj, Array("Cash", "Charge", "GiftCert")))(1)
Dim key As Variant
For Each key In Payments.Keys
    If Not Payments(key) = 0 Then
        coll.Add Array(key & ":   ", Format(Payments(key), "$0.00"))
        PopFormatDict GetNewFormatObj(, , xlRight), CStr(coll.Count)
    End If
Next key
Dim TotalPayments As Double
TotalPayments = Payments("Cash") + Payments("Charge") + Payments("GiftCert")
If TotalPayments > 0 Then
    coll.Add Array("Amount Due:   $", Format(ValueMatch(iDataObj, "Total") - TotalPayments), "$0.00")
    PopFormatDict GetNewFormatObj(True, , xlRight), CStr(coll.Count)
End If
End Sub
Private Function SpclFormatting()
Dim Rg1 As range
Set Rg1 = Sheet13.range("TopLeftReportCell").Offset(Me.HeaderRows - 1, 0)
Dim rg2 As range
Set rg2 = Sheet13.range("TopLeftReportCell").Offset(500, Me.ReportCols - 1)

Sheet13.range(Rg1, rg2).Font.name = "Calibri"

Sheet13.range(Rg1, rg2).Font.Size = 16
End Function


