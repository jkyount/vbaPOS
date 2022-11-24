VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tclsSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pReportType As Variant
Private DataOrder As Variant
Private BlankRows() As Variant
Private pDataSource As Variant

Private pReportCols As Integer
Private pColWidths() As Variant
Private pRowHeight As Integer
Private pHeaderRows As Integer
Private pColFormats As Variant
Private pColAlignments As Variant


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
Public Property Let DataSource(value As Variant)
Me.ReportType.DataSource = value
End Property
Public Property Get DataSource() As Variant
Set DataSource = Me.ReportType.DataSource
End Property

Public Property Get RequiredParams() As Dictionary
Set RequiredParams = Me.ReportType.RequiredParams
End Property

Public Property Let ReportType(value As Variant)
Set pReportType = value
End Property
Public Property Get ReportType() As Variant
Set ReportType = pReportType
End Property

Public Function GetHeaderLines(iReport As aclsReport) As Collection
Set GetHeaderLines = Me.ReportType.GetHeaderLines(iReport)
End Function

Private Sub Class_Initialize()
Me.ColWidths = Array(15, 9.57, 1.29, 3.71, 11.3)
Me.ColFormats = Array(xlGeneralFormat, "0.00", xlGeneralFormat, "0", xlGeneralFormat)
Me.ColAlignments = Array(xlLeft, xlLeft, xlLeft, xlLeft, xlLeft)
Me.HeaderRows = 6
Me.ReportCols = 5
Me.RowHeight = 24
End Sub

Public Function GetReportLines(iReport As aclsReport) As Collection
BlankRows = GetBlankRows
DataOrder = GetDataOrder
Dim coll As New Collection
Dim StatPairs As New Dictionary
Dim i As Integer
coll.Add Array("", "$", "/", "#", "")
PopFormatDict GetNewFormatObj(, 14, xlCenter), CStr(coll.Count)
Set StatPairs = GetStatPairs(iReport)
Dim key As Variant
Dim arr As Variant
For Each key In StatPairs.Keys
    arr = Array(key, StatPairs(key)(1), "/", StatPairs(key)(2), "")
    If key Like "Blank*" Then
        arr = Array("//////////////////////////////////", "", "", "", "")
        PopFormatDict GetNewFormatObj(, , , , xlEdgeBottom), CStr(coll.Count)
        PopFormatDict GetNewFormatObj(True, , xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count + 1)
    End If
    coll.Add arr
Next key
coll.Add Array("Transactions", StatPairs("SubTotal")(2), "", "", "")
If StatPairs("SubTotal")(2) = 0 Then
    coll.Add Array("Avg", 0, "", "", "", "")
End If
If Not StatPairs("SubTotal")(2) = 0 Then
    coll.Add Array("Avg", StatPairs("SubTotal")(1) / StatPairs("SubTotal")(2), "", "", "", "")
End If
coll.Add Array("Cash Due", StatPairs("Cash")(1) - StatPairs("ChargeTip")(1), "", "", "")
PopFormatDict GetNewFormatObj(, , , , xlEdgeBottom), CStr(coll.Count)

ClosedCheckSummary iReport, coll
OpenCheckSummary iReport, coll
Set GetReportLines = coll
End Function

Private Sub ClosedCheckSummary(iReport As aclsReport, coll As Collection)
coll.Add Array("Closed Checks Summary", "", "", "", "")
PopFormatDict GetNewFormatObj(True, , xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)
Dim Checks As New Collection
Set Checks = Me.ReportType.GetClosedChecks(iReport)
Dim chk As Dictionary
For Each chk In Checks
    coll.Add Array(chk("CheckNumber"), "$" & CCur(chk("Total")), "", "", GetPayTypeValue(chk))
Next chk
PopFormatDict GetNewFormatObj(, , , , xlEdgeBottom), CStr(coll.Count)
End Sub

Private Sub OpenCheckSummary(iReport As aclsReport, coll As Collection)
coll.Add Array("Open Checks Summary", "", "", "", "")
PopFormatDict GetNewFormatObj(True, , xlCenterAcrossSelection, , xlEdgeBottom), CStr(coll.Count)
Dim Checks As New Collection
Set Checks = Me.ReportType.GetOpenChecks(iReport)
Dim chk As Dictionary
For Each chk In Checks
    coll.Add Array(chk("CheckNumber"), "$" & CCur(chk("Total")), "", "", GetPayTypeValue(chk))
Next chk
End Sub

Public Function GetStatPairs(iReport As aclsReport) As Dictionary
Dim dict As New Dictionary
Dim arr As Variant
Dim i As Integer
Dim BlankCount As Integer
BlankCount = 1
Dim BlankArr(1 To 2) As Variant
BlankArr(1) = "////////////////////////////////"
BlankArr(2) = ""
        
For i = 1 To UBound(DataOrder)
    If InValues(i, BlankRows) Then
        arr = BlankArr
        dict.Add "BlankRow" & BlankCount, BlankArr
        BlankCount = BlankCount + 1
    End If
    If Not InValues(i, BlankRows) Then
        arr = Me.ReportType.GetStatPair(CStr(DataOrder(i)), iReport)
        If Not IsEmpty(arr) Then
            dict.Add DataOrder(i), arr
        End If
    End If
Next i
Set GetStatPairs = dict
End Function

Public Function NullToZero(arr As Variant) As Variant
Dim i As Integer
For i = 1 To UBound(arr)
    If IsNull(arr(i)) Then arr(i) = CStr(0)
Next i
NullToZero = arr
End Function

Private Function GetPayTypeValue(chk As Dictionary) As String
Dim PayType As String
Dim dict As New Dictionary
dict.Add "$", chk("Cash")
dict.Add "Chg", chk("Charge")
dict.Add "GC", chk("GiftCert")
Dim key As Variant
For Each key In dict.Keys
    If dict(key) > 0 Then PayType = PayType & "/" & key
Next key
If chk("ChargeTip") > 0 Then PayType = PayType & "*"
If Not PayType = "" Then
    GetPayTypeValue = right(PayType, Len(PayType) - 1)
End If
Set dict = Nothing
End Function

Private Function GetBlankRows() As Variant
GetBlankRows = Array(7, 10, 15)
End Function

Private Function GetDataOrder() As Variant
Dim arr(1 To 19) As String
arr(1) = "SubTotal"
arr(2) = "Taxable"
arr(3) = "Tax"
arr(4) = "Charge"
arr(5) = "Cash"
arr(6) = "GiftCert"
arr(7) = "Blank1"
arr(8) = "Discount"
arr(9) = "Void"
arr(10) = "Blank2"
arr(11) = "FoodIn"
arr(12) = "Carryout"
arr(13) = "Beer"
arr(14) = "Wine"
arr(15) = "Blank3"
arr(16) = "Transactions"
arr(17) = "Avg"
arr(18) = "ChargeTip"
arr(19) = "CashDue"
GetDataOrder = arr
End Function
