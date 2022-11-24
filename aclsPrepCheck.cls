VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsPrepCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCheck As String
Private pPrintRange As range
Private pTopOfCheck As range
Private pPrintData As Variant
Private pPrintType As Variant
Const RowsFromTop As Integer = 3

Public Property Let check(value As String)
pCheck = value
End Property
Public Property Get check() As String
check = pCheck
End Property

'Public Property Get PrintArea() As Range
'Set PrintArea = GetPrintArea
'End Property
Public Property Let PrintRange(value As range)
Set pPrintRange = value
End Property
Public Property Get PrintRange() As range
Set PrintRange = pPrintRange
End Property
Public Property Let TopOfCheck(value As range)
Set pTopOfCheck = value
End Property
Public Property Get TopOfCheck() As range
Set TopOfCheck = pTopOfCheck
End Property
Public Property Let PrintData(value As Variant)
pPrintData = value
End Property
Public Property Get PrintData() As Variant
PrintData = pPrintData
End Property

Public Property Let PrintType(value As Variant)
Set pPrintType = value
End Property
Public Property Get PrintType() As Variant
Set PrintType = pPrintType
End Property



Public Function NewPrintObject(check As String, PrintType As Variant) As aclsPrepCheck
Dim x As New aclsPrepCheck
x.PrintParams check
x.PrintType = PrintType
Set NewPrintObject = x
End Function

Public Sub PrintParams(check As String)
Me.check = check
Me.PrintRange = Sheet13.range("PrepCheckRange")
Me.TopOfCheck = Sheet13.range("TopOfCheck_Prep")
End Sub

Public Function GetPrintArea(obj As Variant) As range
Dim Rg1 As range
Dim rg2 As range

Set Rg1 = obj.TopOfCheck
Set rg2 = obj.TopOfCheck.Offset(200, 0).Offset(0, 1).End(xlUp)
Set GetPrintArea = Sheet13.range(Rg1, rg2)

End Function
Public Sub FormatPrintRange()
Dim PrintRange As range
Set PrintRange = Me.PrintRange
With PrintRange
    .Font.name = "Lucida Sans Unicode"
    .Font.Size = 18
    .Font.Bold = True
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeLeft).weight = xlThin
    .Borders(xlEdgeRight).weight = xlThin
End With

End Sub


Public Sub InitializePrintRange(obj As Variant)
obj.PrintRange.Clear
obj.PrintRange.Columns(1).HorizontalAlignment = xlLeft
obj.PrintRange.Columns(2).HorizontalAlignment = xlLeft
End Sub
Public Sub WritePrintLines(coll As Collection)

AddCarryoutHeader Me

Dim range As range
Set range = Me.PrintRange
Dim row As Integer
row = RowsFromTop

Me.PrintType.WritePrintLines range, coll, row
AddTimeStamp range, row
End Sub

Private Sub AddCarryoutHeader(obj As aclsPrepCheck)
Dim iOrder As New aclsOrder
iOrder.ImportCheckDetails iOrder, Me.check
If iOrder.ValueDict("DineIn") = False Then
    Me.PrintRange.Rows(1).Cells(, 1).value = "[ : : .  C A R R Y O U T  . : : ]"
    Me.PrintRange.Rows(1).HorizontalAlignment = xlCenterAcrossSelection
End If
Set iOrder = Nothing
End Sub

Private Sub AddTimeStamp(range As range, row As Integer)
range.Rows(row).Cells(, 2).value = "----------------------"
range.Rows(row + 1).Cells(, 1).value = Format(Now, "hh:mm")
range.Rows(row + 2).Cells(, 2).value = "----------------------"
range.Rows(row).HorizontalAlignment = xlCenterAcrossSelection
range.Rows(row + 1).HorizontalAlignment = xlCenterAcrossSelection
range.Rows(row + 2).HorizontalAlignment = xlCenterAcrossSelection
End Sub

Public Sub AddCheckFooter(check As String)

End Sub



Public Function GetPrintColl(obj As Variant) As Collection
Set GetPrintColl = obj.PrintType.GetPrintColl(obj.check)
End Function


