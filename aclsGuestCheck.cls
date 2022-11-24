VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsGuestCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCheck As String
Private pPrintRange As range
Private pTopOfCheck As range
Private pPrintData As Variant

Public Property Let check(value As String)
pCheck = value
End Property
Public Property Get check() As String
check = pCheck
End Property
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


Public Function NewPrintObject(check As String) As aclsGuestCheck
Dim x As New aclsGuestCheck
x.PrintParams check
Set NewPrintObject = x
End Function

Public Sub PrintParams(check As String)
Me.check = check
Me.PrintRange = Sheet13.range("GuestCheckRange")
Me.TopOfCheck = Sheet13.range("TopOfCheck_Guest")
Me.PrintData = GetPrintData(check)
End Sub
Public Function GetPrintArea(obj As Variant) As range
Dim Rg1 As range
Dim rg2 As range

Set Rg1 = obj.TopOfCheck
Set rg2 = Rg1.Offset(200, 0).End(xlUp).Offset(0, 1)
Set GetPrintArea = Sheet13.range(Rg1, rg2)
End Function

Public Sub FormatPrintRange()

End Sub

Private Function GetPrintData(check As String) As Variant
Dim x As New zclsDailyCheckDetail
GetPrintData = FilteredMatch(x.Wrap(GetNewMatchObj(, check)), "ItemName", "Price", "Seat", "LocalGroup")
End Function

Public Sub InitializePrintRange(obj As aclsGuestCheck)
obj.PrintRange.value = ""
obj.PrintRange.Columns(1).HorizontalAlignment = xlLeft
obj.PrintRange.Columns(2).HorizontalAlignment = xlCenter
End Sub

Public Sub WritePrintLines(coll As Collection)
Dim i As Integer, SeatLines As Integer, seat As Integer, FirstPrintRow As Integer
Dim range As range
Set range = Me.PrintRange
Dim SeatIndicator As String

If coll.Count = 0 Then
    range.value = ""
    MsgBox "Attempted to write an empty collection."
    Exit Sub
End If
FirstPrintRow = 1
Dim iOrder As New aclsOrder
iOrder.ImportCheckDetails iOrder, Me.check
If iOrder.ValueDict("DineIn") = False Then
    FirstPrintRow = 2
    range.Rows(1).Cells(, 1).value = "[ : : : . .  C A R R Y O U T  . . : : : ]"
    range.Rows(1).HorizontalAlignment = xlCenterAcrossSelection
End If
Set iOrder = Nothing
Dim arr As Variant
SeatLines = 1
seat = coll("Line1").seat
SeatIndicator = "- - - - - - - - - - - Seat " & seat & "- - - - - - - - - - -"
range.Rows(FirstPrintRow).Cells(, 1).value = SeatIndicator
For i = FirstPrintRow To coll.Count
    If Not coll("Line" & i).seat = seat Then
        seat = coll("Line" & i).seat
        SeatIndicator = "- - - - - - - - - - - Seat " & seat & "- - - - - - - - - - -"
        range.Rows(i + SeatLines).Cells(, 1).value = SeatIndicator
        SeatLines = SeatLines + 1
    End If
    arr = coll("Line" & i).Data
    range.Rows(i + SeatLines).Cells(, 1).value = ">  " & arr(1)(0, 0)
    If coll("Line" & i).primary = False Then
        range.Rows(i + SeatLines).Cells(, 1).value = "        " & arr(1)(0, 0)
    End If
    range.Rows(i + SeatLines).Cells(, 2).value = arr(1)(1, 0)
    
Next i
End Sub

Public Sub AddCheckFooter(check As String)
Dim DetailMatch As New zclsDailyCheckDetail
Dim Ord As New aclsOrder
Ord.ImportCheckDetails Ord, check
Ord.OrderType = Ord.SameOrderType
Dim subtotal As Currency
Dim Tax As Currency
Dim ServiceCharge As Currency
subtotal = SumMatch(DetailMatch.Wrap(GetNewMatchObj(, check)), "Price")
ServiceCharge = Ord.ValueDict("ServiceCharge")
Tax = Ord.OrderType.GetTax(check)

Dim Rg1 As range
Set Rg1 = Sheet13.range("TopOfCheck_Guest").Offset(200, 0).End(xlUp).Offset(2, 0)
Dim rg2 As range
Set rg2 = Sheet13.range("TopOfCheck_Guest").Offset(200, 0).End(xlUp).Offset(22, 1)
Dim FooterRg As range
Set FooterRg = range(Rg1, rg2)
With FooterRg
    .Rows(1).HorizontalAlignment = xlRight
    .Rows(1).Cells(, 1).value = "-------------------"
    
    .Rows(2).HorizontalAlignment = xlRight
    .Rows(2).Cells(, 1).value = "Subtotal:"
    .Rows(2).Cells(, 2).value = subtotal
    
    .Rows(3).HorizontalAlignment = xlRight
    .Rows(3).Cells(, 1).value = "Tax:"
    .Rows(3).Cells(, 2).value = Tax
    
    Dim k As Integer
    k = 3
    If Not ServiceCharge = 0 Then
    k = k + 1
    .Rows(k).HorizontalAlignment = xlRight
    .Rows(k).Cells(, 1).value = "Service Charge:"
    .Rows(k).Cells(, 2).value = ServiceCharge
    End If
    k = k + 1
    .Rows(k).HorizontalAlignment = xlRight
    .Rows(k).Cells(, 1).value = "Cash Total:"
    .Rows(k).Cells(, 2).value = Tax + subtotal
    k = k + 1
    .Rows(k).HorizontalAlignment = xlRight
    .Rows(k).Cells(, 1).value = "Charge Total:"
    .Rows(k).Cells(, 2).value = (Tax + subtotal) + ((Tax + subtotal) * 0.035)
    
    
    
''    .Offset(-1, 0).HorizontalAlignment = xlRight
''    .Offset(-1, 0).Value = "-------------------"
''    .HorizontalAlignment = xlRight
''    .Value = "Subtotal:"
''    .Offset(0, 1).Value = subtotal
''    .Offset(1, 0).HorizontalAlignment = xlRight
''    .Offset(1, 0).Value = "Tax:"
''    .Offset(1, 1).Value = Tax
''    .Offset(2, 0).HorizontalAlignment = xlRight
''    .Offset(2, 0).Value = "Total:"
''    .Offset(2, 1).Value = Tax + subtotal


    Dim cash As Currency
    Dim charge As Currency
    Dim giftcert As Currency
    cash = Ord.ValueDict("Cash")
    charge = Ord.ValueDict("Charge")
    giftcert = Ord.ValueDict("GiftCert")

    
    k = k + 2


    If Not Ord.Payments = 0 Then
       
    .Rows(k).HorizontalAlignment = xlRight
    .Rows(k).Cells(, 1).value = "----------------------------"
    k = k + 1
'        .Offset(3, 0).HorizontalAlignment = xlCenter
'        .Offset(3, 0).Value = "----------------------------"

         If Not cash = 0 Then
            .Rows(k).HorizontalAlignment = xlRight
            .Rows(k).Cells(, 1).value = "Cash paid:"
            .Rows(k).Cells(, 2).value = "(" & cash & ")"
            k = k + 1
'            .Offset(k, 0).HorizontalAlignment = xlRight
'            .Offset(k, 0).Value = "Cash paid:"
'            .Offset(k, 1).Value = "(" & cash & ")"
'            k = k + 1
        End If
        If Not charge = 0 Then
            .Rows(k).HorizontalAlignment = xlRight
            .Rows(k).Cells(, 1).value = "Charge paid:"
            .Rows(k).Cells(, 2).value = "(" & charge & ")"
            k = k + 1
        End If
'        If Not charge = 0 Then
'            .Offset(k, 0).HorizontalAlignment = xlRight
'            .Offset(k, 0).Value = "Charge paid:"
'            .Offset(k, 1).Value = "(" & charge & ")"
'            k = k + 1
'        End If
        If Not giftcert = 0 Then
            .Rows(k).HorizontalAlignment = xlRight
            .Rows(k).Cells(, 1).value = "Gift Cert paid:"
            .Rows(k).Cells(, 2).value = "(" & giftcert & ")"
            k = k + 1
        End If
    
            .Rows(k).HorizontalAlignment = xlRight
            .Rows(k).Cells(, 1).value = "Amount due:"
            .Rows(k).Cells(, 2).value = Ord.ValueDict("Total") - Ord.Payments
            k = k + 2
'        If Not giftcert = 0 Then
'            .Offset(k, 0).HorizontalAlignment = xlRight
'            .Offset(k, 0).Value = "Gift Cert paid:"
'            .Offset(k, 1).Value = "(" & giftcert & ")"
'            k = k + 1
'        End If
'        .Offset(k, 0).HorizontalAlignment = xlRight
'        .Offset(k, 0).Value = "Amount due:"
'        .Offset(k, 1).Value = Total - Payments
'        k = k + 1
'    End If
    End If
    .Rows(k).Cells(, 1).value = "     Thank you!  Visit us again soon."
    .Rows(k).HorizontalAlignment = xlLeft
    .Rows(k + 1).Cells(, 1).value = "        " & Now
    .Rows(k + 1).HorizontalAlignment = xlLeft
End With
Set FooterRg = Nothing
End Sub

Public Function GetPrintColl(obj As Variant) As Collection
Dim check As String
check = obj.check
Dim coll As New Collection
Set coll = SortCheckLines(RecallCheckLines(check))

Dim line As zclsCheckLines
Dim iDetail As New zclsDailyCheckDetail
For Each line In coll
    line.Data = FilteredMatch(iDetail.Wrap(GetNewMatchObj("CheckNumber", check, "LocalGroup", line.LocalGroup)), "ItemName", "Price")
Next line

Dim PrintColl As New Collection
Dim arr As Variant
Dim row As Integer
row = 1
For Each line In coll
    arr = line.Data
    If Not arr(1)(1, 0) = 0 Then
        PrintColl.Add line, "Line" & row
        row = row + 1
    End If
Next line
Set GetPrintColl = PrintColl
Set iDetail = Nothing
Set coll = Nothing
Set PrintColl = Nothing
Set line = Nothing


End Function
