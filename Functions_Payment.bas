Attribute VB_Name = "Functions_Payment"


Option Explicit
Public Sub ActivatePaymentScreen()
Sheet8.Activate
SyncValues
UpdateGUI
InitializeRanges
InitializeShapes
End Sub
Private Sub InitializeRanges()

Set collCheckData = RecallCheckLines(currentcheck)
Sheet8.range("CheckRange").value = ""
Sheet8.PaymentAmount.value = 0
WriteCheckLines Sheet8.range("CheckRange"), collCheckData
End Sub
Private Sub InitializeShapes()
HideShape "grpDiscountFrame"
HideShape "grpguiDiscounts"
Sheet8.PaymentAmount.Activate
End Sub
Public Sub PayCash()
RestoreState
Dim PaymentAmount As Currency, Total As Currency, AdjustedTotal As Currency
PaymentAmount = Sheet8.PaymentAmount.value
Total = ThisOrder.ValueDict("Total")
AdjustedTotal = Total - ThisOrder.ValueDict("Cash") - ThisOrder.ValueDict("Charge") - ThisOrder.ValueDict("GiftCert")
If PaymentAmount >= AdjustedTotal Then
    'Sheet8.Shapes("ChangeDue").TextFrame.Characters.text = (ThisCheckTotal - CashAmount)
    PaymentAmount = AdjustedTotal
    UpdatePayments currentcheck, "Cash", PaymentAmount
    CloseCheck
    Exit Sub
End If
UpdatePayments currentcheck, "Cash", PaymentAmount
SyncValues
UpdateGUI
MsgBox ("Partial payment applied.")
Sheet8.PaymentAmount.value = 0

End Sub

Private Sub UpdateGUI()
PopGuiWithCheckAttributes ThisOrder.ValueDict, Sheet8
Sheet8.range("Payments").value = ThisOrder.Payments
Sheet8.range("AdjustedCheckTotal").value = ThisOrder.AdjustedTotal
End Sub

Public Sub PayCharge()
RestoreState
Dim PaymentAmount As Currency, Total As Currency, AdjustedTotal As Currency, TipAmount As Currency
PaymentAmount = Sheet8.PaymentAmount.value
Total = ThisOrder.ValueDict("Total")
AdjustedTotal = Total - ThisOrder.ValueDict("Cash") - ThisOrder.ValueDict("Charge") - ThisOrder.ValueDict("GiftCert")
TipAmount = CCur(PaymentAmount - AdjustedTotal)
Debug.Print CCur(TipAmount)
If PaymentAmount >= AdjustedTotal Then
    'Sheet8.Shapes("ChangeDue").TextFrame.Characters.text = (ThisCheckTotal - CashAmount)
    If MsgBox("Tip amount is " & TipAmount & ".  Is that correct?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    UpdatePayments currentcheck, "Charge", AdjustedTotal
    UpdateTips currentcheck, TipAmount
    CloseCheck
    Exit Sub
End If
UpdatePayments currentcheck, "Charge", PaymentAmount
SyncValues
UpdateGUI
MsgBox ("Partial payment applied.")
Sheet8.PaymentAmount.value = 0

End Sub

Public Sub PayGiftCert()
RestoreState
Dim PaymentAmount As Currency, Total As Currency, AdjustedTotal As Currency
PaymentAmount = Sheet8.PaymentAmount.value
Total = ThisOrder.ValueDict("Total")
AdjustedTotal = Total - ThisOrder.ValueDict("Cash") - ThisOrder.ValueDict("Charge") - ThisOrder.ValueDict("GiftCert")
If PaymentAmount >= AdjustedTotal Then
    'Sheet8.Shapes("ChangeDue").TextFrame.Characters.text = (ThisCheckTotal - CashAmount)
    PaymentAmount = AdjustedTotal
    UpdatePayments currentcheck, "GiftCert", PaymentAmount
    CloseCheck
    Exit Sub
End If
UpdatePayments currentcheck, "GiftCert", PaymentAmount
SyncValues
UpdateGUI
MsgBox ("Partial payment applied.")
Sheet8.PaymentAmount.value = 0

End Sub

Public Sub UpdatePayments(check As String, PaymentType As String, Amount As Currency)
Dim iPayment As New zclsDailyCheckIndex
AddToMatch iPayment.Wrap(GetNewUpdateObj(, check, PaymentType, Amount))
End Sub
Public Sub UpdateTips(check As String, TipAmount As Currency)
Dim iPayment As New zclsDailyCheckIndex
AddToMatch iPayment.Wrap(GetNewUpdateObj(, check, "ChargeTip", TipAmount))
End Sub

Public Sub Discount(Optional Price As Variant)
Application.ScreenUpdating = False
Dim ItemID As Integer
ItemID = CInt(Application.caller)
Dim ItemType As Variant
Set ItemType = GetItemType(CStr(GetItemClassCode(ItemID)))
ItemType.ItemInitialize ItemID


Set ThisItem = New aclsItem
ThisItem.ItemType = ItemType
Set ThisItem = PopItem(ItemType.ItemID, ThisItem)
ThisItem.Price = -(ThisOrder.ValueDict("SubTotal") * ThisItem.DiscountPrice) + ThisItem.Price

If -(ThisItem.Price) > ThisOrder.ValueDict("SubTotal") Then
    MsgBox "Cannot apply discount greater than this order's subtotal.  Please select a different discount amount."
    Set ThisItem = Nothing
    Set ItemType = Nothing
    Exit Sub
End If
AddDiscount

OrderSend currentcheck
ActivatePaymentScreen
Set ThisItem = Nothing
Set ItemType = Nothing
Application.ScreenUpdating = True
End Sub

Public Sub CustomDiscount()
If Sheet8.inbxDiscountAmount <= 0 Or Not IsNumeric(Sheet8.inbxDiscountAmount) Then
    MsgBox "Discount amount must be a number greater than 0."
    Exit Sub
End If
Application.ScreenUpdating = False
Dim ItemID As Integer
ItemID = CInt(Application.caller)
Dim ItemType As Variant
Set ItemType = GetItemType(CStr(GetItemClassCode(ItemID)))
ItemType.ItemInitialize ItemID

Set ThisItem = New aclsItem
ThisItem.ItemType = ItemType
Set ThisItem = PopItem(ItemType.ItemID, ThisItem)
ThisItem.Price = -(Sheet8.inbxDiscountAmount.value)
If -(ThisItem.Price) > ThisOrder.ValueDict("SubTotal") Then
    MsgBox "Cannot apply discount greater than this order's subtotal.  Please select a different discount amount."
    Set ThisItem = Nothing
    Set ItemType = Nothing
    Exit Sub
End If
AddDiscount

OrderSend currentcheck
ActivatePaymentScreen
Set ThisItem = Nothing
Application.ScreenUpdating = True
End Sub

Public Sub AddDiscount()

Dim iDetail As New zclsDailyCheckDetail
iDetail.AddItem ThisItem, GetNextLocalGroup(currentcheck), GetNextEntityGroup(currentcheck)
Set iDetail = Nothing

End Sub

Public Sub AddServiceCharge()
Dim iIndex As New zclsDailyCheckIndex
Update iIndex.Wrap(GetNewUpdateObj("CheckNumber", currentcheck, "ServiceCharge", (ThisOrder.ValueDict("SubTotal") * 0.2)))
Set iIndex = Nothing
UpdateTotal currentcheck
End Sub





