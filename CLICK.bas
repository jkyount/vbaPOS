Attribute VB_Name = "CLICK"
Option Explicit

Public Sub SHEET1_CategoryClick()
CategorySelect (Application.caller)
End Sub

Public Sub SHEET1_CustomItemClick()
Sheet1.CustomItemName.value = ""
Sheet1.CustomItemPrice.value = ""

HideShow "grpCustomItem"
End Sub
Public Sub SHEET1_SpecialInstructionClick()
Sheet1.SpecialInstructionText.Enabled = True
HideShape "grpguiMod"
HideShape "frmModConnector"
HideShow "grpSpecialInstruction"
Select Case Sheet1.Shapes("grpSpecialInstruction").Visible
    Case True
        Sheet1.SpecialInstructionText.Activate
    Case False
        Sheet1.SpecialInstructionText.Enabled = False
        Sheet1.SpecialInstructionText.Visible = False
End Select

Sheet1.SpecialInstructionText.value = ""
End Sub
Public Sub SHEET1_ModClick()
If GetCurrentFamily = "Mod" Then
    HideShape "grpguiMod"
    HideShape "frmModConnector"
    Sheet1.Shapes("Mod").line.ForeColor.RGB = rgbBlack
    CItem("1").OrderRank.CheckForRequiredComponents CItem("1")
    Exit Sub
End If
    
HideShape "grpSpecialInstruction"
'HideShow "grpguiMod"
Dim iFamily As New zclsFamily
iFamily.Family = "Mod"
ShowCategoryItems iFamily, GetCurrentPage
SetCurrentFamily "Mod"
HideShow "frmModConnector"
Set iFamily = Nothing
End Sub
Public Sub SHEET1_CancelItemClick()
ResetOrderState
ClearCollection CItem
End Sub
Public Sub SHEET1_Logout()
CancelOrder

ActivateLoginScreen
End Sub
Public Sub SHEET1_PrintGuestClick()
SyncValues
OrderSend currentcheck
PrintGuestCheck currentcheck
End Sub

Public Sub SHEET1_PaymentClick()
RestoreState
OrderSend currentcheck
ActivatePaymentScreen
Sheet8.PaymentAmount.value = ""
Sheet8.PaymentAmount.Activate
End Sub
Public Sub SHEET1_SendClick()
RestoreState
OrderSend currentcheck
ActivateHomeScreen
SCharge
End Sub
Public Sub SHEET1_PrintPrepClick()
SyncValues
OrderSend currentcheck
PrintPrepCheck currentcheck
'PrintPantryCheck currentcheck
End Sub
Public Sub SHEET1_SplitClick()
RestoreState
OrderSend currentcheck
ActivateSplitCheckScreen
End Sub

Public Sub SHEET1_CombineClick()
RestoreState
OrderSend currentcheck
ActivateCombineCheckScreen
End Sub
Public Sub SHEET1_DineInClick()
SetToDineIn currentcheck
UpdateTotal currentcheck
PopGuiWithCheckAttributes GetCheckAttributes(currentcheck), Sheet1
End Sub
Public Sub SHEET1_CarryoutClick()
SetToCarryout currentcheck
UpdateTotal currentcheck
PopGuiWithCheckAttributes GetCheckAttributes(currentcheck), Sheet1
End Sub
Public Sub SHEET1_WholeToppingClick()
WholeTopping
End Sub
Public Sub SHEET1_HalfToppingClick()
HalfTopping
End Sub
Public Sub SHEET1_EditSeatClick()
If Match(GetDailyCheckDetailObj.Wrap(GetNewMatchObj("CheckNumber", currentcheck))) = False Then Exit Sub
RestoreState
OrderSend currentcheck
ActivateEditSeatScreen
End Sub
Public Sub SHEET1_CancelClick()
CancelOrder
End Sub
Public Sub SHEET1_ScrollForwardMenuCategoryClick()
ScrollForward "MenuCategory"
End Sub

Public Sub SHEET1_ScrollBackwardMenuCategoryClick()
ScrollBackward "MenuCategory"
End Sub

Public Sub SHEET1_ScrollForwardComponentClick()
ScrollForward "Component"
End Sub

Public Sub SHEET1_ScrollBackwardComponentClick()
ScrollBackward "Component"
End Sub


Public Sub SHEET7_ViewTablesClick()
ShowShape "grpFloor"

IndicateTablesInUse
End Sub

Public Sub SHEET7_AccentColorClick()
HideShow "frmAccentColor"
End Sub

Public Sub SHEET7_TableClick()
Dim bname As String
bname = Application.caller
ActivateTable bname
End Sub

Public Sub SHEET7_BeginNewCheckClick()
NewDineIn
End Sub

Public Sub SHEET7_SwitchUserClick()
ActivateLoginScreen
End Sub

Public Sub SHEET7_CheckClick()
Dim bname As String
bname = Application.caller
Dim check As String
check = Sheet7.range("CheckSlot" & bname).value

SelectCheck check
End Sub
Public Sub SHEET7_OpenChecksClick()
DisplayOpenChecks
End Sub
Public Sub SHEET7_ConfigClick()
fConfig.Show
End Sub
Public Sub SHEET7_ServerReportClick()
If ThisEmployee.ServerNum = 0 Then
    RestoreState
End If
Dim iReport As New aclsReport
Dim iSales As New tclsSalesReport
Dim iServerReport As New rclsServerReport
iServerReport.DataSource = GetDailyDataObj
iSales.ReportType = iServerReport
iReport.ServerNum = ThisEmployee.ServerNum
iReport.RunReport iSales
'iServerReport.ServerReport ThisEmployee.ServerNum
End Sub

Public Sub SHEET7_ReprintClosedCheckClick()
fReprintClosedCheck.ShowForReprintCheck
End Sub
Public Sub SHEET7_TransferCheckClick()
fReprintClosedCheck.ShowForTransferCheck
End Sub


Public Sub SHEET8_DiscountsClick()
ShowShape "grpDiscountFrame"
ShowItemPage GetFamilyObj("Discounts"), 1
Sheet8.range("A1").Select
End Sub
Public Sub SHEET8_PaymentClick()
HideShape "grpDiscountFrame"
HideShape "grpguiDiscounts"
Sheet8.PaymentAmount.Activate
End Sub
Public Sub SHEET8_CashClick()
PayCash
End Sub
Public Sub SHEET8_ChargeClick()
PayCharge
End Sub
Public Sub SHEET8_GiftCertClick()
PayGiftCert
End Sub
Public Sub SHEET8_BackToOrderClick()
RecallCheck ThisOrder, Sheet1
ActivateOrderScreen
End Sub
Public Sub SHEET10_PrintDoneClick()
ExecuteSplit
End Sub
Public Sub SHEET10_PrintAndSplitAgainClick()
SplitAgain
End Sub
Public Sub SHEET10_CancelClick()
CancelSplit
End Sub
Public Sub SHEET13_CheckClick()
Dim bname As String
bname = Application.caller
Dim check As String
check = Sheet7.range("CheckSlot" & bname).value
SelectCheck check
End Sub
Public Sub SHEET14_DoneClick()
ExecuteSeatEdit
End Sub

