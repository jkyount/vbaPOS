Attribute VB_Name = "Functions_Home"
Option Explicit

Public Sub ActivateHomeScreen()
Sheet7.Activate
SHEET7_ViewTablesClick
Sheet7.Shapes("ServerName").TextFrame.Characters.text = ThisEmployee.FirstName
Set ThisOrder = Nothing
End Sub

Public Sub ChangeAccentColor()
Dim color As XlRgbColor
Dim bname As String
bname = Application.caller
color = Sheet7.Shapes(bname).Fill.ForeColor.RGB
ThisEmployee.SetAccentColor color
IndicateTablesInUse
HideShape "frmAccentColor"
End Sub


Public Sub ActivateTable(TableName As String)
Application.ScreenUpdating = False
ThisTable.ParentTable = TableName

If ThisTable.Checks.Count < 1 Then
    NewDineIn
    Application.ScreenUpdating = True
Exit Sub
End If
If ThisTable.Checks.Count > 1 Then
    DisplayChecks GetChecks("Table", ThisTable.ParentTable, "Closed", False)
    
    ShowShape "BeginNewCheck"
    Application.ScreenUpdating = True
    Exit Sub
End If

If Match(ThisTable.Wrap(GetNewMatchObj("CheckNumber", ThisTable.Checks(1), "ServerNum", ThisEmployee.ServerNum))) = True Then
    SelectCheck ThisTable.Checks(1)
    Application.ScreenUpdating = True
    Exit Sub
End If

DisplayChecks GetChecks("Table", ThisTable.ParentTable, "Closed", False)
ShowShape "BeginNewCheck"
Application.ScreenUpdating = True
End Sub

Public Sub SelectCheck(check As String)
Dim iIndex As New zclsDailyCheckIndex
If Match(iIndex.Wrap(GetNewMatchObj(, check, "ServerNum", ThisEmployee.ServerNum))) = False Then
    MsgBox "Nacho"
    Exit Sub
End If
If ValueMatch(iIndex.Wrap(GetNewMatchObj(, check)), "Closed") = True Then
    UnassignCheck check
    IndicateTablesInUse
    Exit Sub
End If
Application.ScreenUpdating = False
    currentcheck = check
    ThisOrder.check = check
    SyncValues
    
    RecallCheck ThisOrder, Sheet1
    ActivateOrderScreen
    ThisOrder.OrderType.SelectCheck
Application.ScreenUpdating = True
End Sub

Public Sub NewDineIn()
Dim NewCheckNumber As String
NewCheckNumber = GetNextCheck
currentcheck = NewCheckNumber

InitializeCheck NewCheckNumber
ThisTable.Assign NewCheckNumber, ThisEmployee.ServerNum, ThisEmployee.FirstName
Set ThisOrder = NewOrderObject(GetNewDineInObj, NewCheckNumber)
SetOrderInfo ThisOrder
'SetOrderType NewCheckNumber, GetNewDineInObj
UpdateDailyCheckIndex ThisOrder.check, ThisOrder.OrderDetails

SyncValues
Application.ScreenUpdating = False
   
    RecallCheck ThisOrder, Sheet1
    
    ActivateOrderScreen
    SetOrderTypeIndicator "DineIn"
Application.ScreenUpdating = True

End Sub

Public Sub NewCarryout()
Dim NewCheckNumber As String
NewCheckNumber = GetNextCheck
currentcheck = NewCheckNumber


InitializeCheck NewCheckNumber

'SetOrderType NewCheckNumber, GetNewCarryoutObj
Set ThisOrder = NewOrderObject(GetNewCarryoutObj, NewCheckNumber)
SetOrderInfo ThisOrder
UpdateDailyCheckIndex ThisOrder.check, ThisOrder.OrderDetails
SyncValues
Application.ScreenUpdating = False
    
    RecallCheck ThisOrder, Sheet1
    
    ActivateOrderScreen
    SetOrderTypeIndicator "Carryout"
Application.ScreenUpdating = True

End Sub



Public Sub RecallCheck(Order As aclsOrder, sheet As Worksheet)
Set collCheckData = RecallCheckLines(Order.check)
WriteCheckLines sheet.range("CheckRange"), collCheckData
PopGuiWithCheckAttributes Order.ValueDict, sheet
End Sub


Public Sub CheckTransfer(check As String, NewServerNum As Integer)

Dim iOrder As New aclsOrder
iOrder.check = check
iOrder.ImportCheckDetails iOrder, check
iOrder.OrderType = iOrder.SameOrderType

Dim TransferEmployee As New zclsEmployee
TransferEmployee.IDNumber = ValueMatch(ThisEmployee.Wrap(GetNewMatchObj("ServerNum", NewServerNum)), "IDNumber")
'TransferEmployee.ServerNum = NewServerNum
iOrder.TransferCheck check, TransferEmployee

End Sub




