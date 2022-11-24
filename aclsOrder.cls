VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pCheck As String
Private pValueDict As Dictionary
Private pOrderType As Variant
Private pOrderDetails As Dictionary
Private pNewOrderDetails As Dictionary
Private pTableVal As Boolean


Public Property Let check(value As String)
pCheck = value
End Property
Public Property Get check() As String
check = pCheck
End Property

Public Property Let ValueDict(value As Dictionary)
Set pValueDict = value
End Property
Public Property Get ValueDict() As Dictionary
Set ValueDict = pValueDict
End Property

Public Property Let OrderDetails(value As Dictionary)
Set pOrderDetails = value
End Property
Public Property Get OrderDetails() As Dictionary
Set OrderDetails = pOrderDetails
End Property

Public Property Let NewOrderDetails(value As Dictionary)
Set pNewOrderDetails = value
End Property
Public Property Get NewOrderDetails() As Dictionary
Set NewOrderDetails = NewOrderInfo(Me)
End Property

Public Property Let OrderType(value As Variant)
Set pOrderType = value
End Property
Public Property Get OrderType() As Variant
Set OrderType = pOrderType
End Property


Public Property Get TableVal() As Boolean
TableVal = GetTableVal
End Property

Public Property Get Payments() As Currency
Payments = Me.ValueDict("Cash") + Me.ValueDict("Charge") + Me.ValueDict("GiftCert")
End Property
Public Property Get AdjustedTotal() As Currency
AdjustedTotal = Me.ValueDict("Total") - (Me.ValueDict("Cash") + Me.ValueDict("Charge") + Me.ValueDict("GiftCert"))
End Property


Public Function NewOrderObject(OrderType As Variant, check As String) As aclsOrder
Dim iOrder As New aclsOrder
iOrder.OrderType = OrderType
iOrder.check = check
iOrder.ImportCheckDetails iOrder, check
Set NewOrderObject = iOrder
End Function

Public Function SetOrderInfo(Order As aclsOrder, Optional dict As Dictionary)
If dict Is Nothing Then
    Order.OrderDetails = NewOrderInfo(Order)
End If
If Not dict Is Nothing Then
    Order.OrderDetails = dict
End If
End Function

Public Function NewOrderInfo(Order As aclsOrder) As Dictionary
Dim dict As New Dictionary
dict.Add "OrderName", Order.OrderType.GetOrderName
dict.Add "PickupTime", Order.OrderType.GetPickupTime
dict.Add "DineIn", Order.OrderType.GetDineIn
dict.Add "Table", Order.OrderType.GetTable
dict.Add "Phone", Order.OrderType.GetPhone
dict.Add "ServerNum", ThisEmployee.ServerNum
dict.Add "ServerName", ThisEmployee.FirstName
dict.Add "CheckDate", Format(Now, "mmddyy")
dict.Add "CheckOpen", Format(Now, "hh:mm")
Set NewOrderInfo = dict
Set dict = Nothing
End Function

Public Function SameOrderInfo() As Dictionary
Set SameOrderInfo = Me.ValueDict
End Function

Public Function SameOrderType() As Variant
If Me.ValueDict("DineIn") = True Then
    Set SameOrderType = GetNewDineInObj
End If
If Me.ValueDict("DineIn") = False Then
Debug.Print Me.ValueDict("DineIn")
    Set SameOrderType = GetNewCarryoutObj
End If
End Function

Public Sub ImportCheckDetails(OrderObj As aclsOrder, check As String)
OrderObj.ValueDict = GetCheckAttributes(check)
End Sub

Public Function SplitCheckInitialOrderDetails() As Dictionary
Dim dict As New Dictionary
dict.Add "OrderName", Me.ValueDict("OrderName")
dict.Add "PickupTime", Me.ValueDict("PickupTime")
dict.Add "DineIn", Me.ValueDict("DineIn")
dict.Add "Table", Me.ValueDict("Table")
dict.Add "Phone", Me.ValueDict("Phone")
dict.Add "ServerNum", Me.ValueDict("ServerNum")
dict.Add "ServerName", Me.ValueDict("ServerName")
dict.Add "CheckDate", Me.ValueDict("CheckDate")
dict.Add "CheckOpen", Me.ValueDict("CheckOpen")
Set SplitCheckInitialOrderDetails = dict
End Function

Public Function CreateNewOrder(OrderType As Variant, check As String) As aclsOrder
Dim iOrder As New aclsOrder
iOrder.OrderType = OrderType
InitializeOrder check
iOrder.check = check
iOrder.ImportCheckDetails iOrder, check
Set CreateNewOrder = iOrder
End Function

Public Sub InitializeOrder(check As String)
InitializeCheck check
Me.OrderType.Initialize check
End Sub

Public Sub TransferCheck(check As String, TransferEmployee As zclsEmployee)
ReplaceDictValue Me.ValueDict, "ServerNum", TransferEmployee.ServerNum
ReplaceDictValue Me.ValueDict, "ServerName", TransferEmployee.FirstName
UpdateDailyCheckIndex check, Me.ValueDict
Me.OrderType.TransferCheck check, TransferEmployee
End Sub

Private Function GetTableVal() As Boolean
Dim iTable As New zclsTable
If Match(iTable.Wrap(GetNewMatchObj("CheckNumber", Me.ValueDict("CheckNumber")))) = True Then
    GetTableVal = True
End If
End Function


