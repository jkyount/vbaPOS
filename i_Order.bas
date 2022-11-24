Attribute VB_Name = "i_Order"
Option Explicit

Public Sub SyncValues(Optional Order As aclsOrder)
If Order Is Nothing Then Set Order = ThisOrder
Order.ImportCheckDetails Order, currentcheck
Order.OrderType = Order.SameOrderType
End Sub

Public Function NewOrderObject(OrderType As Variant, check As String) As aclsOrder
Dim iOrder As New aclsOrder
Set NewOrderObject = iOrder.NewOrderObject(OrderType, check)
Set iOrder = Nothing
End Function

Public Function SetOrderInfo(Order As aclsOrder, Optional dict As Dictionary)
Dim iOrder As New aclsOrder
iOrder.SetOrderInfo Order, dict
Set iOrder = Nothing
End Function
