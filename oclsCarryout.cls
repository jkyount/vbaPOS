VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oclsCarryout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public Function GetOrderName() As String
Dim OrderName As String
OrderName = InputBox("Name for this order?")
If Not OrderName = "" Or IsNull(OrderName) Then
    GetOrderName = OrderName
    Exit Function
End If
OrderName = ""
End Function

Public Function GetPickupTime() As String
Dim PickupTime As String
PickupTime = InputBox("Quoted pick-up time?")
If Not PickupTime = "" Or IsNull(PickupTime) Then
    GetPickupTime = PickupTime
    Exit Function
End If
PickupTime = ""
End Function

Public Function GetDineIn() As Boolean
GetDineIn = False
End Function

Public Function GetTable() As String
GetTable = "CARRYOUT"
End Function

Public Function GetPhone() As String
Dim Phone As String
Phone = InputBox("Phone number?")
If Not Phone = "" Or IsNull(Phone) Then
    GetPhone = Phone
    Exit Function
End If
Phone = ""
End Function

Public Sub CancelOrder()

End Sub

Public Sub SelectCheck()
SetOrderTypeIndicator "Carryout"
End Sub

Public Sub CloseCheck(check As String)
'5/20
If ThisOrder.TableVal = True Then
    UnassignCheck check
End If
End Sub

Public Sub CreateNewCheck(check As String)
'5/20
If ThisOrder.TableVal = True Then
    ThisTable.Assign check, ThisEmployee.ServerNum, ThisEmployee.FirstName
End If
End Sub

Public Function GetSplitOrderName() As String
'5/20
If ThisOrder.TableVal = True Then
    GetSplitOrderName = GetNextTable(ThisOrder.ValueDict("Table"))
    Exit Function
End If
GetSplitOrderName = ThisOrder.ValueDict("OrderName") & "/2"
End Function

Public Sub TransferCheck(check As String, TransferEmployee As zclsEmployee)

End Sub

Public Function GetTax(check As String) As Currency
'5/17/22
GetTax = pGetTax(check)
End Function

Private Function pGetTax(check As String) As Currency
'5/17/22
Dim DetailMatch As New zclsDailyCheckDetail
pGetTax = SumMatch(DetailMatch.Wrap(GetNewMatchObj(, check, "AlwaysTax", True)), "Price") * 0.08
Set DetailMatch = Nothing
End Function

Public Function GetFoodInTotal(check As String) As Currency
'5/17/22
GetFoodInTotal = 0
End Function

Public Function GetCarryoutTotal(check As String) As Currency
'5/17/22
Dim iDetail As New zclsDailyCheckDetail
GetCarryoutTotal = SumMatch(iDetail.Wrap(GetNewMatchObj(, check, "Category", "Food")), "Price")
Set iDetail = Nothing
End Function
