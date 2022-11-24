VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oclsDineIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public Function GetOrderName() As String
GetOrderName = ThisTable.Table
End Function

Public Function GetPickupTime() As String
GetPickupTime = "---"
End Function

Public Function GetDineIn() As Boolean
GetDineIn = True
End Function

Public Function GetTable() As String
GetTable = ThisTable.ParentTable
End Function

Public Function GetPhone() As String
GetPhone = ""
End Function

Public Sub CancelOrder()
ThisTable.UnassignCheck currentcheck
End Sub

Public Sub SelectCheck()
RecallTableState currentcheck
SetOrderTypeIndicator "DineIn"
End Sub

Public Sub CloseCheck(check As String)
'5/20
UnassignCheck check
End Sub

Public Sub CreateNewCheck(NewCheck As String)


ThisTable.Assign NewCheck, ThisEmployee.ServerNum, ThisEmployee.FirstName
End Sub

Public Function GetSplitOrderName() As String

GetSplitOrderName = GetNextTable(ThisOrder.ValueDict("Table"))
End Function

Public Sub TransferCheck(check As String, TransferEmployee As zclsEmployee)
Dim iTable As New zclsTable
Update iTable.Wrap(GetNewUpdateObj("CheckNumber", check, "ServerNum", TransferEmployee.ServerNum))
Update iTable.Wrap(GetNewUpdateObj("CheckNumber", check, "ServerName", TransferEmployee.FirstName))
Set iTable = Nothing
End Sub

Public Function GetTax(check As String) As Currency
'5/17/22
GetTax = pGetTax(check)
End Function

Private Function pGetTax(check As String) As Currency
'5/17/22
Dim iDetail As New zclsDailyCheckDetail
pGetTax = ((SumMatch(iDetail.Wrap(GetNewMatchObj(, check)), "Price")) * 0.08)
Set iDetail = Nothing
End Function

Public Function GetFoodInTotal(check As String) As Currency
'5/17/22
Dim iDetail As New zclsDailyCheckDetail
GetFoodInTotal = SumMatch(iDetail.Wrap(GetNewMatchObj(, check, "Category", "Food")), "Price")
Set iDetail = Nothing
End Function

Public Function GetCarryoutTotal(check As String) As Currency
'5/17/22
GetCarryoutTotal = 0
End Function





