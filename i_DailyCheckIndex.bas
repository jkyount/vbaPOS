Attribute VB_Name = "i_DailyCheckIndex"
Option Explicit

Public Sub DailyCheckIndex_CalculateTotals(check As String) '<------- REVISED VERSION??
Dim Ord As New aclsOrder
Dim iDetail As New zclsDailyCheckDetail
Dim dict As New Dictionary
Dim iIndex As New zclsDailyCheckIndex
Ord.ImportCheckDetails Ord, check
Ord.OrderType = Ord.SameOrderType

dict.Add "SubTotal", SumMatch(iDetail.Wrap(GetNewMatchObj(, check)), "Price")
Debug.Print dict("SubTotal")
dict.Add "Beer", SumMatch(iDetail.Wrap(GetNewMatchObj(, check, "Category", "Beer")), "Price")
dict.Add "Wine", SumMatch(iDetail.Wrap(GetNewMatchObj(, check, "Category", "Wine")), "Price")
dict.Add "Discount", SumMatch(iDetail.Wrap(GetNewMatchObj(, check, "Family", "Discounts")), "Price")
dict.Add "FoodIn", Ord.OrderType.GetFoodInTotal(check)
dict.Add "Carryout", Ord.OrderType.GetCarryoutTotal(check)
UpdateFromDict iIndex.Wrap(GetNewMatchObj(, check)), dict
UpdateTax check
UpdateTotal check
Set iIndex = Nothing
Set iDetail = Nothing
Set Ord = Nothing
Set dict = Nothing
End Sub

Public Sub UpdateTax(check As String)
Dim iDailyCheckIndex As New zclsDailyCheckIndex
iDailyCheckIndex.UpdateTax (check)
Set iDailyCheckIndex = Nothing
End Sub

Public Sub UpdateTotal(check As String)
Dim iDailyCheckIndex As New zclsDailyCheckIndex
iDailyCheckIndex.UpdateTotal (check)
Set iDailyCheckIndex = Nothing
End Sub

Public Sub UpdateDailyCheckIndex(check As String, Params As Dictionary)
Dim iDailyCheckIndex As New zclsDailyCheckIndex
iDailyCheckIndex.UpdateDailyCheckIndex check, Params
Set iDailyCheckIndex = Nothing
End Sub

Public Function GetChecks(Optional Where As String = "", Optional Equals As Variant, Optional AndWhere As String = "", Optional Equals2 As Variant) As Variant
Dim iDataObject As New aclsDataObject
Dim iDailyCheckIndex As New zclsDailyCheckIndex
Set iDataObject = iDailyCheckIndex.Wrap(GetNewMatchObj(Where, Equals, AndWhere, Equals2))
GetChecks = iDailyCheckIndex.FormatRecordset(GetRecordsetMatch(iDataObject, ConstructMatchQuery(iDataObject)))
End Function

Public Function GetCheckAttributes(check As String) As Dictionary
Dim iIndex As New zclsDailyCheckIndex
Set GetCheckAttributes = GetValueDict(iIndex.Wrap(GetNewMatchObj("CheckNumber", check)))(1)
Set iIndex = Nothing
End Function

Public Sub DailyCheckIndex_CloseOrder(check As String)
Dim iDailyCheckIndex As New zclsDailyCheckIndex
iDailyCheckIndex.CloseOrder check
Set iDailyCheckIndex = Nothing
End Sub




