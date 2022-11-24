Attribute VB_Name = "i_Menu"
Option Explicit

Public Function GetMenuObj() As zclsMenu
Dim iMenu As zclsMenu
Set iMenu = New zclsMenu
Set GetMenuObj = iMenu
Set iMenu = Nothing
End Function
Public Function GetNextItemID(Optional Family As String) As String
'STILL STRING
Dim iMenu As New zclsMenu
GetNextItemID = iMenu.GetNextItemID(Family)
Set iMenu = Nothing
End Function

Public Sub CreateNewSpecialItem(ItemID As Integer, ItemName As String, Price As Currency)
Dim iMenu As New zclsMenu
iMenu.CreateNewSpecialItem ItemID, ItemName, Price
Set iMenu = Nothing
End Sub

Public Function GetItemClassCode(ItemID As Integer) As Integer
Dim iMenu As New zclsMenu
GetItemClassCode = iMenu.GetItemClassCode(ItemID)
Set iMenu = Nothing
End Function
