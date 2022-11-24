Attribute VB_Name = "i_ItemButton"
Option Explicit

'Public Function GetBtnObj(ItemID As Integer, BtnType As Variant) As aclsItemButton
'Dim BtnDict As Dictionary
'Set BtnDict = GetBtnDict(ItemID, BtnType)
'Dim iItemBtn As aclsItemButton
'Set iItemBtn = New aclsItemButton
'iItemBtn.BtnType = BtnType
'iItemBtn.name = BtnDict("Name")
'iItemBtn.Caption = BtnDict("Caption")
'iItemBtn.Family = BtnDict("Family")
'iItemBtn.StyleDict = BtnDict("StyleDict")
'iItemBtn.Location = BtnDict("Location")
'iItemBtn.LocalGroup = BtnDict("LocalGroup")
'
'iItemBtn.Shape = BtnType.GetShape(ItemID)
'Set GetBtnObj = iItemBtn
'Set iItemBtn = Nothing
'Set BtnDict = Nothing
'End Function
'
'Private Function GetBtnDict(ItemID As Integer, BtnType As Variant)
'Dim dict As Dictionary
'Set dict = New Dictionary
'
'dict.Add "Name", GetName(ItemID)
'dict.Add "Caption", GetCaption(ItemID)
'dict.Add "Family", GetFamily(ItemID)
'dict.Add "StyleDict", GetStyleDict(dict("Family"))
'dict.Add "Location", GetLocation(ItemID)
'dict.Add "LocalGroup", BtnType.GetLocalGroup(ItemID)
'End Function
'
'Private Function GetName(ItemID As Integer) As String
'GetName = "Item" & ItemID
'End Function
'
'Private Function GetCaption(ItemID As Integer) As String
'GetCaption = ValueMatch(GetMenuObj.Wrap(GetNewMatchObj("ID", ItemID)), "ItemName")
'End Function
'Private Function GetFamily(ItemID As Integer) As String
'GetFamily = ValueMatch(GetMenuObj.Wrap(GetNewMatchObj("ID", ItemID)), "Family")
'End Function
'Private Function GetStyleDict(Family As String) As Dictionary
'Dim iFamily As New zclsFamily
'iFamily.Family = Family
'Set GetStyleDict = iFamily.StyleDict
'Set iFamily = Nothing
'End Function
'Private Function GetShape(ItemID As Integer) As Shape
'
'End Function
'Private Function GetLocation(ItemID As Integer) As Worksheet
'Dim Family As String
'Family = ValueMatch(GetMenuObj.Wrap(GetNewMatchObj("ID", ItemID)), "Family")
'End Function
'
'Private Function GetLocalGroup(ItemID As Integer) As Integer
'
'End Function



