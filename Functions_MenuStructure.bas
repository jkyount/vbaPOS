Attribute VB_Name = "Functions_MenuStructure"
Option Explicit

Public Function GetItemFamilies() As Variant
Dim iFamily As New zclsFamily
GetItemFamilies = FltrdOrdrdMtch(iFamily.Wrap(GetNewMatchObj("NOT ID", 0, "NOT MultiMenu", True)), "ID ASC", "Family")
Set iFamily = Nothing
End Function

Public Function GetUnrestrictedItemFamilies() As Variant
Dim iFamily As New zclsFamily
GetUnrestrictedItemFamilies = FltrdOrdrdMtch(iFamily.Wrap(GetNewMatchObj("NOT ID = 0 AND NOT Restricted", True, "NOT MultiMenu", True)), "ID ASC", "Family")
Set iFamily = Nothing
End Function

Public Function GetUnfixedItemFamilies() As Variant
Dim iFamily As New zclsFamily
GetUnfixedItemFamilies = FltrdOrdrdMtch(iFamily.Wrap(GetNewMatchObj("NOT ID = 0 AND NOT Fixed", True, "NOT MultiMenu", True)), "ID ASC", "Family")
Set iFamily = Nothing
End Function


Public Function GetItemCategories() As Variant
Dim iCategory As New zclsCategory
GetItemCategories = FilteredMatch(iCategory.Wrap(GetNewMatchObj("NOT ID", 0)), "Category")
Set iCategory = Nothing
End Function

Public Function GetFamilyGroups() As Variant
Dim iFamilyGroup As New zclsFamilyGroup
GetFamilyGroups = FilteredMatch(iFamilyGroup.Wrap(GetNewMatchObj("NOT ID", 0)), "FamilyGroup")
Set iFamilyGroup = Nothing
End Function

Public Function GetMultiMenuIDs() As Variant
Dim iFamily As New zclsFamily
GetMultiMenuIDs = FilteredMatch(iFamily.Wrap(GetNewMatchObj("MultiMenu", True)), "MultiMenuID")
Set iFamily = Nothing
End Function

Public Function GetMenuStyles() As Variant
Dim iMenuStyle As New aclsMenuStyle
GetMenuStyles = FilteredMatch(iMenuStyle.Wrap(GetNewMatchObj("NOT ID", 0)), "MenuStyle")
Set iMenuStyle = Nothing
End Function

'Public Function CIntItemID(ItemID As String) As Integer
'CIntItemID = CInt(right(ItemID, Len(ItemID) - 4))
'End Function


