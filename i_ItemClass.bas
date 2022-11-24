Attribute VB_Name = "i_ItemClass"
Option Explicit

Public Function GetItemType(ClassCode As String) As Variant
Dim iItemclass As New zclsItemClass
Set GetItemType = iItemclass.GetItemType(ClassCode)
Set iItemclass = Nothing
End Function
