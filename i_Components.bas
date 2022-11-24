Attribute VB_Name = "i_Components"
Option Explicit

Public Function GetNewChild(ID As Integer) As bclsChild
Dim iChild As New bclsChild
Set GetNewChild = iChild.GetNew(ID)
Set iChild = Nothing
End Function

Public Function GetNewChildren() As bclsChildren
Dim iChildren As New bclsChildren
Set GetNewChildren = iChildren.GetNew
Set iChildren = Nothing
End Function

Public Function GetNewParent(ID As Integer) As bclsParent
Dim iParent As New bclsParent
Set GetNewParent = iParent.GetNew(ID)
Set iParent = Nothing
End Function

Public Function GetItemByID(ID As Integer) As aclsItem
Dim iItem As aclsItem
For Each iItem In CItem
    If iItem.CollID = ID Then
        Set GetItemByID = iItem
        Exit Function
    End If
Next iItem

End Function

Public Function GetItemByParentID(ParentID As Integer) As aclsItem
Dim iItem As aclsItem
For Each iItem In CItem
    If iItem.Parent.ID = ParentID Then
        Set GetItemByParentID = iItem
        Exit Function
    End If
Next iItem
End Function

Public Function GetPrimaryItem() As aclsItem
Dim iItem As aclsItem
For Each iItem In CItem
    If iItem.Parent.ID = -1 Then
        Set GetPrimaryItem = iItem
        Exit Function
    End If
Next iItem
End Function

Public Sub RemoveItemFromQueue(item As aclsItem)
Dim child As bclsChild
If item.Children.coll.Count > 0 Then
    For Each child In item.Children.coll
        RemoveItemFromQueue GetItemByID(child.ID)
        item.Children.coll.Remove CStr(child.ID)
    Next child
End If
CItem.Remove CStr(item.CollID)

End Sub

Public Function NormalizePrintParameters(coll As Collection) As Collection

Dim ThisChild As aclsItem
Dim child As bclsChild
For Each child In coll(1).Children.coll
    Set ThisChild = GetItemByID(child.ID)
    ThisChild.ItemType.InheritParentPrintParams ThisChild
    ApplyParentParamToChildren ThisChild, ThisChild.PrintKitchen
Next child
Set NormalizePrintParameters = coll
End Function

Public Sub ApplyParentParamToChildren(item As aclsItem, param As Variant)
Dim child As bclsChild
    If item.Children.coll.Count > 0 Then
    For Each child In item.Children.coll
        ApplyParentParamToChildren GetItemByID(child.ID), param
        GetItemByID(child.ID).PrintKitchen = param
    Next child
End If
item.PrintKitchen = param
End Sub

Public Function FormatItemCollection(coll As Collection) As Collection
Dim tempcoll As Collection
Set tempcoll = FormatSides(FormatChildSpacing(OrderByParent(NormalizePrintParameters(coll))))

'FormatChildSpacing tempcoll
'FormatSides tempcoll
Set FormatItemCollection = tempcoll
End Function

