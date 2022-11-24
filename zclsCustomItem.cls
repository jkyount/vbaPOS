VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsCustomItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Implements zclsItem

Private pItemID As Integer

Public Property Let ItemID(value As Integer)
pItemID = value
End Property
Public Property Get ItemID() As Integer
ItemID = pItemID
End Property

Public Sub ItemInitialize(ItemID As Integer)
Dim NextID As Integer, ItemName As String
Dim Price As Currency
NextID = GetNextItemID("CustomItem")
ItemName = Sheet1.OLEObjects("CustomItemName").Object.text

Price = Sheet1.OLEObjects("CustomItemPrice").Object.value
Me.ItemID = NextID
CreateNewSpecialItem Me.ItemID, ItemName, Price
End Sub

Public Sub InheritParentPrintParams(item As aclsItem)
End Sub

Public Sub SpclConfig(item As aclsItem)

End Sub

Public Sub zclsItem_RefreshPreviewWindow()

End Sub

Public Sub zclsItem_UpdateGUI()
HideShape "grpCustomItem"
End Sub


Public Function GetItemIndicator() As String
GetItemIndicator = "> "
End Function
