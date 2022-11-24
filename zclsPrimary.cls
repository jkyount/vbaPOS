VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsPrimary"
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
Me.ItemID = ItemID
'ThisItem.ItemID = ItemID
End Sub

Public Sub SpclConfig(item As aclsItem)

End Sub

Public Sub InheritParentPrintParams(item As aclsItem)
End Sub


Public Sub zclsItem_RefreshPreviewWindow()
RefreshPreviewWindow
End Sub

Public Sub zclsItem_UpdateGUI()
ShowShape "frmSaladFrame"
End Sub


Public Function GetItemIndicator() As String
GetItemIndicator = "> "
End Function



