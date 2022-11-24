VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bclsParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pID As Integer
Private pItem As aclsItem
Private pItemID As Integer

Public Property Let ID(value As Integer)
pID = value
End Property
Public Property Get ID() As Integer
ID = pID
End Property

Public Property Let item(value As aclsItem)
Set pItem = value
End Property
Public Property Get item() As aclsItem
Set item = GetItemByID(Me.ID)
End Property

Public Function GetNew(ID As Integer) As bclsParent
Dim x As New bclsParent
x.ID = ID
Set GetNew = x
Set x = Nothing
End Function


Public Property Let ItemID(value As Integer)
pItemID = value
End Property
Public Property Get ItemID() As Integer
ItemID = pItemID
End Property

'Public Sub ItemInitialize(ItemID As Integer)
'Me.ItemID = ItemID
'End Sub

Public Sub SetParent()
ThisItem.Parent = GetNewParent(-1)
End Sub

Public Sub InheritParentPrintParams(item As aclsItem)
End Sub

Public Sub BuildItemCollection()
ClearCollection CItem
ThisItem.IsPrimaryItem = True
CItem.Add ThisItem, CStr(ThisItem.CollID)
SetCurrentParent ThisItem.CollID
End Sub

Public Sub RefreshPreviewWindow()
RefreshPreviewWindow
End Sub

'Public Sub UpdateGUI()
'ShowShape "frmSaladFrame"
'End Sub

Public Sub CheckForRequiredComponents(item As aclsItem)
If item.ItemOptions = False Then
    HideShape "frmSaladFrame"
    OrderQueuedItems CItem
    Exit Sub
End If

If Not item.RequiredComponents.Count = 0 Then
    If MissingComponents = True Then
        Exit Sub
    End If
End If
SetCurrentParent item.CollID
HideShape "ComponentBLOCK"
HideShape "grpScrollCategoryItems"
ShowShape "btnDone"
SetCurrentFamily ""
End Sub

'Public Function GetItemIndicator() As String
'GetItemIndicator = "> "
'End Function








