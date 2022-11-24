VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMenuItems 
   Caption         =   "MenuItems"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   OleObjectBlob   =   "fMenuItems.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fMenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ItemID As Integer
Private FamilyName As String

Private Sub Cancel_Click()
ResetGUI
End Sub

Private Sub ChangeFamily_Click()
If NoItemSelected = True Then Exit Sub
fChangeFamily.SelectedItem.value = ItemSelect.value
fChangeFamily.Show
InitializeComboBoxes
ResetGUI
End Sub

Private Sub EditMenuItem_Click()
If NoItemSelected = True Then Exit Sub
ItemID = CInt(GetFirstWord(ItemSelect.value))
FamilyName = MenuItems_FamilySelect.value
ActivateItemEditScreen ItemID
EnableCtrls GetCtrlDict(Me, "ItemSelect", "MenuItems_FamilySelect"), False
DisplayCtrls GetCtrlDict(Me, "SaveChanges", "RemoveItem"), True
End Sub

Private Sub AddNewItem_Click()
If ValidRequest = False Then
    Exit Sub
End If
FamilyName = MenuItems_FamilySelect.value
ItemID = CInt(GetNextItemID)

ActivateItemEditScreen ItemID
Family.value = MenuItems_FamilySelect.value
ItemSelect.value = ItemID
EnableCtrls GetCtrlDict(Me, "ItemSelect", "MenuItems_FamilySelect"), False
SaveNewItem.Visible = True
DisplayCtrls GetCtrlDict(Me, "SaveChanges", "RemoveItem"), False
End Sub














Private Sub SaveChanges_Click()
UpdateMenu ItemID
'UpdateButton ItemID
ResetGUI
MsgBox "Item parameters updated."
End Sub

Private Sub RemoveItem_Click()
Dim iMenu As New zclsMenu
iMenu.Remove ItemID
'Dim iBtn As New aclsItemButton
'iBtn.Remove FamilyName, ItemID
ResetGUI
MsgBox "Item removed."
End Sub

Private Sub SaveNewItem_Click()
Dim iFamily As New zclsFamily
iFamily.Family = FamilyName
If IsEmpty(iFamily.Members) Then
    'InitializeNewFamily iFamily, ItemID, Me.ItemName.value
    iFamily.Activate
    UpdateMenu ItemID
    ResetGUI
    MsgBox "Item parameters updated."
    Exit Sub
End If
UpdateMenu ItemID
'AddNewButton ItemID
'Dim iBtn As New aclsItemButton
'iBtn.PositionAll GetFamilyButtonObj(iFamily.Family)
ResetGUI
MsgBox "Item parameters updated."
End Sub

Private Sub UserForm_Activate()
Me.Height = ActiveWindow.Height
Me.Width = ActiveWindow.Width
Me.Top = ActiveWindow.Top
Me.Left = ActiveWindow.Left
InitializeComboBoxes
ClassCode.value = 0
ResetGUI
End Sub

Public Sub InitializeComboBoxes()
Dim arr As Variant
arr = GetUnfixedItemFamilies
    PopDropDown Me.Family, arr
    PopDropDown Me.MenuItems_FamilySelect, arr
Dim iFamilyGroup As New zclsFamilyGroup
iFamilyGroup.FamilyGroup = "Component"
arr = iFamilyGroup.Members
    PopDropDown Me.Req1, arr
    PopDropDown Me.Req2, arr
arr = GetItemCategories
    PopDropDown Me.Category, arr
Set iFamilyGroup = Nothing
End Sub

Private Sub MenuItems_FamilySelect_Change()
If ValidComboBoxValue(MenuItems_FamilySelect) = False Then Exit Sub
UpdateItemSelectList MenuItems_FamilySelect.value
FamilyName = MenuItems_FamilySelect.value
End Sub

Private Sub UpdateItemSelectList(FamilyName As String)
ItemSelect.Clear
ItemSelect.Enabled = True
Dim iMenu As New zclsMenu
PopDropDown Me.ItemSelect, iMenu.PairItemNameAndID(iMenu.GetItemsInFamily(FamilyName))
Set iMenu = Nothing
End Sub

Private Sub ActivateItemEditScreen(ItemID As Integer)
DisplaySelectedItemParams ItemID
ItemEditFrame.Visible = True
End Sub

Private Sub DisplaySelectedItemParams(ItemID As Integer)
Dim iMenu As New zclsMenu
PopFormWithValues GetValueDict(iMenu.Wrap(GetNewMatchObj("ID", ItemID)))(1), ItemEditFrame
Set iMenu = Nothing
End Sub

'Public Sub UpdateButton(ItemID As Integer)
'Dim iItemButton As New aclsItemButton
'Dim iFamily As New zclsFamily
'Set iFamily = iFamily.GetFamilyButtonObj(FamilyName)
'iItemButton.SetCaption iFamily.BtnLocation.Shapes(CStr(ItemID)), ItemName.value
'End Sub

'Private Sub AddNewButton(ItemID As Integer)
'Set ThisItem = New aclsItem
'ThisItem.Initialize ItemID
'Dim iBtn As New aclsItemButton
'iBtn.AddNew ThisItem, ThisItem.Family
'Set ThisItem = Nothing
'Set iBtn = Nothing
'End Sub

Private Sub UpdateMenu(ItemID As Integer)
SaveMenuItemChanges ItemID
End Sub

Private Sub ResetGUI()
MenuItems_FamilySelect.Enabled = True
ItemSelect.Enabled = False
ItemEditFrame.Visible = False
SaveChanges.Visible = False
SaveNewItem.Visible = False
RemoveItem.Visible = False
MenuItems_FamilySelect.value = "Select a family.."
ItemSelect.value = ""
ClassCode.value = 0
End Sub

Private Sub SaveMenuItemChanges(ItemID As Integer)
Dim iMenu As New zclsMenu
Dim iDataObject As New aclsDataObject
Set iDataObject = iMenu.Wrap(iDataObject.GetNewMatchObj("ID", ItemID))
Dim iFamily As New zclsFamily
ClassCode.value = CInt(ValueMatch(iFamily.Wrap(GetNewMatchObj("Family", MenuItems_FamilySelect.value)), "ClassCode"))
UpdateFromDict iDataObject, GetFormValueDict(ItemEditFrame)
Set iMenu = Nothing
Set iDataObject = Nothing
Set iFamily = Nothing
End Sub

Private Function ValidRequest() As Boolean
If NoFamilySelected = True Then
    MsgBox "Invalid family name.  Please select a family from the list."
    ValidRequest = False
    Exit Function
End If
If ValidComboBoxValue(MenuItems_FamilySelect) = False Then
    MsgBox "Invalid family name.  Please select a family from the list."
    ValidRequest = False
    Exit Function
End If
'Dim iMenu As New zclsMenu
'If Match(iMenu.Wrap(GetNewMatchObj("Family", FamilyName, "ItemName", ""))) = True Then
'    ValidRequest = True
'    Exit Function
'End If
ValidRequest = True
End Function

'Public Sub InitializeNewFamily(iFamily As zclsFamily, ItemID As Integer, BtnCaption As String)
'Dim BtnName As String
'BtnName = CStr(ItemID)
'Dim NewShp As Shape
'Set NewShp = Sheet1.Shapes("TemplateButton" & iFamily.MenuStyle).Duplicate
'NewShp.name = BtnName
'NewShp.TextFrame.Characters.text = BtnCaption
'NewShp.ZOrder (msoSendToBack)
'
'Dim BlankShp As Shape
'Set BlankShp = Sheet1.Shapes("TemplateButtonBlank").Duplicate
'BlankShp.name = iFamily.Family & "Blank"
'BlankShp.ZOrder (msoSendToBack)
'
'Dim iBtn As New aclsItemButton
'iBtn.LocalGroup = 1
'iBtn.Shape = NewShp
'Set iFamily = iFamily.GetFamilyButtonObj(iFamily.Family)
'iBtn.Position iFamily, iBtn
'
'BlankShp.Left = NewShp.Left
'BlankShp.Top = NewShp.Top
'
'Dim grp As Shape
'Set grp = Sheet1.Shapes.range(Array(NewShp.name, BlankShp.name)).Group
'grp.name = ("grpgui" & iFamily.Family)
'
'
'Do Until grp.ZOrderPosition > Sheet1.Shapes(iFamily.StyleDict("AboveZOrder")).ZOrderPosition
'    grp.ZOrder msoBringForward
'Loop
'grp.Visible = msoFalse
'End Sub

Private Function NoItemSelected() As Boolean
NoItemSelected = False
If ItemSelect.value = "" Then
    NoItemSelected = True
End If
End Function

Private Function NoFamilySelected() As Boolean
NoFamilySelected = False
If MenuItems_FamilySelect.value Like "Select a *" Then NoFamilySelected = True
If MenuItems_FamilySelect.value = "" Then NoFamilySelected = True
End Function


