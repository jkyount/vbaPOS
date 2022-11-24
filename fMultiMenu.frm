VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMultiMenu 
   Caption         =   "UserForm1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19410
   OleObjectBlob   =   "fMultiMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fMultiMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddMembers_Click()
If Not ValidComboBoxValue(AvailableFamilies) = True Then
    MsgBox "Invalid selection.  Please choose a value from the dropdown menu."
    Exit Sub
End If

Dim iMultiMenu As New aclsMultiMenu
iMultiMenu.AddMember Family.value, AvailableFamilies.value
Set iMultiMenu = Nothing
AddOrRemoveMember_Cleanup
MsgBox "Members updated."
End Sub

Private Sub AddOrRemoveMember_Cleanup()
RefreshMultiMenuComboBoxes Family.value
Dim iFamily As New zclsFamily
iFamily.Family = Family.value
UpdateMenuButtons iFamily.FamilyGroup
ResetGUI
End Sub

Private Sub Cancel_Click()
ResetGUI
End Sub

Private Sub RemoveMember_Click()
If Not ValidComboBoxValue(CurrentFamilies) = True Then
    MsgBox "Invalid selection.  Please choose a value from the dropdown menu."
    Exit Sub
End If
Dim iMultiMenu As New aclsMultiMenu
iMultiMenu.RemoveMember Family.value, CurrentFamilies.value
Set iMultiMenu = Nothing
AddOrRemoveMember_Cleanup
MsgBox "Members updated."
End Sub

Private Sub RemoveMultiMenu_Click()
Dim iFamily As New zclsFamily
iFamily.Family = Me.Family.value
Dim FamilyGroup As String
FamilyGroup = iFamily.FamilyGroup
Set iFamily = Nothing
RemoveAllMembers Me.Family.value
DeleteMultiMenu Me.Family.value
UpdateMenuButtons FamilyGroup
End Sub

Private Sub DeleteMultiMenu(MultiMenuName As String)
Dim iMultiMenu As New aclsMultiMenu
iMultiMenu.Delete MultiMenuName
Set iMultiMenu = Nothing
End Sub

Private Sub RemoveAllMembers(MultiMenuName As String)
Dim iMultiMenu As New aclsMultiMenu
Dim coll As Collection
Set coll = iMultiMenu.Members(MultiMenuName)
Dim i As Integer
For i = 1 To coll.Count
    iMultiMenu.RemoveMember MultiMenuName, coll(i)
Next i
Set iMultiMenu = Nothing
Set coll = Nothing
End Sub

Private Sub SaveNewMultiMenu_Click()
Me.MultiMenu.value = True
Me.MultiMenuID.value = DisplayName.value
Dim iFamily As New zclsFamily
iFamily.AddNew GetFormValueDict(EditMultiMenuFrame)
Set iFamily = Nothing
MsgBox ("MultiMenu created.  MultiMenus with no members will not be active on the order screen.  A MultiMenu will be automatically activated after it has been assigned at least one member, and automatically deactivated if all of its members are removed.")
InitializeComboBoxes
ResetGUI
End Sub

Private Sub UserForm_Activate()
Me.Height = ActiveWindow.Height
Me.Width = ActiveWindow.Width
Me.Top = ActiveWindow.Top
Me.Left = ActiveWindow.Left
InitializeComboBoxes
ResetGUI
End Sub

Public Sub InitializeComboBoxes()
Dim iFamily As New zclsFamily
Dim arr As Variant
Dim iDataObj As aclsDataObject
Set iDataObj = iFamily.Wrap(GetNewMatchObj("MultiMenu", True))
arr = FilteredMatch(iDataObj, "Family")
PopDropDown Me.MultiMenuSelect, arr
Set iFamily = Nothing
End Sub

Private Sub ResetGUI()
MultiMenuSelect.Enabled = True
DisplayCtrls GetCtrlDict(Me, "EditMultiMenuFrame", "SaveChanges", "SaveNewMultiMenu", "RemoveMultiMenu"), False
ClearControls
End Sub

Private Sub ClearControls()
Me.Family.value = ""
Me.DisplayName.value = ""
Me.MultiMenuID.value = ""
Me.MenuStyle.value = "Condensed"
Me.FamilyGroup.value = "MenuCategory"
Me.MultiMenu.value = True
EnableCtrls GetCtrlDict(Me, "FamilyGroup", "MenuStyle"), False
End Sub

Private Sub EditMultiMenu_Click()
If MultiMenuSelect.value = "" Then
    Exit Sub
End If
EnableCtrls GetCtrlDict(Me, "AvailableFamilies", "CurrentFamilies", "AddMembers", "RemoveMember"), True
DisplayCtrls GetCtrlDict(Me, "SaveChanges", "RemoveMultiMenu"), True
EnableCtrls GetCtrlDict(Me, "Family", "MultiMenuSelect", "FamilyGroup", "MenuStyle"), False
CurrentFamilies.value = ""
ActivateMultiMenuEditScreen MultiMenuSelect.value
Family.value = MultiMenuSelect.value
AvailableFamilies.value = ""
End Sub

Private Sub ActivateMultiMenuEditScreen(SelectedMultiMenu As String)
Dim iMultiMenu As New aclsMultiMenu
RefreshMultiMenuComboBoxes SelectedMultiMenu
DisplaySelectedFamilyParams SelectedMultiMenu
EditMultiMenuFrame.Visible = True
End Sub

Private Sub RefreshMultiMenuComboBoxes(SelectedMultiMenu As String)
Dim iMultiMenu As New aclsMultiMenu
PopDropDown Me.AvailableFamilies, iMultiMenu.GetAvailableFamilies(SelectedMultiMenu)
PopDropDown Me.CurrentFamilies, iMultiMenu.GetCurrentFamilies(SelectedMultiMenu)
Set iMultiMenu = Nothing
End Sub

Private Sub DisplaySelectedFamilyParams(SelectedFamily As String)
Dim iFamily As New zclsFamily
PopFormWithValues GetValueDict(iFamily.Wrap(GetNewMatchObj("Family", SelectedFamily)))(1), EditMultiMenuFrame
Set iFamily = Nothing
End Sub

Private Sub AddNewMultiMenu_Click()
If ValidRequest = False Then
    Exit Sub
End If
EnableCtrls GetCtrlDict(Me, "Family", "FamilyGroup"), True
DisplayCtrls GetCtrlDict(Me, "SaveNewMultiMenu"), True
DisplayCtrls GetCtrlDict(Me, "SaveChanges", "RemoveMultiMenu"), False
EnableCtrls GetCtrlDict(Me, "MultiMenu", "MenuStyle", "AvailableFamilies", "CurrentFamilies", "AddMembers", "RemoveMember"), False
ActivateMultiMenuEditScreen "None"
MultiMenuSelect.value = ""
Family.value = ""
AvailableFamilies.value = ""
CurrentFamilies.value = ""
End Sub

Private Sub SaveChanges_Click()
UpdateMenu Family.value
'Dim iBtn As New aclsItemButton
'Dim iFamily As zclsFamily
'Set iFamily = iFamily.GetFamilyButtonObj(Family.value)
''iBtn.PositionAll iFamily
ResetGUI
MsgBox "MultiMenu updated."
'Set iBtn = Nothing
End Sub

Private Sub UpdateMenu(SelectedFamily As String)
SaveFamilyChanges SelectedFamily
Dim iFamily As New zclsFamily
iFamily.Family = SelectedFamily
UpdateMenuButtons iFamily.FamilyGroup
End Sub

Private Sub SaveFamilyChanges(SelectedFamily As String)
Dim iFamily As New zclsFamily
Dim iDataObject As New aclsDataObject
Set iDataObject = iFamily.Wrap(iDataObject.GetNewMatchObj("Family", SelectedFamily))
UpdateFromDict iDataObject, GetFormValueDict(EditMultiMenuFrame)
Set iFamily = Nothing
Set iDataObject = Nothing
End Sub

Private Function ValidRequest() As Boolean
ValidRequest = True
End Function




