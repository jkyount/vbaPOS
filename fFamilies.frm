VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fFamilies 
   Caption         =   "UserForm1"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20010
   OleObjectBlob   =   "fFamilies.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fFamilies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActivateFamilyEditScreen(SelectedFamily As String)
DisplaySelectedFamilyParams SelectedFamily
EditFamilyFrame.Visible = True
End Sub

Private Sub AddNewFamily_Click()
ActivateFamilyEditScreen ""
DisplayCtrls GetCtrlDict(Me, "SaveNewFamily", "Cancel"), True
CtrlEnabled Family, True
Family.value = ""
EnableCtrls GetCtrlDict(Me, "FamilySelect", "Active"), False
FamilySelect.value = ""
Active.value = False
End Sub

Private Sub Cancel_Click()
ResetGUI
End Sub

Private Sub ClearControls()
ClassCode.value = 0
Dim ctrl As Control
For Each ctrl In fFamilies.EditFamilyFrame.Controls
    If Not TypeName(ctrl) = "Label" Then
        If Not TypeName(ctrl) = "Frame" Then
            ctrl.value = False
            If Not TypeName(ctrl) = "CheckBox" Then
                ctrl.value = ""
            End If
        End If
    End If
Next ctrl
End Sub
Private Sub ClearErrorLabels()
Dim ctrl As Control
    For Each ctrl In fFamilies.EditFamilyFrame.Controls
        If TypeName(ctrl) = "Label" Then
            If ctrl.name Like "Error*" Then
                ctrl.Caption = ""
            End If
        End If
    Next ctrl
End Sub

Private Sub CommitChanges()
If ValidFormValues = True Then
    UpdateMenu Family.value
'    Dim iBtn As New aclsItemButton
'    Dim iFamily As zclsFamily
'    Set iFamily = New zclsFamily
'    Set iFamily = iFamily.GetFamilyButtonObj(Family.value)
'    'Btn.PositionAll iFamily
    ResetGUI
    MsgBox "Family structure updated."
    'Set iBtn = Nothing
End If
End Sub

Private Sub CtrlDisplay(ctrl As MSForms.Control, value As Boolean)
ctrl.Visible = value
End Sub

Private Sub CtrlEnabled(ctrl As MSForms.Control, value As Boolean)
ctrl.Enabled = value
End Sub

Private Sub DisplaySelectedFamilyParams(SelectedFamily As String)
Dim iFamily As New zclsFamily
PopFormWithValues GetValueDict(iFamily.Wrap(GetNewMatchObj("Family", SelectedFamily)))(1), EditFamilyFrame
Set iFamily = Nothing
Dim i As Integer
For i = 0 To UserClassCode.ListCount
    
    If UserClassCode.List(i) Like ClassCode.value & "*" Then
        UserClassCode.value = UserClassCode.List(i)
        Exit For
    End If
Next i

End Sub

Private Sub EditFamily_Click()
If FamilySelect.value = "" Or FamilySelect.value = "Select a family.." Then
    Exit Sub
End If
ActivateFamilyEditScreen FamilySelect.value
Family.value = FamilySelect.value
EnableCtrls GetCtrlDict(Me, "Family", "FamilySelect"), False
DisplayCtrls GetCtrlDict(Me, "SaveChanges", "RemoveFamily", "Cancel"), True
End Sub



Private Sub FamilyGroup_Change()
If Not FamilyGroup.value = "Component" Then
    MenuStyle.Enabled = True
    Exit Sub
End If
MenuStyle.value = "Component"
MenuStyle.Enabled = False

End Sub

Public Sub InitializeComboBoxes()
Dim iMenu As New zclsMenu
Dim arr As Variant
arr = GetUnrestrictedItemFamilies
PopDropDown Me.FamilySelect, arr
arr = GetFamilyGroups
PopDropDown Me.FamilyGroup, arr
arr = GetMultiMenuIDs
PopDropDown Me.MultiMenuParent, arr
arr = GetMenuStyles
PopDropDown Me.MenuStyle, arr
Set iMenu = Nothing
Dim iItemclass As New zclsItemClass
arr = iItemclass.GetAssignableClassNames
PopDropDown Me.UserClassCode, arr
End Sub

Private Sub ResetGUI()
EnableCtrls GetCtrlDict(Me, "FamilySelect", "MenuStyle"), True
DisplayCtrls GetCtrlDict(Me, "EditFamilyFrame", "SaveChanges", "SaveNewFamily", "RemoveFamily", "Cancel"), False
ClearControls
ClearErrorLabels
FamilySelect.value = "Select a family.."
End Sub







Private Sub RemoveFamily_Click()
If Me.MultiMenu.value = True Then
    MsgBox ("This family is designated as a MultiMenu.  It can only be removed from the MultiMenu configuration menu.")
End If
Dim iFamily As zclsFamily
Set iFamily = New zclsFamily
iFamily.Family = Family.value
Dim FamilyGroup As String
FamilyGroup = iFamily.FamilyGroup
iFamily.Remove Family.value
InitializeComboBoxes
ResetGUI
UpdateMenuButtons FamilyGroup
Set iFamily = Nothing
MsgBox "Family removed."
End Sub

Private Sub SaveChanges_Click()
CommitChanges
End Sub

Private Sub SaveNewFamily_Click()
If ValidNewFamilyValues = True Then
    Dim iFamily As New zclsFamily
    iFamily.Family = Family.value
    Dim FamilyDict As Dictionary
    Set FamilyDict = GetFormValueDict(EditFamilyFrame)
    iFamily.AddNew FamilyDict
    Dim iMenu As New zclsMenu
    'iMenu.AddNewFamily FamilyDict
    UpdateMenuButtons iFamily.FamilyGroup
    InitializeComboBoxes
    ResetGUI
    Set iMenu = Nothing
    Set iFamily = Nothing
    MsgBox "Family added."
End If
End Sub

Private Sub SaveFamilyChanges(SelectedFamily As String)
Dim iFamily As New zclsFamily
Dim iDataObject As New aclsDataObject
Set iDataObject = iFamily.Wrap(iDataObject.GetNewMatchObj("Family", SelectedFamily))
UpdateFromDict iDataObject, GetFormValueDict(EditFamilyFrame)
Set iFamily = Nothing
Set iDataObject = Nothing
End Sub

Private Sub UserClassCode_Change()
If Not UserClassCode.value = False Then
    If Not UserClassCode = "" Then
        Me.ClassCode.value = CInt(GetFirstWord(Me.UserClassCode.value))
    End If
End If
End Sub

Private Sub UserForm_Activate()
Me.Height = ActiveWindow.Height
Me.Width = ActiveWindow.Width
Me.Top = ActiveWindow.Top
Me.Left = ActiveWindow.Left
ClassCode.value = 0
InitializeComboBoxes
ResetGUI
End Sub

Private Sub UpdateMenu(SelectedFamily As String)
Dim iFamily As New zclsFamily
iFamily.Family = SelectedFamily
SaveFamilyChanges SelectedFamily
UpdateMenuButtons iFamily.FamilyGroup
End Sub

Private Function ValidFormValues() As Boolean
ClearErrorLabels
ValidFormValues = True
Dim iFamily As zclsFamily
Set iFamily = New zclsFamily
'Set iFamily = iFamily.GetFamilyButtonObj(FamilySelect.value)
iFamily.Family = FamilySelect.value
If Active.value = True Then
    If iFamily.Count = 0 Then
        Error_Active.Caption = "Family cannot be activated until it posesses at least 1 menu item member."
        ValidFormValues = False
    End If
End If

If DisplayName.value = "" Then
    Error_DisplayName.Caption = "Enter a valid display name."
    ValidFormValues = False
End If

If FamilyGroup.value = "" Then
    Error_FamilyGroup.Caption = "Select a family group."
    ValidFormValues = False
End If

If MenuStyle.value = "" Then
    Error_MenuStyle.Caption = "Select a menu style."
    ValidFormValues = False
End If

If MenuStyle = "Component" And Not FamilyGroup = "Component" Then
    Error_MenuStyle.Caption = "Menu Style invalid for selected Family Group."
    ValidFormValues = False
End If

If FamilyGroup = "Component" And Not MenuStyle = "Component" Then
    Error_FamilyGroup.Caption = "Family Group invalid for selected MenuStyle."
    ValidFormValues = False
End If

If ClassCode.value = 0 Then
    Error_ClassCode.Caption = "Select a family class."
    ValidFormValues = False
End If
End Function

Private Function ValidNewFamilyValues() As Boolean
ValidNewFamilyValues = True
If Family.value = "" Then
    Error_Family.Caption = "Enter a valid family name."
    ValidNewFamilyValues = False
End If

Dim arr As Variant
arr = GetItemFamilies
Dim i As Integer
For i = 1 To UBound(arr)
    If Family.value = arr(i)(0, 0) Then
        Error_Family.Caption = "This family name has already been assigned.  Please enter a different name."
        ValidNewFamilyValues = False
        Exit For
    End If
Next i

If ClassCode.value = 0 Then
    Error_ClassCode.Caption = "Select a family class."
    ValidNewFamilyValues = False
End If

If DisplayName.value = "" Then
    Error_DisplayName.Caption = "Enter a valid display name."
    ValidNewFamilyValues = False
End If

If FamilyGroup.value = "" Then
    Error_FamilyGroup.Caption = "Select a family group."
    ValidNewFamilyValues = False
End If

If MenuStyle.value = "" Then
    Error_MenuStyle.Caption = "Select a menu style."
    ValidNewFamilyValues = False
End If

If MenuStyle = "Component" And Not FamilyGroup = "Component" Then
    Error_MenuStyle.Caption = "Menu Style invalid for selected Family Group."
    ValidNewFamilyValues = False
End If

If FamilyGroup = "Component" And Not MenuStyle = "Component" Then
    Error_FamilyGroup.Caption = "Family Group invalid for selected MenuStyle."
    ValidNewFamilyValues = False
End If

    
End Function






