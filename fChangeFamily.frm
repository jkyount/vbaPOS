VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fChangeFamily 
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5835
   OleObjectBlob   =   "fChangeFamily.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fChangeFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub ChangeFamily_Click()
'TransferToNewFamily CInt((GetFirstWord(SelectedItem.value))), FamilySelect.value
'MsgBox ("Item reassigned.")
'Me.Hide
'
'End Sub



Private Sub UserForm_Activate()
Dim iDataObj As aclsDataObject
Dim iFamily As New zclsFamily
Dim iMenu As New zclsMenu
Dim iFamilyGroup As New zclsFamilyGroup
Dim ItemID As Integer
ItemID = CInt(GetFirstWord(SelectedItem.value))
Dim SelectedItemFamily As String
SelectedItemFamily = ValueMatch(iMenu.Wrap(GetNewMatchObj("ID", ItemID)), "Family")
Dim SelectedItemFamilyGroup As String
SelectedItemFamilyGroup = ValueMatch(iFamily.Wrap(GetNewMatchObj("Family", SelectedItemFamily)), "FamilyGroup")

PopDropDown FamilySelect, iFamilyGroup.GetMembers(SelectedItemFamilyGroup)




End Sub

