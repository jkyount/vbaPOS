VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fAddEmployee 
   Caption         =   "UserForm1"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465.001
   OleObjectBlob   =   "fAddEmployee.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fAddEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddNewEmployee_Click()
Dim iEmployee As New zclsEmployee
Dim dict As New Dictionary
Set dict = GetNewEmployeeDict
AddNewRecord iEmployee, GetFormValueDict(fNewEmployeeInfo)
fEmployees.InitializeComboBoxes
fEmployees.ResetGUI
MsgBox "Employee added."
Me.Hide

End Sub


