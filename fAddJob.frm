VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fAddJob 
   Caption         =   "Add Job"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
   OleObjectBlob   =   "fAddJob.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fAddJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub AddJob_Click()
If ValidInput = True Then
    PayRate.BackColor = rgbWhite
    Dim JobCode As Integer
    Dim iEmployee As New zclsEmployee
    JobCode = iEmployee.JobNameToCode(Jobs.value)
    Dim ServerNum As Integer
    ServerNum = CInt(GetFirstWord(fEmployees.Employees.value))
    AddEmployeeJob JobCode, PayRate.value, ServerNum
End If
MsgBox "Job added successfully."
Me.Hide
End Sub

Private Sub Jobs_Change()
If ValidComboBoxValue(Me.Jobs) = False Then Jobs.value = ""
End Sub

Private Sub UserForm_Activate()
Dim iEmployee As New zclsEmployee
PopDropDown Me.Jobs, iEmployee.GetAllJobs
PayRate.value = 0
PayRate.BackColor = rgbWhite

End Sub

Public Function ValidInput() As Boolean
ValidInput = True
If Not IsNumeric(PayRate.value) = True Then
    MsgBox "Please enter a valid pay rate."
    PayRate.BackColor = rgbRed
    ValidInput = False
    Exit Function
End If

If ValidComboBoxValue(Me.Jobs) = False Then
    MsgBox "Invalid Job.  Select a job from the list, or add a new job from the configuration menu."
    Jobs.BackColor = rgbRed
    ValidInput = False
    Exit Function
End If
End Function


