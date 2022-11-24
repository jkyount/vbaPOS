VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fEmployees 
   Caption         =   "Configuration"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   OleObjectBlob   =   "fEmployees.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ServerNum As Integer
Private iEmployee As New zclsEmployee
Private iDataObject As New aclsDataObject
Private JobDict As New Dictionary

Private Sub AddJob_Click()
fAddJob.Show
End Sub

Private Sub AddNewEmployee_Click()
fAddEmployee.Show
End Sub

Private Sub Cancel_Click()
ResetGUI
End Sub

Private Function UpdateJobDict() As Dictionary
Set JobDict = GetJobDict(ServerNum)
End Function

Private Sub Job1_Change()
If ValidComboBoxValue(Me.Job1) = False Then
    Job1.value = ""
    Exit Sub
End If
End Sub

Private Sub Job2_Change()
If ValidComboBoxValue(Me.Job2) = False Then
    Job2.value = ""
    Exit Sub
End If
End Sub

Private Sub Job3_Change()
If ValidComboBoxValue(Me.Job3) = False Then
    Job3.value = ""
    Exit Sub
End If
End Sub

Private Sub RemoveEmployee_Click()
If MsgBox("This will permamently delete this employee from your database.  Proceed?", vbYesNo) = vbYes Then
    DeleteMatch iEmployee.Wrap(GetNewMatchObj("ServerNum", ServerNum))
    InitializeComboBoxes
    ResetGUI
End If
End Sub

Private Sub SaveChanges_Click()
If ValidInput = True Then
 UpdateFromDict iEmployee.Wrap(GetNewMatchObj("ServerNum", ServerNum)), GetEmployeeUpdateDict
End If
   
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
Dim arr As Variant
arr = iEmployee.GetEmployees
PopDropDown Me.Employees, arr
arr = iEmployee.GetAllJobs
Dim i As Integer
For i = 1 To 3
    PopDropDown Me.Controls("Job" & i), arr
    fEmployees.Controls("Payrate" & i).BackColor = rgbWhite
Next i
End Sub


Public Sub ResetGUI()
DisplayCtrls GetCtrlDict(Me, "EmployeeInfo", "RemoveEmployee", "FunctionButtons"), False
SaveChanges.Visible = True
Employees.Enabled = True

Dim ctrl As MSForms.Control
For Each ctrl In fEmployees.Controls
    If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then
        ctrl.BackColor = rgbWhite
    End If
Next ctrl

End Sub

Private Sub EditEmployeeInfo_Click()
DisplayCtrls GetCtrlDict(Me, "EmployeeInfo", "RemoveEmployee", "FunctionButtons", "SaveChanges"), True
Employees.Enabled = False
ServerNum = CInt(GetFirstWord(Employees.value))
UpdateJobDict
DisplayPersonalInfo ServerNum
DisplayJobInfo
End Sub

Private Sub DisplayPersonalInfo(ServerNum As Integer)
PopFormWithValues GetValueDict(iEmployee.Wrap(GetNewMatchObj("ServerNum", ServerNum)))(1), EmployeeInfo
End Sub

Private Sub DisplayJobInfo()
Dim key As Variant
For Each key In JobDict.Keys
    Dim JobCode As Integer
    fEmployees.EmployeeInfo.Controls(key).value = iEmployee.JobCodeToName(JobDict(key)("Job"))
Next key
End Sub

Private Function GetEmployeeUpdateDict() As Dictionary
Dim dict As New Dictionary
Dim key As String
Dim val As Variant
Dim ctrl As MSForms.Control
For Each ctrl In fEmployees.EmployeeInfo.Controls
    If Not TypeName(ctrl) = "Label" Then
        If Not TypeName(ctrl) = "Frame" Then
        key = ctrl.name
        val = ctrl.value
        If ctrl.name Like ("Job*") = True Then
            val = iEmployee.JobNameToCode(ctrl.value)
        End If
        
        dict.Add key, val
        End If
    End If
Next ctrl

Set GetEmployeeUpdateDict = dict
End Function

Public Function ValidInput() As Boolean
ValidInput = True

If IDIsAvailable(IDNumber.value) = False Then
    MsgBox "Login ID already in use.  Please enter a different Login ID."
    IDNumber.BackColor = rgbRed
    ValidInput = False
    Exit Function
End If

Dim i As Integer

For i = 1 To 3
    If Not IsNumeric(fEmployees.Controls("Payrate" & i).value) = True _
    Or fEmployees.Controls("Payrate" & i).value < 0 Then
        MsgBox "Please enter a valid pay rate."
        fEmployees.Controls("Payrate" & i).BackColor = rgbRed
        ValidInput = False
        Exit Function
    End If
Next i

If Not IsNumeric(IDNumber.value) = True Then
    MsgBox "Please enter a numeric LoginID."
    IDNumber.BackColor = rgbRed
    ValidInput = False
    Exit Function
End If

For i = 1 To 3
    If ValidComboBoxValue(fEmployees.Controls("Job" & i)) = False Then
        MsgBox "Invalid Job.  Select a job from the list, or add a new job from the job configuration menu."
        fEmployees.Controls("Job" & i).BackColor = rgbRed
        ValidInput = False
        Exit Function
    End If
Next i
End Function

Public Function IDIsAvailable(ID As Long) As Boolean
Dim arr As Variant
arr = FilteredMatch(iEmployee.Wrap(GetNewMatchObj("NOT IDNumber", 0)), "IDNumber", "ServerNum")
Dim i As Integer
For i = 1 To UBound(arr)
    If ID = arr(i)(0, 0) Then
        If arr(i)(1, 0) = ServerNum Then
            IDIsAvailable = True
            Exit Function
        End If
        IDIsAvailable = False
        Exit Function
    End If
Next i
Set iEmployee = Nothing
IDIsAvailable = True
End Function

