Attribute VB_Name = "i_Employee"
Option Explicit
Public Function GetEmployeeObj(ID As Long)
Dim iEmployee As zclsEmployee
Set iEmployee = New zclsEmployee
iEmployee.IDNumber = ID
Set GetEmployeeObj = iEmployee
Set iEmployee = Nothing
End Function
Public Function IsLoginValid(IDNumber As Long) As Boolean
Dim iEmployee As New zclsEmployee
IsLoginValid = iEmployee.IsLoginValid(IDNumber)
Set iEmployee = Nothing
End Function

Public Sub AddEmployeeJob(JobCode As Integer, PayRate As Double, ServerNum As Integer)
Dim iEmployee As New zclsEmployee
iEmployee.AddNewJob JobCode, PayRate, ServerNum
Set iEmployee = Nothing
End Sub

Public Sub RemoveEmployeeJob(JobCode As Integer, ServerNum As Integer)
Dim iEmployee As New zclsEmployee
iEmployee.RemoveJob JobCode, ServerNum
Set iEmployee = Nothing
End Sub

Public Function GetJobDict(ServerNum As Integer) As Dictionary
Dim iEmployee As New zclsEmployee
Set GetJobDict = iEmployee.GetJobDict(ServerNum)
Set iEmployee = Nothing
End Function
