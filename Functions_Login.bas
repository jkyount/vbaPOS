Attribute VB_Name = "Functions_Login"
Option Explicit

Public Sub ActivateLoginScreen()
Sheet11.Activate
Sheet11.LoginBox.value = ""
Sheet11.LoginBox.Activate
End Sub


Public Sub Login()
If IsNumeric(Sheet11.LoginBox.value) = False Then
    Exit Sub
End If
Dim LoginID As Long
LoginID = Sheet11.LoginBox.value

If IsLoginValid(LoginID) = False Then
    MsgBox ("Invalid ID")
    Sheet11.LoginBox.value = ""
    Exit Sub
End If
ThisEmployee.Reset
ThisEmployee.IDNumber = LoginID
If ThisEmployee.ClockedIn = False Then
    MsgBox ("Please clock in.")
    Sheet11.LoginBox.value = ""
    Exit Sub
End If

IndicateTablesInUse
ActivateHomeScreen
End Sub

Public Sub CLICK_ClockInOut()

If IsNumeric(Sheet11.LoginBox.value) = False Then Exit Sub
Dim LoginID As Long
LoginID = Sheet11.LoginBox.value
ThisEmployee.Reset
ThisEmployee.IDNumber = LoginID
If IsLoginValid(LoginID) = False Then
    MsgBox ("Invalid ID")
    Sheet11.LoginBox.value = ""
    Exit Sub
End If
If Not ThisEmployee.ClockedIn = False Then
    Select Case MsgBox("Clock out?", vbYesNo)
        Case vbYes
            ClockOut LoginID
            Sheet11.LoginBox.value = ""
            Exit Sub
        Case vbNo
            Sheet11.LoginBox.value = ""
            Exit Sub
    End Select
End If
ClockIn LoginID
Sheet11.LoginBox.value = ""
End Sub

Public Sub ClockIn(ID As Long)
ThisEmployee.ClockIn ID
TimeClock_ClockIn ID
End Sub

Public Sub ClockOut(ID As Long)
ThisEmployee.ClockOut ID
TimeClock_ClockOut ID
End Sub



