Attribute VB_Name = "Functions_EndOfDay"
Option Explicit


Public Sub EndDay()
'If Worksheets("Reports").Range("OutstandingAmountCell").value > 0 Then
'    MsgBox ("Please close any outstanding checks before running the Z report.")
'    Exit Sub
'End If
'HideShow "grpEndOfDayWarning"

UpdateMonthlyTotals
ArchiveCheckDetail
ArchiveCheckIndex
''ArchiveDailyItemCounts
ClearAllChecks
'ClearDailyItemCounts
ClearCustomItems
ClearTableStates
ClockOutAllEmployees
ActivateLoginScreen
End Sub

Public Sub UpdateMonthlyTotals()
'v/
Dim iMonthlyTotals As New aclsMonthlyTotals
iMonthlyTotals.Update
MsgBox ("Monthly totals updated.")
Set iMonthlyTotals = Nothing
End Sub

Public Sub ClearAllChecks()
'v/
Dim iIndex As New zclsDailyCheckIndex
Dim iDetail As New zclsDailyCheckDetail
DeleteMatch iIndex.Wrap(GetNewMatchObj("NOT CheckNumber", ""))
DeleteMatch iDetail.Wrap(GetNewMatchObj("NOT CheckNumber", ""))
MsgBox ("Today's Checks removed from database.")
Set iIndex = Nothing
Set iDetail = Nothing
End Sub

Public Sub ClockOutAllEmployees()
'v/
Dim iTimeClock As New aclsTimeClock
Dim iEmployee As New zclsEmployee
iTimeClock.ClockOutAll
iEmployee.ClockOutAll
Set iEmployee = Nothing
Set iTimeClock = Nothing
End Sub

Public Sub ArchiveCheckDetail()
'v/
Dim iDataObject As New aclsDataObject
Dim iDetail As New zclsDailyCheckDetail
ArchiveData iDetail.Wrap(iDataObject)
Set iDetail = Nothing
Set iDataObject = Nothing
End Sub

Public Sub ArchiveCheckIndex()
'v/
Dim iCheckIndex As New zclsDailyCheckIndex
Dim iDataObject As New aclsDataObject
ArchiveData iCheckIndex.Wrap(iDataObject)
Set iCheckIndex = Nothing
Set iDataObject = Nothing
MsgBox ("Today's checks have been archived.")
End Sub

Public Sub ClearCustomItems()
'v/
Dim iMenu As New zclsMenu
iMenu.ClearCustomItems
Set iMenu = Nothing
MsgBox ("Custom items have been deleted.")
End Sub

