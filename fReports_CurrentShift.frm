VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fReports_CurrentShift 
   Caption         =   "UserForm1"
   ClientHeight    =   8265.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16935
   OleObjectBlob   =   "fReports_CurrentShift.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fReports_CurrentShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iReport As New aclsReport



Private Sub CommandButton20_Click()
EndDay
Me.Hide
End Sub

Private Sub CurrentClockIn_Click()
Dim iCurrentClockIn As New rclsCurrentClockIn
iCurrentClockIn.DataSource = GetDailyDataObj
iReport.RunReport iCurrentClockIn
End Sub

Private Sub ItemCounts_Click()
Dim iItemCount As New rclsItemCount
iItemCount.DataSource = GetDailyDataObj
iReport.RunReport iItemCount
End Sub

Private Sub SalesReport_Click()
Dim iSales As New tclsSalesReport

Dim iSalesSummary As New rclsSalesSummary
iSales.ReportType = iSalesSummary
iSales.DataSource = GetDailyDataObj
iReport.RunReport iSales

End Sub














Private Sub ServerReport_Click()
Dim iSales As New tclsSalesReport
Dim iServerReport As New rclsServerReport
iServerReport.DataSource = GetDailyDataObj
iSales.ReportType = iServerReport
iReport.ServerNum = GetFirstWord(Server.value)
iReport.RunReport iSales
End Sub

Private Sub UserForm_Activate()
Me.Height = ActiveWindow.Height
Me.Width = ActiveWindow.Width
Me.Top = ActiveWindow.Top
Me.Left = ActiveWindow.Left
InitializeDropdowns
iReport.StartDate = Now
iReport.EndDate = Now
End Sub

Private Sub InitializeDropdowns()
Dim iEmployee As New zclsEmployee
PopDropDown Me.Server, iEmployee.GetEmployees
PopDropDown Me.Checks, FormatCheckArray(GetChecks("NOT CheckNumber", ""))
End Sub

Private Function FormatCheckArray(Checks As Variant) As Variant
Dim TempArray() As Variant
Dim junkarray(0 To 0, 0 To 0) As Variant
ReDim TempArray(1 To UBound(Checks) + 1)
Dim i As Integer
For i = 0 To UBound(Checks)
    junkarray(0, 0) = Checks(i, 0) & " [" & Checks(i, 1) & "]  [" & Checks(i, 2) & "]  [" & Format(Checks(i, 3), "$0.00") & "]  [" & Checks(i, 4) & "]"
    TempArray(i + 1) = junkarray
Next i
FormatCheckArray = TempArray
End Function

Private Sub ViewCheck_Click()
Dim iCheckReport As New rclsCheckReport
iCheckReport.DataSource = GetDailyDataObj
iReport.check = GetFirstWord(Me.Checks.value)
iReport.RunReport iCheckReport
End Sub
