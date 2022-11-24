VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fReports 
   Caption         =   "UserForm1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   OleObjectBlob   =   "fReports.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DateBox As MSForms.TextBox
Private ReportType As Variant


Private Sub btnFromDate_Click()
Set DateBox = fReports.FromDate
fDatePicker.Show
End Sub

Private Sub btnToDate_Click()
Set DateBox = fReports.ToDate
fDatePicker.Show
End Sub

Private Sub CurrentShift_Click()
frmDateRange.Visible = False

End Sub

Private Sub DateRange_Click()
frmDateRange.Visible = True
End Sub



Private Sub CurrentClockIn_Click()
ResetForm
Dim iCurrentClockIn As New rclsCurrentClockIn
Set ReportType = iCurrentClockIn
ViewReport.Visible = True
ShowRequiredParams ReportType.RequiredParams
End Sub

Private Sub ItemCount_Click()
ResetForm
DisplayCtrls GetCtrlDict(Me, "frmDateSelect", "ViewReport"), True
'frmDateSelect.Visible = True
'ViewReport.Visible = True
Dim iItemCount As New rclsItemCount
Set ReportType = iItemCount
End Sub



Private Sub MonthlyTotals_Click()
ResetForm
Dim iMonthlyTotals As New rclsMonthlyTotals
Set ReportType = iMonthlyTotals
DisplayCtrls GetCtrlDict(Me, "frmDateSelect", "ViewReport"), True
'frmDateSelect.Visible = True
'ViewReport.Visible = True
ShowRequiredParams ReportType.RequiredParams

End Sub

Private Sub SalesReport_Click()
ResetForm
Dim iSalesReport As New tclsSalesReport
Set ReportType = iSalesReport
Dim iSalesSummary As New rclsSalesSummary
ReportType.ReportType = iSalesSummary
DisplayCtrls GetCtrlDict(Me, "frmDateSelect", "ViewReport"), True
'frmDateSelect.Visible = True
'ViewReport.Visible = True
ShowRequiredParams ReportType.RequiredParams
End Sub

Private Sub ServerReport_Click()
ResetForm
Dim iEmployee As New zclsEmployee
PopDropDown Me.ServerNum, iEmployee.GetEmployees
Dim iSalesReport As New tclsSalesReport
Set ReportType = iSalesReport
Dim iServerReport As New rclsServerReport
ReportType.ReportType = iServerReport
DisplayCtrls GetCtrlDict(Me, "frmDateSelect", "ViewReport", "frmReportParams"), True
'frmDateSelect.Visible = True
'frmReportParams.Visible = True
'ViewReport.Visible = True
ShowRequiredParams ReportType.RequiredParams

End Sub

Private Sub ShowRequiredParams(RequiredParams As Dictionary)
Dim key As Variant
On Error Resume Next
    For Each key In RequiredParams.Keys
        fReports.frmReportParams.Controls("frm" & key).Visible = True
    Next key
On Error GoTo 0
End Sub


Private Sub TimeClockDetail_Click()
ResetForm
Dim iTimeClockDetail As New rclsTimeClockDetail
Set ReportType = iTimeClockDetail
DisplayCtrls GetCtrlDict(Me, "frmDateSelect", "ViewReport"), True
'frmDateSelect.Visible = True
'ViewReport.Visible = True
ShowRequiredParams ReportType.RequiredParams
End Sub

Private Sub UserForm_Initialize()
Me.Top = ActiveWindow.Top
Me.Left = ActiveWindow.Left
Me.Height = ActiveWindow.Height
Me.Width = ActiveWindow.Width
ResetForm

End Sub

Private Sub ResetForm()
DisplayCtrls GetCtrlDict(Me, "frmDateSelect", "ViewReport", "frmServerNum", "frmParam2", "frmParam3", "frmReportParams"), False
FromDate.value = ""
ToDate.value = ""
'frmDateSelect.Visible = False
'
'
'frmServerNum.Visible = False
'frmParam2.Visible = False
'frmParam3.Visible = False
'frmReportParams.Visible = False
'ViewReport.Visible = False
End Sub

Public Sub SetDateValue(tDate As Date)
DateBox.value = tDate
End Sub


Private Sub ViewReport_Click()
If ReportType Is Nothing Then
    Exit Sub
End If

ValidateRequiredParams ReportType.RequiredParams, GetSelectedParamDict
Dim iReport As New aclsReport
SetReportParams iReport

SetDataSource ReportType
iReport.RunReport ReportType
ResetForm
End Sub

Private Sub SetDataSource(ReportType As Variant)
ReportType.DataSource = GetArchiveDataObj
End Sub

Private Sub SetReportParams(iReport As aclsReport)
    If Not FromDate.value = "" Then
    iReport.StartDate = FromDate.value
    End If
    If Not ToDate.value = "" Then
    iReport.EndDate = ToDate.value
    End If
    iReport.ServerNum = CInt(GetFirstWord(ServerNum.value))
End Sub

Private Function GetSelectedParamDict() As Dictionary
Dim dict As New Dictionary
Dim ctrl As MSForms.Control
For Each ctrl In Me.frmReportParams.Controls
    If TypeName(ctrl) = "ComboBox" Then
        If Not ctrl.value = "" Then
            dict.Add ctrl.name, ctrl.value
        End If
    End If
Next ctrl
Set GetSelectedParamDict = dict
End Function
