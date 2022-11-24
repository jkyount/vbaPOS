VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RepViewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   OleObjectBlob   =   "RepViewer.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "RepViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pReportType As Variant

Private Sub PrintReport_Click()

Dim iReport As New aclsReport
iReport.PrintReport pReportType
End Sub

Public Sub ShowRepViewer(ReportType As Variant)
Set pReportType = ReportType
RepViewer.Show
Set pReportType = Nothing
End Sub

Private Sub UserForm_Activate()
Me.Height = ActiveWindow.Height
Me.Width = ActiveWindow.Width
Me.Top = ActiveWindow.Top
Me.Left = ActiveWindow.Left

End Sub


