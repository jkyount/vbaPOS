VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fConfig 
   Caption         =   "Configuration"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   OleObjectBlob   =   "fConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Employees_Click()
Me.Hide
fEmployees.Show

End Sub



Private Sub Families_Click()
Me.Hide
fFamilies.Show


End Sub

Private Sub MenuItems_Click()
Me.Hide
fMenuItems.Show

End Sub

Private Sub MultiMenus_Click()
Me.Hide
fMultiMenu.Show

End Sub

Private Sub PastReports_Click()
Me.Hide
fReports.Show

End Sub

Private Sub Reports_Click()
Me.Hide
fReports_CurrentShift.Show

End Sub

Private Sub UserForm_Activate()
Me.Top = ActiveWindow.Top
Me.Left = ActiveWindow.Left
Me.Height = ActiveWindow.Height
Me.Width = ActiveWindow.Width
Me.MainFrame.Top = Me.Top + 44
Me.MainFrame.Left = Me.Left + 44
Me.MainFrame.Height = Me.Height - 110
Me.MainFrame.Width = Me.Width - 88
Me.InnerFrame.Top = (Me.MainFrame.Height / 2) - (Me.InnerFrame.Height / 2)
Me.InnerFrame.Left = (Me.MainFrame.Width / 2) - (Me.InnerFrame.Width / 2)



End Sub

