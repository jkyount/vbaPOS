VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fInitialize 
   Caption         =   "UserForm1"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15825
   OleObjectBlob   =   "fInitialize.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fInitialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Activate()
Initialize
Me.Hide
End Sub


Private Sub Initialize()

Dim i As Long
For i = 1 To 2500
Debug.Print i
Next i
End Sub




