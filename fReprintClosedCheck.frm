VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fReprintClosedCheck 
   Caption         =   "Reprint Closed Check"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   OleObjectBlob   =   "fReprintClosedCheck.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fReprintClosedCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PrintCheck_Click()
PrintGuestCheck GetFirstWord(ClosedChecks.value)
Me.Hide
Unload Me
End Sub

Private Sub TransferCheck_Click()
CheckTransfer GetFirstWord(ClosedChecks.value), ThisEmployee.ServerNum
MsgBox "Check transferred."
IndicateTablesInUse
Me.Hide
Unload Me
End Sub

Public Sub ShowForReprintCheck()
Me.Caption = "Reprint Closed Check"
RestoreState
PopDropDown ClosedChecks, GetCheckArray(GetNewMatchObj("ServerNum", ThisEmployee.ServerNum, "Closed", True))
Me.PrintCheck.Visible = True
Me.TransferCheck.Visible = False
Me.Show
End Sub

Private Function GetCheckArray(iDataObj As aclsDataObject) As Variant
Dim arr As Variant
arr = GetChecks(iDataObj.Field1, iDataObj.Value1, iDataObj.Field2, iDataObj.Value2)

Dim TempArray() As Variant
Dim junkarray(0 To 0, 0 To 0) As Variant
ReDim TempArray(1 To UBound(arr) + 1)
Dim i As Integer
For i = 0 To UBound(arr)
    junkarray(0, 0) = arr(i, 0) & " [" & arr(i, 1) & "]  [" & arr(i, 2) & "]  [" & Format(arr(i, 3), "$0.00") & "]  [" & arr(i, 4) & "]"
    TempArray(i + 1) = junkarray
Next i
GetCheckArray = TempArray

End Function

Public Sub ShowForTransferCheck()

PopDropDown ClosedChecks, GetCheckArray(GetNewMatchObj("NOT ServerNum", 0, "Closed", False))
Me.Caption = "Transfer Check"
Me.PrintCheck.Visible = False
Me.TransferCheck.Visible = True
Me.Show
End Sub


