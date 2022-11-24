Attribute VB_Name = "gui_Payment"
Option Explicit


Public Sub sizeshapes()

Dim i As Integer
For i = 19 To 24
Sheet7.Shapes("grpOpenChecksDisplay" & i).Top = Sheet7.Shapes("grpOpenChecksDisplay" & i - 18).Top
Sheet7.Shapes("grpOpenChecksDisplay" & i).Left = Sheet7.Shapes("grpOpenChecksDisplay1").Left + 475
Next i
End Sub

Public Sub MakeNamedRanges()



Dim i As Integer
Dim rg As range
For i = 7 To 11
Set rg = Sheet7.range("CheckSlot" & i)
rg.Offset(2, 0).name = "CheckSlot" & (i + 1)
Next i
End Sub
