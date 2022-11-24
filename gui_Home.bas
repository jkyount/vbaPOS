Attribute VB_Name = "gui_Home"

Option Explicit



Public Sub IndicateTablesInUse()
Application.ScreenUpdating = False
Dim TablesInUse As New Collection
Set TablesInUse = GetTablesInUse
Dim ServerTables As New Collection
Set ServerTables = GetServerTables(ThisEmployee.ServerNum)
Dim member As Variant
Dim i As Integer
For i = 2 To 13
Sheet7.Shapes("InUseTable" & i).Visible = msoFalse
Sheet7.Shapes("VisTable" & i).Fill.ForeColor.RGB = RGB(255, 195, 87)
Next i
For Each member In TablesInUse
    Sheet7.Shapes("InUse" & member).Visible = msoTrue
Next member
For Each member In ServerTables
        Sheet7.Shapes("Vis" & member).Fill.ForeColor.RGB = ThisEmployee.AccentColor
Next member
'Application.ScreenUpdating = True
End Sub


Public Sub DisplayChecks(arrChecks As Variant)
Application.ScreenUpdating = False
HideShape "grpFloor"
HideShape "BeginNewCheck"

Dim i As Integer
Dim k As Integer

Dim rg As range
Dim CheckSlotRange As range
i = 0
Do Until i > UBound(arrChecks, 1) Or i + 1 > 12
    Set rg = Sheet7.range("CheckSlot" & i + 1)
    rg.value = arrChecks(i, 0)
    rg.Interior.color = 3243501
    rg.BorderAround 1, -4138
    rg.Offset(0, 2).value = arrChecks(i, 1)
    rg.Offset(0, 2).BorderAround 1, -4138
     rg.Offset(0, 2).Interior.color = 15921906
    rg.Offset(0, 4).value = arrChecks(i, 2)
    rg.Offset(0, 4).BorderAround 1, -4138
    rg.Offset(0, 4).Interior.color = 15921906
    rg.Offset(0, 6).value = arrChecks(i, 3)
    rg.Offset(0, 6).BorderAround 1, -4138
    rg.Offset(0, 6).Interior.color = 15921906
    rg.Offset(0, 8).value = arrChecks(i, 4)
    rg.Offset(0, 8).BorderAround 1, -4138
     rg.Offset(0, 8).Interior.color = 15921906
    ShowShape "" & i + 1
    i = i + 1
Loop

If i > UBound(arrChecks, 1) Then
    Sheet7.Shapes("BeginNewCheck").Top = Sheet7.Shapes("" & i + 1).Top
    Sheet7.Shapes("BeginNewCheck").Left = Sheet7.Shapes("" & i + 1).Left
    Do Until i + 1 > 12
        Set rg = Sheet7.range("CheckSlot" & i + 1)
        Set CheckSlotRange = Sheet7.range(rg, rg.Offset(0, 8))
        CheckSlotRange.value = ""
        CheckSlotRange.Borders.LineStyle = xlNone
        CheckSlotRange.Interior.color = 15189684
        HideShape "" & i + 1
        i = i + 1
    Loop
End If


'If UBound(arrChecks, 1) + 1 > 12 Then ShowShape "NextChecks"
    
'Application.ScreenUpdating = True
End Sub

Public Sub DisplayOpenChecks()

DisplayChecks GetChecks("Closed", False)

End Sub


