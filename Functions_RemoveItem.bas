Attribute VB_Name = "Functions_RemoveItem"
Option Explicit

Public Sub DeleteItem()
Dim EntityGroup As Integer
EntityGroup = GetEntityGroup(ActiveCell.row, collCheckData)
RemoveItem currentcheck, EntityGroup
Set collCheckData = RecallCheckLines(currentcheck)
WriteCheckLines Sheet1.range("CheckRange"), collCheckData
End Sub


Public Function GetEntityGroup(SelectedLine As Integer, coll As Collection) As Integer

Dim iDetail As New zclsDailyCheckDetail
Dim LocalGroup As Integer, GuiRow As Integer
GuiRow = GetGuiRow(SelectedLine, GetRangeOffset)
LocalGroup = GetLocalGroup(GuiRow, coll)
GetEntityGroup = ValueMatch(iDetail.Wrap(GetNewMatchObj(, currentcheck, "LocalGroup", LocalGroup)), "EntityGroup")
End Function
Public Function GetGuiRow(SelectedLine As Integer, RangeOffset As Integer) As Integer
GetGuiRow = SelectedLine - RangeOffset
End Function
Public Function GetRangeOffset() As Integer
Select Case ActiveCell.Parent.CodeName
    Case "Sheet1"
        GetRangeOffset = (Sheet1.range("CheckRange").Rows(1).row - 1)
        Exit Function
    Case "Sheet8"
        GetRangeOffset = (Sheet8.range("CheckRange").Rows(1).row - 1)
        Exit Function
    Case "Sheet10"
        GetRangeOffset = (Sheet10.range("OriginalCheckRange").Rows(1).row - 1)
        Exit Function
    Case "Sheet14"
        GetRangeOffset = (Sheet14.range("Seat1").Rows(1).row - 1)
        Exit Function
End Select
End Function

Private Function GetLocalGroup(GuiRow As Integer, coll As Collection) As Integer
Dim member As Variant
For Each member In coll
    If member.GuiRow = GuiRow Then
        GetLocalGroup = member.LocalGroup
        Exit Function
    End If
Next member
End Function







