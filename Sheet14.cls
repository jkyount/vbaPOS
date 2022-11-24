VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Private Sub Worksheet_Activate()

Sheet14.ScrollArea = "A1:AC25"
End Sub

Public Sub Worksheet_SelectionChange(ByVal Target As range)

Dim i As Integer, j As Integer, ClickedRange As Integer
Dim NewSeat As Integer, line As Integer
Dim SeatRanges As Collection
Set SeatRanges = GetSeatRanges
Dim dict As Dictionary
Set dict = GetSeatDict



If EditSeat_GetClickCount = 0 Then
    Dim rg As range
    For Each rg In SeatRanges
        ClickedRange = ClickedRange + 1
        If Not Intersect(rg, Target) Is Nothing Then
            SetOriginalSeat ClickedRange
            Exit For
        End If
    Next rg
    If dict.Exists(CStr(ClickedRange)) = False Then
        Exit Sub
    End If
    If (Target.row - 5) > dict(CStr(ClickedRange)).Count Then
        EditSeat_SetClickCount 0
        Sheet14.range("AllSeatRanges").Interior.color = RGB(255, 255, 255)
        Exit Sub
    End If
    If Intersect(Sheet14.range("AllSeatRanges"), Target) Is Nothing Then Exit Sub
    EditSeat_SetClickCount 1
    SetFirstClickedCell Target
    Sheet14.range("AllSeatRanges").Interior.color = RGB(255, 255, 255)
    HighlightEntityGroup Target, dict(CStr(ClickedRange))
    Exit Sub
End If
If EditSeat_GetClickCount > 0 Then
    If Intersect(Sheet14.range("AllSeatRanges"), Target) Is Nothing Then
        EditSeat_SetClickCount 0
        Sheet14.range("AllSeatRanges").Interior.color = RGB(255, 255, 255)
        Exit Sub
    End If
    For Each rg In SeatRanges
        ClickedRange = ClickedRange + 1
        If Not Intersect(rg, Target) Is Nothing _
        And Not Intersect(rg, GetFirstClickedCell) Is Nothing _
        Then
            If (Target.row - 5) > dict(CStr(ClickedRange)).Count Then
                EditSeat_SetClickCount 0
                Sheet14.range("AllSeatRanges").Interior.color = RGB(255, 255, 255)
                Exit Sub
            End If
            EditSeat_SetClickCount 1
            SetFirstClickedCell Target
            Sheet14.range("AllSeatRanges").Interior.color = RGB(255, 255, 255)
            HighlightEntityGroup Target, dict(CStr(ClickedRange))
            Exit Sub
        End If
    Next rg
    For i = 1 To 12
        If Not Intersect(SeatRanges(i), Target) Is Nothing Then
            NewSeat = i
            line = GetFirstClickedCell.row
            Debug.Print GetEntityGroup(line, dict(CStr(GetOriginalSeat)))
            
            EditSeat GetEntityGroup(line, dict(CStr(GetOriginalSeat))), NewSeat
            Sheet14.range("AllSeatRanges").Interior.color = RGB(255, 255, 255)
            EditSeat_SetClickCount 0
            Exit Sub
        End If
    Next i
End If
End Sub
