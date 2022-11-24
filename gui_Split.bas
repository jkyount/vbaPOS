Attribute VB_Name = "gui_Split"
Option Explicit

Public Sub HighlightEntityGroup(Target As range, coll As Collection)
Dim i As Integer
Dim CollLine As Integer
Dim RangeLine As Integer
RangeLine = Target.row - GetRangeOffset
CollLine = GetCollRow(coll, RangeLine)
Target.Interior.color = RGB(100, 200, 200)
If Not CollLine = 1 Then
    i = 1
    Do While coll("Line" & CollLine - i).EntityGroup = coll("Line" & CollLine).EntityGroup
        Target.Offset(-i, 0).Interior.color = RGB(100, 200, 200)
        i = i + 1
        If CollLine - i = 0 Then Exit Do
    Loop
End If
If Not CollLine = coll.Count Then
     i = 1
    Do While coll("Line" & CollLine + i).EntityGroup = coll("Line" & CollLine).EntityGroup
        Target.Offset(i, 0).Interior.color = RGB(100, 200, 200)
        i = i + 1
        If CollLine + i > coll.Count Then Exit Do
    Loop
End If
End Sub
