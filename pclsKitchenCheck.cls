VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pclsKitchenCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function GetPrintColl(check As String) As Collection

Dim coll As Collection
Set coll = FormatPrintLines(check)

Dim PrintColl As New Collection
Dim row As Integer
Dim line As zclsCheckLines
row = 1
For Each line In coll
    If line.PrintKitchen = True Then
        PrintColl.Add line, "Line" & row
        row = row + 1
    End If
Next line
Set GetPrintColl = PrintColl
Set coll = Nothing
Set PrintColl = Nothing
Set line = Nothing
End Function

Private Function FormatPrintLines(check As String) As Collection
Dim coll As New Collection
Set coll = SortCheckLines(RecallCheckLines(check))

Dim line As zclsCheckLines
Dim iDetail As New zclsDailyCheckDetail
For Each line In coll
    line.Data = FilteredMatch(iDetail.Wrap(GetNewMatchObj("CheckNumber", check, "LocalGroup", line.LocalGroup)), "ItemName")
Next line
Set FormatPrintLines = coll
Set iDetail = Nothing
Set coll = Nothing
Set line = Nothing
End Function

Public Sub WritePrintLines(range As range, coll As Collection, row As Integer)


Dim ItemCount As Integer
ItemCount = 1
Dim i As Integer
For i = 1 To coll.Count
    Dim KitchenData As Variant
    KitchenData = coll(i).Data
    range.Rows(row).Cells(, 1).value = ""
    range.Rows(row).Cells(, 2).value = "    " & KitchenData(1)(0, 0) & " "
    If coll("Line" & i).primary = True Then
        range.Rows(row).Cells(, 1).NumberFormat = "@"
        range.Rows(row).Cells(, 1).value = "(" & CStr(ItemCount) & ")"
        range.Rows(row).Cells(, 2).value = KitchenData(1)(0, 0)
        ItemCount = ItemCount + 1
    End If
    row = row + 2
Next i

End Sub
