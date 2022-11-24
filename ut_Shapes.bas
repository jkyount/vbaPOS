Attribute VB_Name = "ut_Shapes"
Option Explicit

Public Sub Dontouch()

ActiveSheet.range("A1").value = ""

End Sub

Public Sub HideShowMe()
Dim bname As String
bname = Application.caller
HideShow bname
End Sub

Public Sub HideShow(shp As String)
Select Case ActiveSheet.Shapes(shp).Visible
    Case msoTrue
        ActiveSheet.Shapes(shp).Visible = msoFalse
    Case msoFalse
        ActiveSheet.Shapes(shp).Visible = msoTrue
End Select
End Sub



Public Sub HideShape(shp As String)
ActiveSheet.Shapes(shp).Visible = msoFalse
End Sub
Public Sub ShowShape(shp As String)
ActiveSheet.Shapes(shp).Visible = msoTrue
End Sub

Public Sub ClearShapeText(shp As String)
ActiveSheet.Shapes(shp).TextFrame.Characters.text = ""
End Sub

Public Sub SetShapeText(shp As String, text As String)
ActiveSheet.Shapes(shp).TextFrame.Characters.text = text
End Sub

Public Sub SetShapeTrans(shp As String, FillVal As Double, Optional LineVal As Double = -1)
ActiveSheet.Shapes(shp).Fill.Transparency = FillVal
If Not LineVal = -1 Then ActiveSheet.Shapes(shp).line.Transparency = LineVal
End Sub

Public Sub SetShapeLineWidth(shp As String, value As Double)
ActiveSheet.Shapes(shp).line.weight = value
End Sub

Public Sub GroupShapes()

Dim i As Integer
Dim grp As Shape
For i = 1 To 24
Set grp = Sheet7.Shapes.range(Array("lblOpenChecksName" & i, "lblOpenChecksTime" & i, "lblOpenChecksTotal" & i, "btnOpenChecks" & i, "lblOpenChecksServer" & i)).Group
grp.name = ("grpOpenChecks" & i)
Next i

End Sub

Public Sub RenameShape(sheet As Worksheet, OldName As String, NewName As String)
sheet.Shapes(OldName).name = NewName
End Sub


Public Function GetFirstWord(str As String) As String
If str = "" Then
    GetFirstWord = "0"
    Exit Function
    
End If
Dim arr As Variant
arr = Split(str, " ")
GetFirstWord = arr(0)

End Function

Public Sub Border(cell As range, weight As Integer, CellsUp As Integer, CellsAcross As Integer)
Dim rg As range
Dim sheet As String
sheet = cell.Parent.name
Set rg = Worksheets(sheet).range(cell, cell.Offset(CellsUp, CellsAcross))
rg.BorderAround xlSolid, weight
End Sub

Public Function InValues(val As Integer, Vals() As Variant) As Boolean
Dim i As Integer
For i = 0 To UBound(Vals)
    If val = Vals(i) Then
        InValues = True
        Exit Function
    End If
Next i
End Function

Public Function LongToRgb(value As Long) As XlRgbColor
LongToRgb = RGB(value Mod 256, Int(value / 256) Mod 256, Int(value / 256 / 256) Mod 256)
End Function


