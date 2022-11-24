Attribute VB_Name = "Functions_EditSeat"
Option Explicit
Dim OriginalSeat As Integer
Dim FirstClickedCell As range
Dim ClickCount As Integer
Dim SeatDict As Dictionary
Dim collEditSeatData As New Collection



Public Sub ActivateEditSeatScreen()
Sheet14.Activate
Init
SyncValues
UpdateGUI
UpdateRanges
End Sub


Private Sub Init()
Set collEditSeatData = DefineWriteLines(DuplicateCheckLines(collCheckData), 1, 2)
CopyToTemp currentcheck
ClickCount = 0
Set FirstClickedCell = Nothing
Set SeatDict = GenerateSeatDict(collEditSeatData)
End Sub

Private Sub UpdateGUI()
PopGuiWithCheckAttributes ThisOrder.ValueDict, Sheet14
End Sub

Private Sub UpdateRanges()
Sheet14.range("AllSeatRanges").value = ""
WriteCheckLinesBySeat SeatDict
End Sub

Public Function GetSeatDict() As Dictionary
If SeatDict.Count = 0 Then
    Set GetSeatDict = GenerateSeatDict(collEditSeatData)
    Exit Function
End If
Set GetSeatDict = SeatDict
End Function
Public Function GenerateSeatDict(coll As Collection) As Dictionary
Dim i As Integer, k As Integer
Dim dict As Dictionary
Dim seat As Collection
If Not coll.Count = 0 Then
    i = 2
    Set seat = New Collection
    Set dict = New Dictionary
    seat.Add coll(1), "Line1"
    '----\/\/\/----THIS IF STATEMENT WAS COMMENTED OUT ----'
        If Not i > coll.Count Then
        Do While coll(i).seat = coll(i - 1).seat
            seat.Add coll(i), "Line" & i
            i = i + 1
            If i > coll.Count Then Exit Do
        Loop
        
        
        End If
        dict.Add CStr(coll(i - 1).seat), SortCheckLines(seat)
    Do While Not i > coll.Count
    Set seat = New Collection
        seat.Add coll(i), "Line1"
            i = i + 1
            k = 2
        If Not i > coll.Count Then
        Do While coll(i).seat = coll(i - 1).seat
            seat.Add coll(i), "Line" & k
            i = i + 1
            k = k + 1
            If i > coll.Count Then Exit Do
        Loop
        End If
        dict.Add CStr(coll(i - 1).seat), SortCheckLines(seat)
    Loop
End If
   Set GenerateSeatDict = dict
End Function


Private Sub WriteCheckLinesBySeat(SeatDict As Dictionary)
Sheet14.range("AllSeatRanges").value = ""
Dim i As Integer
Dim SeatNumbers As Variant
SeatNumbers = SeatDict.Keys
For i = 1 To SeatDict.Count
    WriteCheckLines Sheet14.range("Seat" & SeatNumbers(i - 1)), SeatDict(SeatNumbers(i - 1))
Next i
End Sub

Public Function GetSeatRanges() As Collection
Dim SeatRanges As New Collection
Dim rg As range
Dim i As Integer
For i = 1 To 12
    Set rg = Sheet14.range("Seat" & i)
    SeatRanges.Add rg, CStr(i)
Next i
Set GetSeatRanges = SeatRanges
End Function

Public Sub SetOriginalSeat(seat As Integer)
OriginalSeat = seat
End Sub

Public Function GetOriginalSeat() As Integer
GetOriginalSeat = OriginalSeat
End Function

Public Sub SetFirstClickedCell(Target As range)
Set FirstClickedCell = Target
End Sub

Public Function GetFirstClickedCell() As range
Set GetFirstClickedCell = FirstClickedCell
End Function

Public Sub EditSeat_SetClickCount(value As Integer)
ClickCount = value
End Sub

Public Function EditSeat_GetClickCount() As Integer
EditSeat_GetClickCount = ClickCount
End Function

Public Sub UpdateSeatDict(dict As Dictionary)
Set SeatDict = dict
End Sub

Public Sub EditSeat(EntityGroup As Integer, NewSeat As Integer)
TransferSeat EntityGroup, NewSeat
UpdateCheckCollection
UpdateSeatDict GenerateSeatDict(collEditSeatData)
WriteCheckLinesBySeat SeatDict
End Sub

Private Sub UpdateCheckCollection()
Dim iTemp As New zclsTemp
Set collEditSeatData = SortCheckLines(GetCheckLines(currentcheck, GetValueDict(iTemp.Wrap(GetNewMatchObj(, currentcheck)))))
Set iTemp = Nothing
End Sub

Private Sub TransferSeat(EntityGroup As Integer, NewSeat As Integer)
Dim iTemp As New zclsTemp
Update iTemp.Wrap(GetNewUpdateObj("EntityGroup", EntityGroup, "Seat", NewSeat))
End Sub

Public Sub ExecuteSeatEdit()
pExecuteSeatEdit
Set SeatDict = Nothing
Set collEditSeatData = Nothing
RecallCheck ThisOrder, Sheet1
ActivateOrderScreen
End Sub
Private Sub pExecuteSeatEdit()
SubmitChanges currentcheck
End Sub



