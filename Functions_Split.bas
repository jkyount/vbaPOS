Attribute VB_Name = "Functions_Split"
Option Explicit

Dim ClickCount As Integer
Dim collOriginal As New Collection
Dim collSplit As New Collection
Dim ClickRange As range


Public Sub SetClickRange(rg As range)
Set ClickRange = rg
End Sub

Public Function GetClickRange() As range
Set GetClickRange = ClickRange
End Function

Public Sub SetClickCount(cnt As Integer)
ClickCount = cnt
End Sub

Public Function GetClickCount() As Integer
GetClickCount = ClickCount
End Function

Public Function GetcollOriginal()
Set GetcollOriginal = collOriginal
End Function

Public Function GetcollSplit()
Set GetcollSplit = collSplit
End Function

Public Function GetCollRow(coll As Collection, activerow As Integer) As Integer
Dim z As zclsCheckLines
For Each z In coll
    If z.GuiRow = activerow Then
        GetCollRow = z.row
        Set z = Nothing
        Exit Function
    End If
Next z
GetCollRow = -1
Set z = Nothing
End Function

Public Sub ActivateSplitCheckScreen()
Dim iState As New aclsState
Init
Sheet10.Activate
'Sheet10.ScrollArea = "A1:T40"
InitializeGUI
SetShapeText "SplitIndicator", "Splitting Check: " & currentcheck & "."
SetShapeText "OrderName", iState.OrderName
SetShapeText "ServerName", ThisEmployee.FirstName
Sheet10.range("OrderName").value = iState.OrderName
Sheet10.range("NewOrderName").value = ThisOrder.OrderType.GetSplitOrderName
Sheet10.range("NewCheckNumber").value = PeekNextCheck
End Sub

Private Sub Init()
PopGuiWithCheckAttributes ThisOrder.ValueDict, Sheet10
Set collOriginal = DuplicateCheckLines(collCheckData)
ClearCollection collSplit
CopyToTemp currentcheck
ClickCount = 0
Set ClickRange = Nothing
End Sub

Private Sub InitializeGUI()
HideAllSeatButtons
PositSeatButtons "Original", collOriginal
WriteCheckLines Sheet10.range("OriginalCheckRange"), collOriginal
Sheet10.range("SplitCheckRange").value = ""
End Sub

Private Sub UpdateGUI()
InitializeGUI
WriteCheckLines Sheet10.range("SplitCheckRange"), collSplit
PositSeatButtons "Split", collSplit
End Sub

Private Sub PositSeatButtons(opt As String, coll As Collection)
Dim i As Integer
Dim seat As Integer
Dim rg As range
Set rg = Sheet10.range(opt & "CheckRange")
If Not coll.Count = 0 Then
    seat = coll(1).seat
    ShowShape opt & seat
    Sheet10.Shapes(opt & seat).Top = rg.Rows(1).Top
    i = 2
    Do While Not i > coll.Count
        Do Until coll(i).seat <> seat
            If i = coll.Count Then Exit Do
            i = i + 1
        Loop
        seat = coll(i).seat
        ShowShape opt & seat
        Sheet10.Shapes(opt & seat).Top = rg.Rows(coll(i).GuiRow - 1).Top
        i = i + 1
    Loop
End If
End Sub

Private Sub HideAllSeatButtons()
Dim i As Integer
For i = 1 To 12
    Sheet10.Shapes("Split" & i).Visible = msoFalse
    Sheet10.Shapes("Original" & i).Visible = msoFalse
Next i
End Sub

Public Sub SplitItem(EntityGroup As Integer, DestinationCheck As String)
TransferItem EntityGroup, DestinationCheck
UpdateCheckCollections
UpdateGUI
End Sub

Public Sub UpdateCheckCollections()
Dim iTemp As New zclsTemp
Dim qry As String
Dim check As String
check = currentcheck
qry = "SELECT * FROM TempCheck WHERE CheckNumber = """ & check & """ ORDER BY LocalGroup ASC"
Set collOriginal = SortCheckLines(GetCheckLines(currentcheck, GetValueDict(iTemp.Wrap(GetNewMatchObj(, check)), qry)))
check = Sheet10.range("NewCheckNumber").value
qry = "SELECT * FROM TempCheck WHERE CheckNumber = """ & check & """ ORDER BY LocalGroup ASC"
Set collSplit = SortCheckLines(GetCheckLines(Sheet10.range("NewCheckNumber").value, GetValueDict(iTemp.Wrap(GetNewMatchObj(, check)), qry)))
Set iTemp = Nothing
End Sub

Private Sub TransferItem(EntityGroup As Integer, DestinationCheck As String)
Dim iTemp As New zclsTemp
Update iTemp.Wrap(GetNewUpdateObj("EntityGroup", EntityGroup, "CheckNumber", DestinationCheck))
End Sub

Public Sub ExecuteSplit()
pExecuteSplit
PrintGuestCheck currentcheck
Set collSplit = Nothing
Set collOriginal = Nothing
ActivateHomeScreen
End Sub

Public Sub SplitAgain()
Application.ScreenUpdating = False
pExecuteSplit
PrintGuestCheck currentcheck
RestoreState
Set collCheckData = DuplicateCollection(collOriginal)
ActivateSplitCheckScreen
Application.ScreenUpdating = True
End Sub

Public Sub SplitSeat()
Dim bname As String, DestinationCheck As String
Dim ClickedSeat As Shape
Dim SeatDict As New Dictionary
bname = Application.caller
Set ClickedSeat = Sheet10.Shapes(bname)
If bname Like "Original*" Then
    Set SeatDict = GenerateSeatDict(collOriginal)
    DestinationCheck = PeekNextCheck
End If
If bname Like "Split*" Then
    Set SeatDict = GenerateSeatDict(collSplit)
    DestinationCheck = currentcheck
End If

Dim SeatItems As New Collection
Set SeatItems = SeatDict(ClickedSeat.TextFrame.Characters.text)
Dim CurrentEntityGroup As Integer, i As Integer

CurrentEntityGroup = SeatItems(1).EntityGroup
    TransferItem CurrentEntityGroup, DestinationCheck
    i = 2
    For i = 2 To SeatItems.Count
        If Not SeatItems(i).EntityGroup = CurrentEntityGroup Then
            TransferItem SeatItems(i).EntityGroup, DestinationCheck
            CurrentEntityGroup = SeatItems(i).EntityGroup
        End If
    Next i
UpdateCheckCollections
UpdateGUI

End Sub

Public Sub SplitAllSeats()
If collSplit.Count > 0 Then
    MsgBox "You have already performed split actions on this check.  To split all seats, first cancel or complete the current operation."
    Exit Sub
End If
Dim SeatDict As New Dictionary
Set SeatDict = GenerateSeatDict(collOriginal)
Application.ScreenUpdating = False
Dim NewCheck As String
Dim CurrentEntityGroup As Integer
Dim line As zclsCheckLines
Dim i As Integer
Dim j As Integer
For j = 1 To (SeatDict.Count - 1) '<--------SeatDict count is base 0.  Loop begins on second element of SeatDict.
    NewCheck = PeekNextCheck
    CurrentEntityGroup = SeatDict.Items(j)(1).EntityGroup
    TransferItem CurrentEntityGroup, NewCheck
    i = 2
    For i = 2 To SeatDict.Items(j).Count
        If Not SeatDict.Items(j)(i).EntityGroup = CurrentEntityGroup Then
            TransferItem SeatDict.Items(j)(i).EntityGroup, NewCheck
            CurrentEntityGroup = SeatDict.Items(j)(i).EntityGroup
        End If
    Next i
    CreateNewCheck
Next j
UpdateOriginalCheck
ActivateHomeScreen
Application.ScreenUpdating = True
End Sub

Private Sub pExecuteSplit()
CreateNewCheck
UpdateOriginalCheck
End Sub


Private Sub CreateNewCheck()
Dim NewCheck As String
NewCheck = GetNextCheck
Dim NewOrder As New aclsOrder
Dim NewOrderName As String


'::::::: Create Record in DailyCheckIndex for the new check number
InitializeCheck NewCheck

'::::::: Get aclsOrder object with new check number assigned, and same orer type as original order
Set NewOrder = NewOrder.NewOrderObject(ThisOrder.SameOrderType, NewCheck) 'Set NewOrder = NewOrder.CreateNewOrder(ThisOrder.SameOrderType, NewCheck)



'::::::: Assign aclsOrder object the same order info as original order
NewOrder.OrderDetails = ThisOrder.SameOrderInfo

'::::::: Modify the only piece of order info that needs to be different, "OrderName," via a method determined by the order type.
ReplaceDictValue NewOrder.OrderDetails, "OrderName", NewOrder.OrderType.GetSplitOrderName

'::::::: Assign check to table, if DineIn.  Do nothing if Carryout
NewOrder.OrderType.CreateNewCheck NewCheck

UpdateDailyCheckIndex NewCheck, NewOrder.OrderDetails
SubmitChanges NewCheck
DailyCheckIndex_CalculateTotals NewCheck
PrintGuestCheck NewCheck
End Sub
Private Sub UpdateOriginalCheck()
SubmitChanges currentcheck
DailyCheckIndex_CalculateTotals currentcheck
End Sub
Private Sub SubmitChanges(check As String)
Dim iDetail As New zclsDailyCheckDetail
iDetail.SubmitChanges check
End Sub

Public Sub CancelSplit()
Set collOriginal = Nothing
Set collSplit = Nothing
RecallCheck ThisOrder, Sheet1
ActivateOrderScreen
End Sub




