Attribute VB_Name = "Functions_Order"
Option Explicit
Dim Qty As Integer
Dim currentseat As Integer
Dim TestColl As Collection
Dim ClickFlag As Boolean

Public Sub SetCurrentParentToNothing()
CurrentParent = 0
End Sub
Public Function GetCurrentParent() As Integer
GetCurrentParent = CurrentParent
End Function

Public Function GetCurrentSeat() As Integer
If currentseat = 0 Then GetCurrentSeat = 1
GetCurrentSeat = currentseat
End Function
Public Sub SetCurrentSeat(val As Integer)
currentseat = val
End Sub

Public Sub ActivateOrderScreen()
Sheet1.Activate
ResetOrderState
SeatOne
CaptureState
End Sub

Public Sub Ssfub()
If ThisEmployee.IDNumber = 0 Then RestoreState
Dim bname As String
bname = Application.caller
Dim ItemID As Integer
ItemID = CInt(bname)
ItemDirector ItemID
End Sub

Public Sub SetCurrentParent(ID As Integer)
CurrentParent = ID
ShowParentFrame GetItemByID(ID).ItemName
End Sub
Public Sub ItemDirector(Optional ItemID As Integer)

If ClickFlag = True Then
    Exit Sub
End If
ClickFlag = True

If ItemID = 0 Then ItemID = CInt(Application.caller)
Set ThisItem = New aclsItem
ThisItem.Initialize ItemID

'/////ItemID becomes INTEGER here////'

Order
Set ThisItem = Nothing
ClickFlag = False
End Sub

Public Sub Order()
ShowShape "ComponentBLOCK"

ThisItem.OrderRank = ThisItem.AssignOrderRank
ThisItem.CollID = GetNextCollID(OrderByCollID(DuplicateCItem(CItem)))
ThisItem.RequiredComponents = ThisItem.GetRequiredComponents

ThisItem.OrderRank.SetParent
ThisItem.Children = GetNewChildren
ThisItem.ItemType.SpclConfig ThisItem
ThisItem.OrderRank.BuildItemCollection

ThisItem.ItemType.zclsItem_RefreshPreviewWindow
ThisItem.ItemType.zclsItem_UpdateGUI
ThisItem.OrderRank.CheckForRequiredComponents ThisItem
End Sub

Public Function PopItem(ItemID As Integer, item As Variant) As aclsItem
Dim iMenu As New zclsMenu
Dim ItemDict As New Dictionary
Set ItemDict = GetValueDict(iMenu.Wrap(GetNewMatchObj("ID", ItemID)))(1)
item.ItemID = ItemID
item.ItemName = ItemDict("ItemName")
item.Price = ItemDict("Price")
item.Family = ItemDict("Family")
If Not IsNull(ItemDict("Req1")) Then
    item.Req1 = ItemDict("Req1")
    item.Req2 = ItemDict("Req2")
End If
If item.Req1 = "" Then
    item.Req1 = item.Req2
    item.Req2 = ""
End If
item.IsPrimaryItem = ItemDict("IsPrimaryItem")
item.Category = ItemDict("Category")
item.CoolerFlag = ItemDict("Flag")
item.AlwaysTax = ItemDict("AlwaysTax")
item.AltScePrice = ItemDict("AltScePrice")
item.AltPrice = ItemDict("AltPrice")
item.DiscountPrice = ItemDict("DiscountPrice")
item.PrintKitchen = ItemDict("PrintKitchen")
item.PrintPantry = ItemDict("PrintPantry")
item.ItemOptions = ItemDict("ItemOptions")

item.seat = GetCurrentSeat
Set PopItem = item

Set iMenu = Nothing
Set ItemDict = Nothing
End Function

Public Sub GetItemCell()
If Sheet1.range("CheckOrientationCell").value = "" Then
    Set ItemCell = Sheet1.range("CheckOrientationCell")
    Exit Sub
End If
Set ItemCell = Sheet1.range("CheckOrientationCell").Offset(100, 0).End(xlUp).Offset(1, 0)
End Sub

Public Sub AdvCell()
Set ItemCell = ItemCell.Offset(1, 0)
End Sub
Public Sub AdvRange(rg As range)
Set rg = rg.Offset(1, 0)
End Sub

Public Sub SetEndOfCheck()
EndOfCheck = ItemCell.row
End Sub

Public Sub NextSeat()
currentseat = currentseat + 1
Sheet1.Shapes("SeatIndicator").TextFrame.Characters.text = "Entry for seat " & currentseat & ""
End Sub
Public Sub SeatOne()
currentseat = 1
Sheet1.Shapes("SeatIndicator").TextFrame.Characters.text = "Entry for seat " & currentseat & ""
End Sub
Public Sub WholeTopping()
CItem("1").ItemType.ToppingArea = 0
Sheet1.Shapes("PzaWhole").line.weight = 3
Sheet1.Shapes("PzaHalf").line.weight = 1
End Sub
Public Sub HalfTopping()
CItem("1").ItemType.ToppingArea = 1
Sheet1.Shapes("PzaWhole").line.weight = 1
Sheet1.Shapes("PzaHalf").line.weight = 3
End Sub
Public Sub SetQuantity()
Dim QtyValidate As Variant
QtyValidate = InputBox("Set quantity for next item Ordered:", , "1")
If IsNumeric(QtyValidate) Then
    If CInt(QtyValidate) > 0 Then
        Qty = QtyValidate
        Exit Sub
    End If
End If
MsgBox "Please enter a positive numeric value."

End Sub

Public Sub SetToDineIn(check As String)
Dim iIndex As New zclsDailyCheckIndex
Update iIndex.Wrap(GetNewUpdateObj(, check, "DineIn", True))
SetOrderTypeIndicator "DineIn"
UpdateTax currentcheck
Set iIndex = Nothing
End Sub
Public Sub SetToCarryout(check As String)
Dim iIndex As New zclsDailyCheckIndex
Update iIndex.Wrap(GetNewUpdateObj(, check, "DineIn", False))
SetOrderTypeIndicator "Carryout"
Set iIndex = Nothing
UpdateTax currentcheck
End Sub

Public Sub CheckQueueForRequiredComponents(coll As Collection)
Dim i As Integer
For i = coll.Count To 1 Step -1
    Dim Reqs As Variant
    Reqs = AllRequirementsFullfilled(coll(i))
    If Not Reqs = True Then
        MsgBox ("Required component " & Reqs & " not selected.")
        SetCurrentParent coll(i).CollID
        Exit Sub
    End If
Next i

Dim FormattedColl As Collection
Set FormattedColl = FormatItemCollection(coll)
'NormalizePrintParameters coll
'FormatItemCollection coll
OrderQueuedItems FormattedColl
ResetOrderState
End Sub

Public Function AllRequirementsFullfilled(item As aclsItem) As Variant
If Not item.RequiredComponents.Count = 0 Then
Dim key As Variant
    For Each key In item.RequiredComponents.Keys
        Dim RequiredComponent As String
        RequiredComponent = item.RequiredComponents(key)
        If item.ValidateRequirements(item, RequiredComponent) = False Then
            ShowComponentFrame RequiredComponent
            DisplayQuickMods RequiredComponent
            ShowParentFrame item.ItemName
            AllRequirementsFullfilled = RequiredComponent
            Exit Function
        End If
    Next key
End If
AllRequirementsFullfilled = True
End Function




Public Function RequiresComponents(item As Variant) As Boolean
If item.Req1 = "" Then
    RequiresComponents = False
    Exit Function
End If
RequiresComponents = True
End Function

Public Function FormatSides(coll As Collection) As Collection

Dim member As aclsItem
Dim child As bclsChild
For Each member In coll
    If Not member.Children.coll.Count = 0 Then
    For Each child In member.Children.coll
        Dim ChildItem As aclsItem
        Set ChildItem = child.item
        If ChildItem.Family = "Drsng" Or ChildItem.Family = "Sce" Then
            member.ItemName = member.ItemName & "  /  " & LTrim(ChildItem.ItemName)
            member.Price = (member.Price) + (ChildItem.Price)
            coll.Remove CStr(child.ID)
        End If
    Next child
    End If
Next member
Set member = Nothing
Set child = Nothing
Set ChildItem = Nothing
Set FormatSides = coll
End Function

Public Sub CLICK_Sheet1_btnDONE()
CheckQueueForRequiredComponents CItem
End Sub
Public Sub CancelOrder()
RestoreState
Dim CheckDetail As New zclsDailyCheckDetail
DeleteMatch CheckDetail.Wrap(GetNewMatchObj("CheckNumber", currentcheck, "Sent", False))
If Match(CheckDetail.Wrap(GetNewMatchObj("CheckNumber", currentcheck, "Sent", True))) = False Then
    ThisOrder.OrderType.CancelOrder
    
    SetCheckUnused currentcheck
    Dim CheckIndex As New zclsDailyCheckIndex
    DeleteMatch CheckIndex.Wrap(GetNewMatchObj("CheckNumber", currentcheck))
End If
Set CheckDetail = Nothing
Set CheckIndex = Nothing
'ResetOrderState
ActivateHomeScreen
End Sub
Public Sub OrderSend(check As String)
SendCurrentItems check
DailyCheckIndex_CalculateTotals check
End Sub



Public Sub OrderQueuedItems(coll As Collection)
Dim q As Integer
q = 1
GetQty
Dim i As Integer


'Dim tempcoll As Collection
'Set tempcoll = FormatItemCollection(coll)
'FormatChildSpacing tempcoll
'FormatSides tempcoll

Dim DailyCheckDetail As New zclsDailyCheckDetail
Do While q <= Qty
    DailyCheckDetail.AddCurrentItems coll
    q = q + 1
Loop
Set collCheckData = RecallCheckLines(currentcheck)
WriteCheckLines Sheet1.range("CheckRange"), collCheckData
PopGuiWithCheckAttributes GetCheckAttributes(currentcheck), Sheet1
SetLastEntry coll
ResetOrderState

Set DailyCheckDetail = Nothing


End Sub
Public Sub ResetOrderState()
GetItemCell
SetEndOfCheck
CloseFrames
Qty = 1
End Sub

Public Sub GetQty()
If IsNull(Qty) Or Qty = 0 Then Qty = 1
End Sub

Public Sub SetLastEntry(coll As Collection)
If coll.Count = 0 Then Exit Sub
Set LastEntry = DuplicateCollection(coll)
ClearCollection CItem
End Sub

Public Sub RepeatLastEntry()
'MsgBox "Unavailable"
'Exit Sub
OrderQueuedItems LastEntry
End Sub

Public Sub CloseCheck()
RestoreState
OrderSend currentcheck
DailyCheckIndex_CloseOrder currentcheck
ThisOrder.OrderType.CloseCheck currentcheck
MsgBox "Check " & currentcheck & " closed."
ActivateHomeScreen
End Sub

Public Sub QuickMod()
Dim bname As String
bname = Application.caller
Sheet1.SpecialInstructionText.value = Sheet1.Shapes(bname).TextFrame.Characters.text
Dim ItemID As Integer
ItemID = GetNextItemID("SpclInstruction")
ItemDirector ItemID
Sheet1.SpecialInstructionText.value = ""
End Sub

Public Function MissingComponents() As Boolean
Dim i As Integer
For i = CItem.Count To 1 Step -1
    Dim Reqs As Variant
    Reqs = AllRequirementsFullfilled(CItem(i))
    If Not Reqs = True Then
        SetCurrentParent CItem(i).CollID
        MissingComponents = True
        Exit Function
    End If
Next i
MissingComponents = False
End Function

Public Sub ScrollParents()

Dim CurrentParentCollIndex As Integer
Dim i As Integer
For i = 1 To CItem.Count
    If CItem(i).CollID = CurrentParent Then CurrentParentCollIndex = i
Next i

If Not CurrentParentCollIndex = CItem.Count Then
    SetCurrentParent CItem(CurrentParentCollIndex + 1).CollID
    Exit Sub
End If

SetCurrentParent CItem(1).CollID

    

End Sub



