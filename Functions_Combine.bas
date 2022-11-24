Attribute VB_Name = "Functions_Combine"
Option Explicit

Public Sub ActivateCombineCheckScreen()
Sheet12.ScrollArea = "A1:U36"
Sheet12.Activate
InitializeRanges
DisplayChecks GetChecks("ServerNum", ThisEmployee.ServerNum, "Closed = False AND NOT CheckNumber", currentcheck)
End Sub


Private Sub DisplayChecks(arrChecks As Variant)
Application.ScreenUpdating = False
Dim i As Integer
Dim k As Integer
Dim rg As range
i = 0
Do Until i > UBound(arrChecks, 1) Or i + 1 > 12
    Set rg = Sheet12.range("CheckSlot" & i + 1)
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

    Dim CheckSlotRange As range
    Do Until i + 1 > 12
        Set rg = Sheet12.range("CheckSlot" & i + 1)
        Set CheckSlotRange = Sheet12.range(rg, rg.Offset(0, 8))
        CheckSlotRange.value = ""
        CheckSlotRange.Borders.LineStyle = xlNone
        CheckSlotRange.Interior.color = 15921906
        HideShape "" & i + 1
        i = i + 1
    Loop

Set rg = Nothing
Set CheckSlotRange = Nothing
Application.ScreenUpdating = True


End Sub

Private Sub AdvRange(rg As range)
Set rg = rg.Offset(1, 0)
End Sub

Public Sub ClickCheckToCombine()

Dim bname As String, CheckToCombine As String
bname = Application.caller
Dim OrderToCombine As New aclsOrder

CheckToCombine = Sheet12.range("CheckSlot" & bname).value
OrderToCombine.ImportCheckDetails OrderToCombine, CheckToCombine

If CheckToCombine = currentcheck Then
    MsgBox "Cannot combine check to itself.  Please select a different check."
    Exit Sub
End If
If SameOrderType(OrderToCombine) = False Then
    MsgBox "Cannot combine orders of different order types (carryout/dine-in)."
    Exit Sub
End If
'Dim dict As New Dictionary
'Dim iIndex As New zclsDailyCheckIndex
'Set dict = GetValueDict(iIndex.Wrap(GetNewMatchObj(, CheckToCombine)))(1)
Sheet12.range("TargetCheckNumber").value = OrderToCombine.ValueDict("CheckNumber")
Sheet12.range("TargetServerName").value = OrderToCombine.ValueDict("ServerName")
Sheet12.range("TargetOrderName").value = OrderToCombine.ValueDict("OrderName")
Sheet12.range("TargetPhone").value = OrderToCombine.ValueDict("Phone")
WriteCheckLines Sheet12.range("CombineCheckRange"), RecallCheckLines(CheckToCombine)
If MsgBox("Combine these checks?", vbYesNo) = vbYes Then
    CombineChecks currentcheck, CheckToCombine
    MsgBox ("Checks combined")
    RecallCheck ThisOrder, Sheet1
    ActivateOrderScreen
    Exit Sub
End If
Set OrderToCombine = Nothing
End Sub

Private Function SameCheck(CheckToCombine As String) As Boolean
SameCheck = False
If CheckToCombine = currentcheck Then SameCheck = True
End Function

Private Function SameOrderType(OrderToCombine As aclsOrder) As Boolean
SameOrderType = False
If OrderToCombine.ValueDict("DineIn") = ThisOrder.ValueDict("DineIn") Then SameOrderType = True
End Function

Public Sub CombineChecks(OriginalCheck As String, CheckToCombine As String)
CopyToTemp OriginalCheck
AppendToTemp CheckToCombine
ReconcileLocalGroup OriginalCheck, CheckToCombine
ReconcileEntityGroup OriginalCheck, CheckToCombine
ReconcileCheckNumber OriginalCheck, CheckToCombine
SubmitChanges OriginalCheck
DailyCheckIndex_CalculateTotals OriginalCheck
Cleanup CheckToCombine
End Sub

Private Sub ReconcileLocalGroup(OriginalCheck As String, CheckToCombine As String)
Dim GreatestLocalGroup As Integer
GreatestLocalGroup = GetNextLocalGroup(OriginalCheck)
Dim iDataObject As New aclsDataObject
Dim iTemp As New zclsTemp
Set iDataObject = iTemp.Wrap(GetNewMatchObj(, CheckToCombine))
Dim rs As ADODB.RecordSet
Set rs = GetRecordsetMatch(iDataObject, ConstructMatchQuery(iDataObject))
Dim i As Integer
For i = 1 To rs.RecordCount
    rs.Fields("LocalGroup").value = GreatestLocalGroup + (i - 1)
    rs.Update
    rs.MoveNext
Next i
Set rs = Nothing
Set iTemp = Nothing
iDataObject.CloseDbs iDataObject
Set iDataObject = Nothing
End Sub

Private Sub ReconcileEntityGroup(OriginalCheck As String, CheckToCombine As String)
Dim GreatestEntityGroup As Integer
GreatestEntityGroup = GetNextEntityGroup(OriginalCheck) - 1
Dim iDataObject As New aclsDataObject
Dim iTemp As New zclsTemp
Set iDataObject = iTemp.Wrap(GetNewMatchObj(, CheckToCombine))
Dim rs As ADODB.RecordSet
Set rs = GetRecordsetMatch(iDataObject, ConstructMatchQuery(iDataObject))
Do Until rs.EOF
    rs.Fields("EntityGroup").value = rs.Fields("EntityGroup").value + GreatestEntityGroup
    rs.Update
    rs.MoveNext
Loop
Set rs = Nothing
iDataObject.CloseDbs iDataObject
Set iDataObject = Nothing
End Sub

Private Sub ReconcileCheckNumber(OriginalCheck As String, CheckToCombine As String)
Dim iTemp As New zclsTemp
Update iTemp.Wrap(GetNewUpdateObj("CheckNumber", CheckToCombine, "CheckNumber", OriginalCheck))
Set iTemp = Nothing
End Sub

Private Sub Cleanup(CheckToCombine)
'5/20
Dim CheckIndex As New zclsDailyCheckIndex
DeleteMatch CheckIndex.Wrap(GetNewMatchObj(, CheckToCombine))
Dim CheckDetail As New zclsDailyCheckDetail
DeleteMatch CheckDetail.Wrap(GetNewMatchObj(, CheckToCombine))
ThisOrder.OrderType.CloseCheck CheckToCombine 'new
'old  -------------ThisOrder.Cleanup---------
'Dim iTable As New zclsTable
'iTable.Table = ValueMatch(iTable.Wrap(GetNewMatchObj(, CheckToCombine)), "Table")
'iTable.UnassignCheck (CheckToCombine)
'-------------------------------------------

End Sub

