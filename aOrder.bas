Attribute VB_Name = "aOrder"

Option Explicit





Public Sub test()

'==========================================================================
Dim time As Double, time2 As Double
time = Timer
'==========================================================================






'==========================================================================
time2 = Timer
Debug.Print time2 - time
'==========================================================================

End Sub








Public Sub Test2()
'==========================================================================
Dim time As Double, time2 As Double
time = Timer
'==========================================================================
Application.ScreenUpdating = False
Dim AccApp As Object
Dim DBPath As String
DBPath = "C:\Jared\POS\Access_Int\Menu.accdb"
Set AccApp = CreateObject("Access.Application")
AccApp.Visible = True
AccApp.OpenCurrentDatabase DBPath

AccApp.DoCmd.OpenForm "MenuItem"
Set AccApp = Nothing
Application.ScreenUpdating = True
'==========================================================
time2 = Timer
Debug.Print time2 - time
'==========================================================================
End Sub







Public Sub Test3()
'==========================================================================
Dim time As Double, time2 As Double
time = Timer
'==========================================================================

Dim iMenu As New zclsMenu
iMenu.OpenDbs
Dim rs As New ADODB.RecordSet
Set rs = iMenu.GetRs

rs.Source = "AllItems"
rs.Open
Do Until rs.EOF
    If rs.Fields("Req2").value = Null Or EmptyCheck(rs.Fields("Req1").value) = True Then rs.Fields("Req1").value = ""
    If rs.Fields("Req2").value = Null Or EmptyCheck(rs.Fields("Req2").value) = True Then rs.Fields("Req2").value = ""
    rs.MoveNext
Loop
iMenu.CloseDbs

    



'==========================================================================
time2 = Timer
Debug.Print time2 - time
'==========================================================================
End Sub

Public Function EmptyCheck(val As String) As Boolean
    Select Case val
        Case "Salad"
            EmptyCheck = False
            Exit Function
        Case "Drsng"
            EmptyCheck = False
            Exit Function
        Case "Pasta"
            EmptyCheck = False
            Exit Function
        Case "Sce"
            EmptyCheck = False
            Exit Function
        Case "PzaTop"
            EmptyCheck = False
            Exit Function
    End Select
    EmptyCheck = True
        
        
End Function
'
'Public Sub TransferToNewFamily(ItemID As Integer, NewFamily As String)
'Dim iBtn As New aclsItemButton
'Dim iDataObj As aclsDataObject
'Set iDataObj = New aclsDataObject
'Dim iMenu As zclsMenu
'Set iMenu = New zclsMenu
'Dim OriginalFamilyName As String
'OriginalFamilyName = ValueMatch(iMenu.Wrap(GetNewMatchObj("ID", ItemID)), "Family")
'Dim iFamily As zclsFamily
'Set iFamily = New zclsFamily
'iFamily.Family = NewFamily
'Dim OriginalFamily As New zclsFamily
'Set OriginalFamily = GetFamilyButtonObj(OriginalFamilyName)
'OriginalFamily.BtnLocation.Shapes(CStr(ItemID)).Delete
'If IsEmpty(iFamily.Members) Then
'    'fMenuItems.InitializeNewFamily iFamily, ItemID, ValueMatch(iMenu.Wrap(GetNewMatchObj("ID", ItemID)), "ItemName")
'    iFamily.Activate
'End If
'Set iDataObj = iMenu.Wrap(GetNewUpdateObj("ID", ItemID, "Family", NewFamily))
'Update iDataObj
''iBtn.PositionAll GetFamilyButtonObj(NewFamily)
'If IsEmpty(OriginalFamily.Members) Then
'    OriginalFamily.BtnLocation.Shapes(OriginalFamily.Family & "Blank").Delete
'    Set iDataObj = Nothing
'    Set iFamily = Nothing
'    Set iMenu = Nothing
'    Set iBtn = Nothing
'    Exit Sub
'End If
''iBtn.PositionAll GetFamilyButtonObj(OriginalFamilyName)
'Set iDataObj = Nothing
'Set iFamily = Nothing
'Set iMenu = Nothing
'Set iBtn = Nothing
'End Sub




Public Sub SCharge()
Dim check As String
check = currentcheck
Dim DetailMatch As New zclsDailyCheckDetail
Dim Ord As New aclsOrder
Ord.ImportCheckDetails Ord, check
Ord.OrderType = Ord.SameOrderType
Dim subtotal As Currency
Dim Tax As Currency
Dim ServiceCharge As Currency
subtotal = SumMatch(DetailMatch.Wrap(GetNewMatchObj(, check)), "Price")
ServiceCharge = Ord.ValueDict("ServiceCharge")
Tax = Ord.OrderType.GetTax(check)

Sheet1.Shapes("Scharge").TextFrame.Characters.text = (Tax + subtotal) + ((Tax + subtotal) * 0.035)
Set DetailMatch = Nothing
Set Ord = Nothing

End Sub



