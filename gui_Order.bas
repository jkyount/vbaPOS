Attribute VB_Name = "gui_Order"
Option Explicit

Public time1 As Double
Public time2 As Double
Dim CurrentFamily As String
Dim CurrentPage As Integer
Dim CurrentBtnGrp As String

Public Sub SetCurrentFamily(Family As String)
CurrentFamily = Family
End Sub
Public Sub SetCurrentPage(Page As Integer)
CurrentPage = Page
End Sub
Public Sub SetCurrentBtnGrp(BtnGrp As String)
CurrentBtnGrp = BtnGrp
End Sub

Public Function GetCurrentFamily() As String
GetCurrentFamily = CurrentFamily
End Function

Public Function GetCurrentPage() As Integer
GetCurrentPage = CurrentPage
End Function
Public Sub PopGuiWithCheckAttributes(dict As Dictionary, sheet As Worksheet)
AssignRangeAttributeValues dict, sheet
AssignShapeAttributeValues dict, sheet
End Sub
Public Sub AssignShapeAttributeValues(dict As Dictionary, sheet As Worksheet)
sheet.Shapes("ServerName").TextFrame.Characters.text = dict("ServerName")
sheet.Shapes("OrderName").TextFrame.Characters.text = dict("OrderName")
End Sub

Public Sub AssignRangeAttributeValues(dict As Dictionary, sheet As Worksheet)

Dim val As Variant
On Error Resume Next
For Each val In dict
    sheet.range(val).value = dict(val)
Next val
On Error GoTo 0

End Sub

Public Sub ShowItemPage(iFamily As zclsFamily, Page As Integer)
Dim coll As New Collection
Set coll = iFamily.GetMembersColl
If Page > coll.Count Then
    Page = Page - 1
    Exit Sub
End If


Dim dict As New Dictionary
Set dict = coll(Page)
Dim BtnGroup As String
BtnGroup = "grpgui" & iFamily.MenuStyle
ShowShape BtnGroup
Dim Group As Shape
Set Group = ActiveSheet.Shapes(BtnGroup)
Dim Btn As Shape
Dim key As Variant
Dim i As Integer
i = 1
For Each key In dict.Keys
    If Not key = "705" Then
    Set Btn = Group.GroupItems.item(i)
    Btn.name = key
    Btn.TextFrame.Characters.text = dict(key)
    i = i + 1
    End If
Next key

For i = i To Group.GroupItems.Count
    Group.GroupItems.item(i).Visible = msoFalse
Next i
Set coll = Nothing
Set dict = Nothing
Set Group = Nothing
Set Btn = Nothing
Set key = Nothing
End Sub

Public Sub ShowCategoryItems(iFamily As zclsFamily, Page As Integer)
ShowItemPage iFamily, Page
CurrentPage = Page
ShowShape "grpScrollCategoryItems"
SetShapeText "ItemScrollFrame", ValueMatch(iFamily.Wrap(GetNewMatchObj("Family", iFamily.Family)), "DisplayName")
End Sub

Public Sub ScrollCategoryItemsForward()
If CurrentPage = 0 Then Exit Sub

Dim iFamily As New zclsFamily
iFamily.Family = CurrentFamily
ShowCategoryItems iFamily, CurrentPage + 1
End Sub

Public Sub ScrollCategoryItemsBackward()
If CurrentPage = 0 Then Exit Sub
If CurrentPage = 1 Then Exit Sub


Dim iFamily As New zclsFamily
iFamily.Family = CurrentFamily
ShowCategoryItems iFamily, CurrentPage - 1
End Sub

Public Sub CategorySelect(Category As String)

If CurrentFamily = Category Then
    HideAllItemButtons
    Exit Sub
End If

CurrentFamily = Category
Dim i As Integer
Dim iFamily As New zclsFamily
Dim iMenu As New zclsMenu
Dim CategoryButtonName As String
'Application.ScreenUpdating = False
CategoryButtonName = Category
iFamily.Family = Category
Dim ButtonGroup As String
ButtonGroup = "grpgui" & iFamily.MenuStyle

If ValueMatch(iFamily.Wrap(GetNewMatchObj("Family", Category)), "MultiMenu") = True Then
    CategoryButtonName = "MultiMenu"
    HideAllItemButtons
    ShowShape ("grpgui" & CategoryButtonName)
    DisplayMultiMenu Category
    HighlightSelectedMenuCategory Category
    CurrentFamily = Category
    'Application.ScreenUpdating = True
    Set iFamily = Nothing
    Set iMenu = Nothing
    Exit Sub
End If

If CountMatch(iMenu.Wrap(GetNewMatchObj("Family", Category, "NOT ItemName", ""))) = 0 Then
    'Application.ScreenUpdating = True
    Exit Sub
End If

If iFamily.MenuStyle = "Condensed" Then
    Sheet1.Shapes(ButtonGroup).Left = Sheet1.Shapes(Category).Left
End If

HideAllItemButtons
HighlightSelectedMenuCategory Category
CurrentPage = 1
CurrentBtnGrp = ButtonGroup
CurrentFamily = Category
ShowCategoryItems iFamily, CurrentPage
'Application.ScreenUpdating = True
Set iFamily = Nothing
Set iMenu = Nothing
End Sub

Public Sub DisplayMultiMenu(Category As String)
Dim time As Double
Dim time2 As Double
time = Timer


Dim iMultiMenu As New aclsMultiMenu
Dim coll As New Collection
Set coll = iMultiMenu.Members(Category)
Dim shp As Shape
Set shp = Sheet1.Shapes("grpguiMultiMenu")
Dim i As Integer
For i = 1 To coll.Count
    shp.GroupItems(i).name = coll(i)
    shp.GroupItems(i).TextFrame.Characters.text = coll(i)
Next i

For i = coll.Count + 1 To shp.GroupItems.Count
    shp.GroupItems(i).Visible = msoFalse
Next i
shp.Left = Sheet1.Shapes(Category).Left
Set iMultiMenu = Nothing
Set coll = Nothing
Set shp = Nothing

time2 = Timer
Debug.Print "DisplayMultiMenu  " & time2 - time
End Sub
Public Sub HighlightSelectedMenuCategory(Category As String)
Dim time As Double
Dim time2 As Double
time = Timer


Dim shp As Shape
For Each shp In Sheet1.Shapes("MenuCategory").GroupItems
    shp.line.ForeColor.RGB = rgbBlack
Next shp
Sheet1.Shapes(Category).line.ForeColor.RGB = rgbWhite

time2 = Timer
Debug.Print "HighlightSelectedMenuCategory  " & time2 - time
End Sub

Public Sub HideAllItemButtons()
Dim shp As Shape
For Each shp In Sheet1.Shapes
    If shp.name Like "grpgui*" Then
        HideShape shp.name
    End If
Next shp
HideShape "grpScrollCategoryItems"
CurrentFamily = ""
CurrentBtnGrp = ""
CurrentPage = 0
End Sub
Public Sub ShowComponentFrame(RequiredComponent As String)
'5/27
     HideAllItemButtons 'NEW
'OLD    HideShape ("frmSalad")
'OLD    HideShape ("frmDrsng")
'OLD    HideShape ("frmPasta")
'OLD    HideShape ("frmSce")

Dim iFamily As New zclsFamily
iFamily.Family = RequiredComponent
Dim Family As String
Family = iFamily.Family
Dim ButtonGroup As String
ButtonGroup = "grpgui" & iFamily.MenuStyle
    ShowCategoryItems iFamily, 1
    ShowShape ("Component")
    'ShowShape RequiredComponent
    Set iFamily = Nothing
    
CurrentPage = 1
CurrentBtnGrp = ButtonGroup
CurrentFamily = Family
End Sub


Public Sub ShowParentFrame(ParentName As String)
SetShapeText "ParentFrame", "Selecting options for " & ParentName & "."
ShowShape "ParentFrame"
ShowShape "ScrollParents"
End Sub
Public Sub DisplayQuickMods(RequiredComponent As String)
Select Case RequiredComponent
    Case "Sce"
        SetShapeText "QuickMod1", "Sce on Side"
        SetShapeText "QuickMod2", "Xtra Sce"
        SetShapeText "QuickMod3", "Xtra Sce Side"
        SetShapeText "QuickMod4", "Light Sce"
        ShowShape "grpQuickMod"
    Case "Drsng"
        SetShapeText "QuickMod1", "Drsng on Side"
        SetShapeText "QuickMod2", "Xtra Drsng Side"
        'SetShapeText "QuickMod3", ""
        'SetShapeText "QuickMod1", ""
        ShowShape "grpQuickMod"
        HideShape "QuickMod3"
        HideShape "QuickMod4"
    Case "PzaTop"
        SetShapeText "QuickMod1", "Half [A]"
        SetShapeText "QuickMod2", "Half [B]"
        SetShapeText "QuickMod3", "Well Done Pza"
        SetShapeText "QuickMod4", "Pza To-GO"
        ShowShape "grpQuickMod"
End Select
End Sub
Public Sub CloseFrames()
'5/27
    HideAllItemButtons 'NEW
    HideShape "frmSaladFrame"
    HideShape "frmPza"
    HideShape "frmModConnector" 'NEW
    HideShape "shpBLOCKBOX"
    HideShape "grpCustomItem"
    HideShape "grpSpecialInstruction"
    HideShape "btnDone"
    HideShape "grpQuickMod"
    HideShape "Component"
    HideShape "ParentFrame"
    HideShape "ScrollParents"
    HideShape "ComponentBLOCK"
    HideShape "grpPreviewWindow"
'    HideShape "grpguiBeerWine"
End Sub
Public Sub SetOrderTypeIndicator(OrderType As String)
SetShapeTrans "btnCarryout", 0.75, 0.75
SetShapeTrans "btnDineIn", 0.75, 0.75
SetShapeTrans "btn" & OrderType, 0, 0
End Sub

Public Sub GoToComponent()
Dim bname As String
bname = Application.caller
ShowComponentFrame bname
DisplayQuickMods bname
End Sub

Public Function GetFirstButtonIndex(FamilyGroup As String) As Integer
Dim iFamily As New zclsFamily
Dim iDataObj As aclsDataObject
Dim qry As String
qry = "SELECT Family FROM Family WHERE FamilyGroup = """ & FamilyGroup & """ ORDER BY ID ASC"
Set iDataObj = iFamily.Wrap(iDataObj)
Dim rs As ADODB.RecordSet
Set rs = GetRecordsetMatch(iDataObj, qry)
Dim i As Integer
Dim GroupCount As Integer
GroupCount = Sheet1.Shapes(FamilyGroup).GroupItems.Count
i = 0
Do
    If rs.Fields("Family") = Sheet1.Shapes(FamilyGroup).GroupItems(1).name Then
        Exit Do
    End If
    i = i + 1
    rs.MoveNext
Loop Until rs.EOF
GetFirstButtonIndex = i
iFamily.CloseDbs
End Function

Public Function GetMenuMembers(FamilyGroup As String) As Collection

Dim iFamily As New zclsFamily
Dim qry As String
qry = "SELECT * FROM Family WHERE FamilyGroup = """ & FamilyGroup & """ AND Active = True AND MultiMenuMember = False ORDER BY ID ASC"
Set GetMenuMembers = CDictCollection(GetRecordsetMatch(iFamily.Wrap(GetNewMatchObj), qry))
iFamily.CloseDbs
Set iFamily = Nothing
End Function



Public Sub ScrollForward(FamilyGroup As String)

Dim FirstButtonIndex As Integer
FirstButtonIndex = GetFirstButtonIndex(FamilyGroup)
Dim i As Integer
Dim GroupCount As Integer
GroupCount = Sheet1.Shapes(FamilyGroup).GroupItems.Count
Dim coll As Collection
Set coll = GetMenuMembers(FamilyGroup)
Dim MenuLength As Integer
Dim iFamGroup As New zclsFamilyGroup
MenuLength = ValueMatch(iFamGroup.Wrap(GetNewMatchObj("FamilyGroup", FamilyGroup)), "MenuLength")
If Not FirstButtonIndex >= coll.Count - MenuLength Then
    For i = 1 To GroupCount
        Sheet1.Shapes(FamilyGroup).GroupItems(i).name = coll(i + FirstButtonIndex + 1)("Family")
        Sheet1.Shapes(FamilyGroup).GroupItems(i).TextFrame.Characters.text = coll(i + FirstButtonIndex + 1)("DisplayName")
    Next i
End If
Set iFamGroup = Nothing
Set coll = Nothing

End Sub


Public Sub ScrollBackward(FamilyGroup As String)
Dim FirstButtonIndex As Integer
FirstButtonIndex = GetFirstButtonIndex(FamilyGroup)

If FirstButtonIndex > 0 Then
    Dim coll As Collection
    Set coll = GetMenuMembers(FamilyGroup)
    
    Dim i As Integer
    Dim GroupCount As Integer
    GroupCount = Sheet1.Shapes(FamilyGroup).GroupItems.Count
    If coll.Count > GroupCount Then
        For i = 1 To GroupCount
        
            Sheet1.Shapes(FamilyGroup).GroupItems(i).name = coll(i - 1 + FirstButtonIndex)("Family")
            Sheet1.Shapes(FamilyGroup).GroupItems(i).TextFrame.Characters.text = coll(i - 1 + FirstButtonIndex)("DisplayName")
        Next i
    End If
End If
End Sub

Public Sub UpdateMenuButtons(FamilyGroup As String)
'5/18/22 moved from fMultiMenu and fFamily
Dim coll As Collection
Set coll = GetMenuMembers(FamilyGroup)
Dim i As Integer
Dim GroupCount As Integer
GroupCount = Sheet1.Shapes(FamilyGroup).GroupItems.Count
For i = 1 To coll.Count
    If i > GroupCount Then Exit For
    Sheet1.Shapes(FamilyGroup).GroupItems(i).name = coll(i)("Family")
    Sheet1.Shapes(FamilyGroup).GroupItems(i).TextFrame.Characters.text = coll(i)("DisplayName")
Next i
If i <= GroupCount Then
    Dim k As Integer
    For k = i To GroupCount
        Sheet1.Shapes(FamilyGroup).GroupItems(k).name = "EmptyBtn"
        Sheet1.Shapes(FamilyGroup).GroupItems(k).TextFrame.Characters.text = ""
    Next k
End If
Set coll = Nothing
End Sub


Public Sub RefreshPreviewWindow()
Dim shp As Shape
Dim ShpText As String
Dim coll As Collection
Dim Spacer As String
Spacer = ""
Dim SpacerCount As Integer
SpacerCount = 1
Set coll = OrderByParent(DuplicateCItem(CItem))

Dim i As Integer, k As Integer
Sheet1.Shapes("grpPreviewWindow").Visible = msoTrue
For i = 1 To coll.Count
    Set shp = Sheet1.Shapes("grpPreviewWindow").GroupItems(i)
    
    shp.name = CStr(coll(i).CollID)
    ShpText = coll(i).ItemName
    
    If Not i = 1 Then
        SpacerCount = 1
        Spacer = ""
        Dim item As aclsItem
        Set item = coll(i)
        Dim ParentID As Integer
        ParentID = item.Parent.ID
        Do Until ParentID = 1
            SpacerCount = SpacerCount + 1
            Set item = GetItemByID(item.Parent.ID)
            ParentID = item.Parent.ID
        Loop
        For k = 1 To SpacerCount
            Spacer = Spacer & "     "
        Next k
        Spacer = Spacer & "-"
        ShpText = Spacer & coll(i).ItemName
    End If
    shp.TextFrame.Characters.text = ShpText
    shp.Visible = msoTrue
Next i

For i = coll.Count + 1 To Sheet1.Shapes("grpPreviewWindow").GroupItems.Count
    Set shp = Sheet1.Shapes("grpPreviewWindow").GroupItems(i)
    shp.Visible = msoFalse
Next i
End Sub

Public Sub ReviewSelections()
Dim bname As String
bname = Application.caller
Dim SelectedItem As aclsItem
Set SelectedItem = GetItemByID(CInt(bname))
If SelectedItem.Parent.ID = -1 Then Exit Sub

SelectedItem.Parent.item.UnassignChild CInt(bname)
RemoveItemFromQueue SelectedItem
RefreshPreviewWindow
SetCurrentParent SelectedItem.Parent.ID
ShowComponentFrame SelectedItem.Family
End Sub

Public Sub ValidateCustomItemRequest()
If Sheet1.OLEObjects("CustomItemName").Object.value = "" Then
    MsgBox "Please provide a name for the item."
    Exit Sub
End If
If Not IsNumeric(Sheet1.OLEObjects("CustomItemPrice").Object.value) Then
    MsgBox "Please provide a price for the item."
    Exit Sub
End If

Ssfub
End Sub

Public Function FormatChildSpacing(coll As Collection) As Collection
Dim i As Integer, k As Integer
Dim SpacerCount As Integer
Dim Spacer As String
For i = 2 To coll.Count
    SpacerCount = 1
    Spacer = ""
    Dim item As aclsItem
    Set item = coll(i)
    Dim ParentID As Integer
    ParentID = item.Parent.ID
    Do Until ParentID = 1
        SpacerCount = SpacerCount + 1
        Set item = GetItemByID(item.Parent.ID)
        ParentID = item.Parent.ID
    Loop
    For k = 1 To SpacerCount
        Spacer = Spacer & "    "
    Next k
    coll(i).ItemName = Spacer & coll(i).ItemName
Next i
Set FormatChildSpacing = coll
End Function



