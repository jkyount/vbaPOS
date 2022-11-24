VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsItemButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'
'Private pShape As Shape
'Private pName As String
'Private pCaption As String
'Private pLocalGroup As Integer
'
'Private pLocation As Worksheet
'Private pStyleDict As Dictionary
'Private pFamily As String
''Private pBtnType As Variant
'
'
'
'
'
'Public Property Let Shape(value As Shape)
'Set pShape = value
'End Property
'Public Property Get Shape() As Shape
'Set Shape = pShape
'End Property
'
'Public Property Let LocalGroup(value As Integer)
'pLocalGroup = value
'End Property
'Public Property Get LocalGroup() As Integer
'LocalGroup = pLocalGroup
'End Property
'
'Public Property Let name(value As String)
'pName = value
'End Property
'Public Property Get name() As String
'name = pName
'End Property
'
'Public Property Let Caption(value As String)
'pCaption = value
'End Property
'Public Property Get Caption() As String
'Caption = pCaption
'End Property
'
'Public Property Let Location(value As Worksheet)
'Set pLocation = value
'End Property
'Public Property Get Location() As Worksheet
'Set Location = pLocation
'End Property
'
'Public Property Let StyleDict(value As Dictionary)
'Set pStyleDict = value
'End Property
'
'Public Property Get StyleDict() As Dictionary
'Set StyleDict = pStyleDict
'End Property
'
'Public Property Let Family(value As String)
'pFamily = value
'End Property
'Public Property Get Family() As String
'Family = pFamily
'End Property
'
''Public Property Let BtnType(value As Variant)
''Set pBtnType = value
''End Property
''
''Public Property Get BtnType() As Variant
''Set BtnType = pBtnType
''End Property
'
'
''Public Function GetBtnDict(iItem As aclsItem, iFamily As zclsFamily) As Dictionary
''
''Dim dict As New Dictionary
''dict.Add "Name", iItem.ItemID
''dict.Add "Caption", iItem.ItemName
''dict.Add "Shape", GenerateNewShape(iFamily)
''dict.Add "Location", iFamily.BtnLocation
''dict.Add "StyleDict", iFamily.StyleDict
''dict.Add "LocalGroup", iFamily.Count
''dict.Add "Family", iFamily.Family
''Set GetBtnDict = dict
''
''End Function
'Private Function GetNewBtnObj(NewItem As Variant) As aclsItemButton
'Dim x As New aclsItemButton
'x.name = CStr(NewItem.ItemID)
'x.Caption = NewItem.ItemName
'Set GetNewBtnObj = x
'End Function
'
''Private Function bGetNewBtnObj(dict As Dictionary) As aclsItemButton
''Dim x As New aclsItemButton
''x.name = dict("Name")
''x.Caption = dict("Caption")
''x.Location = dict("Location")
''x.StyleDict = dict("StyleDict")
''x.LocalGroup = dict("LocalGroup")
''Set bGetNewBtnObj = x
''End Function
'
''Public Sub bAddNew(iItemBtn As aclsItemButton)
''bGetNewShape iItemBtn
''
'''PositionAll iFamily
''End Sub
'
'Public Sub AddNew(NewItem As Variant, Family As String)
'Dim iItemButton As aclsItemButton
'Set iItemButton = GetNewBtnObj(NewItem)
'Dim iFamily As New zclsFamily
'Set iFamily = iFamily.GetFamilyButtonObj(Family)
'iItemButton.LocalGroup = iFamily.Count
'iItemButton.Shape = GetNewShape(iFamily, iItemButton)
''PositionAll iFamily
'Set iItemButton = Nothing
'Set iFamily = Nothing
'End Sub
'
''Public Sub bGetNewShape(iItemBtn As aclsItemButton)
''Dim NewShape As Shape
''Set NewShape = bGenerateNewShape(iItemBtn)
''NameShape NewShape, iItemBtn.name
''SetCaption NewShape, iItemBtn.Caption
''iItemBtn.Shape = NewShape
''Set NewShape = Nothing
''End Sub
'
'Public Function GetNewShape(iFamily As zclsFamily, iItemButton As aclsItemButton) As Shape
'Dim NewShape As Shape
'Set NewShape = GenerateNewShape(iFamily)
'NameShape NewShape, iItemButton.name
'SetCaption NewShape, iItemButton.Caption
'Set GetNewShape = NewShape
'Set NewShape = Nothing
'End Function
'
''Public Function bGenerateNewShape(iItemBtn As aclsItemButton) As Shape
''Dim shp As Shape
''Dim ExistingButtonName As String
''For Each shp In iItemBtn.Location.Shapes("grpgui" & iItemBtn.Family).GroupItems
''    If shp.name Like "Item*" Then
''        ExistingButtonName = shp.name
''        Exit For
''    End If
''Next shp
''Set shp = iItemBtn.Location.Shapes(ExistingButtonName)
''Set bGenerateNewShape = shp.Duplicate
''Set shp = Nothing
''End Function
'
'Public Function GenerateNewShape(iFamily As zclsFamily) As Shape
'Dim shp As Shape
'Dim ExistingButtonName As String
'For Each shp In iFamily.BtnLocation.Shapes("grpgui" & iFamily.Family).GroupItems
'    If Not shp.name Like "*Blank" Then
'        ExistingButtonName = shp.name
'        Exit For
'    End If
'Next shp
'Set shp = iFamily.BtnLocation.Shapes(ExistingButtonName)
'Set GenerateNewShape = shp.Duplicate
'
''ReplaceDictValue iFamily.StyleDict, "BtnTopRef", shp.Top 'new
''ReplaceDictValue iFamily.StyleDict, "BtnLeftRef", shp.Left 'new
'Set shp = Nothing
'
'End Function
''
''Public Sub bPosition(iItemBtn As aclsItemButton)
''iItemBtn.Shape.Top = iItemBtn.StyleDict("BtnTopRef") + (bGetButtonRow(iItemBtn) - 1) * iItemBtn.StyleDict("BtnRowOffset")
''iItemBtn.Shape.Left = iItemBtn.StyleDict("BtnLeftRef") + (bGetButtonColumn(iItemBtn) - 1) * iItemBtn.StyleDict("BtnColumnOffset")
''End Sub
'
'Public Sub Position(iFamily As zclsFamily, iItemButton As aclsItemButton)
'iItemButton.Shape.Top = iFamily.StyleDict("BtnTopRef") + (GetButtonRow(iFamily, iItemButton.LocalGroup) - 1) * iFamily.StyleDict("BtnRowOffset")
'iItemButton.Shape.Left = iFamily.StyleDict("BtnLeftRef") + (GetButtonColumn(iFamily, iItemButton.LocalGroup) - 1) * iFamily.StyleDict("BtnColumnOffset")
'End Sub
'
'Public Sub PositionAll(iFamily As zclsFamily)
'
'Dim iItemButton As New aclsItemButton
'Dim arr As Variant
'arr = iFamily.Members
'Dim FamilyCount As Integer, ShapeCount As Integer
'FamilyCount = iFamily.Count
'
'ShapeCount = iFamily.BtnLocation.Shapes("grpgui" & iFamily.Family).GroupItems.Count
'
''5/23
'If FamilyCount > ShapeCount Then AddMissingButtons iFamily
'If FamilyCount < ShapeCount Then RemoveExtraButtons iFamily
''/5/23
'
'
'
'Dim i As Integer
'For i = 1 To FamilyCount
'    'On Error GoTo NextI
'    iItemButton.Shape = iFamily.BtnLocation.Shapes(CStr(arr(0, i - 1)))
'    iItemButton.LocalGroup = i
'    iItemButton.Position iFamily, iItemButton
'    iItemButton.Format iFamily.StyleDict, iItemButton 'new
''NextI:
'Next i
'
'Set iItemButton = Nothing
'End Sub
'
'Private Sub RemoveExtraButtons(iFamily As zclsFamily)
''5/23
'Dim MembersDict As Dictionary
'Set MembersDict = MembersToDict(iFamily.Members)
'Dim shp As Shape
'For Each shp In iFamily.BtnLocation.Shapes("grpgui" & iFamily.Family).GroupItems
'    If Not MembersDict.Exists(shp.name) Then
'        If shp.name Like "Item*" Then
'            shp.Delete
'        End If
'    End If
'Next shp
'Set MembersDict = Nothing
'Set shp = Nothing
'End Sub
'
'Private Sub AddMissingButtons(iFamily As zclsFamily)
''5/23
'Dim MembersDict As Dictionary
'Set MembersDict = MembersToDict(iFamily.Members)
'Dim ShapeDict As Dictionary
'Set ShapeDict = ShapesToDict(iFamily)
'Dim key As Variant
'
'For Each key In MembersDict.Keys
'    If Not ShapeDict.Exists(key) Then
'        Dim iBtn As New aclsItemButton
'        Set ThisItem = New aclsItem
'        ThisItem.Initialize CInt(key)
'        iBtn.AddNew ThisItem, iFamily.Family
'    End If
'Next key
'Set MembersDict = Nothing
'Set ShapeDict = Nothing
'Set iBtn = Nothing
'Set ThisItem = Nothing
'End Sub
'
'Private Function MembersToDict(Members As Variant) As Dictionary
'Dim i As Integer
'Dim dict As Dictionary
'Set dict = New Dictionary
'For i = 0 To UBound(Members, 2)
'    dict.Add Members(0, i), i + 1
'Next i
'Set MembersToDict = dict
'End Function
'
'Private Function ShapesToDict(iFamily As zclsFamily) As Dictionary
'Dim i As Integer
'Dim Shapes As GroupShapes
'Set Shapes = iFamily.BtnLocation.Shapes("grpgui" & iFamily.Family).GroupItems
'Dim dict As Dictionary
'Set dict = New Dictionary
'For i = 1 To Shapes.Count
'    dict.Add CInt(Shapes(i).name), 1
'Next i
'Set ShapesToDict = dict
'End Function
'
''Public Sub bFormat(iItemBtn As aclsItemButton)
''iItemBtn.Shape.Height = iItemBtn.StyleDict("BtnHeight")
''iItemBtn.Shape.Width = iItemBtn.StyleDict("BtnWidth")
''iItemBtn.Shape.Fill.ForeColor.RGB = LongToRgb(iItemBtn.StyleDict("BtnBackRgb"))
''iItemBtn.Shape.line.ForeColor.RGB = LongToRgb(iItemBtn.StyleDict("BtnLineRgb"))
''iItemBtn.Shape.TextFrame.Characters.Font.color = iItemBtn.StyleDict("BtnTextRgb")
''iItemBtn.Shape.TextFrame.Characters.Font.name = iItemBtn.StyleDict("BtnFont")
''iItemBtn.Shape.TextFrame.Characters.Font.Size = iItemBtn.StyleDict("BtnFontSize")
''iItemBtn.Shape.Glow.color.RGB = LongToRgb(iItemBtn.StyleDict("BtnGlowRgb"))
''iItemBtn.Shape.Glow.Radius = iItemBtn.StyleDict("BtnGlowRadius")
''iItemBtn.Shape.Glow.Transparency = iItemBtn.StyleDict("BtnGlowTrans")
''
''End Sub
'Public Sub Format(StyleDict As Dictionary, iItemButton As aclsItemButton)
'iItemButton.Shape.Height = StyleDict("BtnHeight")
'iItemButton.Shape.Width = StyleDict("BtnWidth")
'iItemButton.Shape.Fill.ForeColor.RGB = LongToRgb(StyleDict("BtnBackRgb"))
'iItemButton.Shape.line.ForeColor.RGB = LongToRgb(StyleDict("BtnLineRgb"))
'iItemButton.Shape.TextFrame.Characters.Font.color = StyleDict("BtnTextRgb")
'iItemButton.Shape.TextFrame.Characters.Font.name = StyleDict("BtnFont")
'iItemButton.Shape.TextFrame.Characters.Font.Size = StyleDict("BtnFontSize")
'iItemButton.Shape.Glow.color.RGB = LongToRgb(StyleDict("BtnGlowRgb"))
'iItemButton.Shape.Glow.Radius = StyleDict("BtnGlowRadius")
'iItemButton.Shape.Glow.Transparency = StyleDict("BtnGlowTrans")
'
'End Sub
'
'Public Sub NameShape(ByRef Shape As Shape, ShpName As String)
'Shape.name = ShpName
'End Sub
'
'Public Sub SetCaption(ByRef Shape As Shape, ShpCaption As String)
'Shape.TextFrame.Characters.text = ShpCaption
'End Sub
'
''Public Function bGetButtonRow(iItemBtn As aclsItemButton) As Integer
''Dim Counter As Integer: Counter = 1
''Dim row As Integer: row = 1
''Dim ThisShapeIndex As Integer: ThisShapeIndex = iItemBtn.LocalGroup
''Do Until Counter = ThisShapeIndex
''    Counter = Counter + 1
''    row = row + 1
''    If row = (iItemBtn.StyleDict("BtnsPerColumn") + 1) Then
''        row = 1
''    End If
''Loop
''bGetButtonRow = row
''End Function
'
'Public Function GetButtonRow(iFamily As zclsFamily, ThisShape As Integer) As Integer
'Dim Counter As Integer
'Dim row As Integer
'row = 1
'Counter = 1
'Do Until Counter = ThisShape
'    Counter = Counter + 1
'    row = row + 1
'    If row = (iFamily.StyleDict("BtnsPerColumn") + 1) Then
'        row = 1
'    End If
'Loop
'GetButtonRow = row
'End Function
'
''Public Function bGetButtonColumn(iItemBtn As aclsItemButton) As Integer
''Dim Counter As Integer: Counter = 1
''Dim Column As Integer: Column = 1
''Dim ShapeCount As Integer: ShapeCount = 1
''Dim ThisShapeIndex As Integer: ThisShapeIndex = iItemBtn.LocalGroup
''Do
''    If ShapeCount = (iItemBtn.StyleDict("BtnsPerColumn") + 1) Then
''        Column = Column + 1
''        ShapeCount = 1
''    End If
''    Counter = Counter + 1
''    ShapeCount = ShapeCount + 1
''Loop Until Counter > ThisShapeIndex
''bGetButtonColumn = Column
''End Function
'
'Public Function GetButtonColumn(iFamily As zclsFamily, ThisShape As Integer) As Integer
'Dim Counter As Integer
'Dim Column As Integer
'Dim ShapeCount As Integer
'Column = 1
'Counter = 1
'ShapeCount = 1
'Do
'    If ShapeCount = (iFamily.StyleDict("BtnsPerColumn") + 1) Then
'        Column = Column + 1
'        ShapeCount = 1
'    End If
'    Counter = Counter + 1
'    ShapeCount = ShapeCount + 1
'Loop Until Counter > ThisShape
'GetButtonColumn = Column
'End Function
'
'Private Function GetLocation(Location As String) As Worksheet
'Dim wksht As Worksheet
'For Each wksht In ActiveWorkbook.Worksheets
'    If wksht.CodeName = Location Then
'        Set GetLocation = wksht
'        Exit Function
'    End If
'Next wksht
'Set wksht = Nothing
'End Function
'
'Public Sub Remove(Family As String, ItemID As Integer)
'Dim iFamily As New zclsFamily
'
'Set iFamily = iFamily.GetFamilyButtonObj(Family)
'iFamily.BtnLocation.Shapes(CStr(ItemID)).Delete
'
'If iFamily.Count = 0 Then
'    iFamily.RemoveFromGUI iFamily
'    UpdateMenuButtons iFamily.FamilyGroup
'    Set iFamily = Nothing
'    Exit Sub
'End If
'PositionAll iFamily
'Set iFamily = Nothing
'End Sub
'
'Public Sub RemoveArray(Family As String, ItemArray As Variant)
'If Not IsEmpty(ItemArray) Then
'Dim iFamily As New zclsFamily
'Set iFamily = iFamily.GetFamilyButtonObj(Family)
'
'Dim i As Integer
'For i = 0 To UBound(ItemArray)
'    iFamily.BtnLocation.Shapes(CStr(ItemArray(0, i))).Delete
'Next i
'On Error Resume Next
'    iFamily.BtnLocation.Shapes(Family & "Blank").Delete
'On Error GoTo 0
'
'Set iFamily = Nothing
'End If
'End Sub




