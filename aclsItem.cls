VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pItemID As Integer
Private pItemName As String
Private pPrice As Currency
Private pFamily As String
Private pReq1 As String
Private pReq2 As String
Private pIsPrimaryItem As Boolean
Private pCategory As String
Private pCoolerFlag As Boolean
Private pAlwaysTax As Boolean
Private pAltScePrice As Boolean
Private pAltPrice As Currency
Private pDiscountPrice As Double
Private pPrintKitchen As Boolean
Private pPrintPantry As Boolean
Private pItemOptions As Boolean
Private pClassCode As Integer
Private pSeat As Integer
Private pItemIndicator As String
Private pItemType As Variant
Private pParentID As Integer
Private pChildID As Integer
Private pCollID As Integer
Private pComponents As Collection
Private pParent As bclsParent
Private pChild As bclsChild
Private pChildren As bclsChildren
Private pRequiredComponents As Dictionary
Private pOrderRank As Variant



Public Property Let ItemID(value As Integer)
pItemID = value
End Property
Public Property Get ItemID() As Integer
ItemID = pItemID
End Property
Public Property Let ItemName(value As String)
pItemName = value
End Property
Public Property Get ItemName() As String
ItemName = pItemName
End Property
Public Property Let Price(value As Currency)
pPrice = value
End Property
Public Property Get Price() As Currency
Price = pPrice
End Property
Public Property Let Family(value As String)
pFamily = value
End Property
Public Property Get Family() As String
Family = pFamily
End Property
Public Property Let Req1(value As String)
pReq1 = value
End Property
Public Property Get Req1() As String
Req1 = pReq1
End Property
Public Property Let Req2(value As String)
pReq2 = value
End Property
Public Property Get Req2() As String
Req2 = pReq2
End Property
Public Property Let IsPrimaryItem(value As Boolean)
pIsPrimaryItem = value
End Property
Public Property Get IsPrimaryItem() As Boolean
IsPrimaryItem = pIsPrimaryItem
End Property
Public Property Let Category(value As String)
pCategory = value
End Property
Public Property Get Category() As String
Category = pCategory
End Property
Public Property Let CoolerFlag(value As Boolean)
pCoolerFlag = value
End Property
Public Property Get CoolerFlag() As Boolean
CoolerFlag = pCoolerFlag
End Property
Public Property Let AlwaysTax(value As Boolean)
pAlwaysTax = value
End Property
Public Property Get AlwaysTax() As Boolean
AlwaysTax = pAlwaysTax
End Property
Public Property Let AltScePrice(value As Boolean)
pAltScePrice = value
End Property
Public Property Get AltScePrice() As Boolean
AltScePrice = pAltScePrice
End Property
Public Property Let AltPrice(value As Currency)
pAltPrice = value
End Property
Public Property Get AltPrice() As Currency
AltPrice = pAltPrice
End Property
Public Property Let DiscountPrice(value As Currency)
pDiscountPrice = value
End Property
Public Property Get DiscountPrice() As Currency
DiscountPrice = pDiscountPrice
End Property
Public Property Let ItemOptions(value As Boolean)
pItemOptions = value
End Property
Public Property Get ItemOptions() As Boolean
ItemOptions = pItemOptions
End Property
Public Property Let PrintKitchen(value As Boolean)
pPrintKitchen = value
End Property
Public Property Get PrintKitchen() As Boolean
PrintKitchen = pPrintKitchen
End Property
Public Property Let PrintPantry(value As Boolean)
pPrintPantry = value
End Property
Public Property Get PrintPantry() As Boolean
PrintPantry = pPrintPantry
End Property
Public Property Let ClassCode(value As Integer)
pClassCode = value
End Property
Public Property Get ClassCode() As Integer
ClassCode = pClassCode
End Property
Public Property Let seat(value As Integer)
pSeat = value
End Property
Public Property Get seat() As Integer
seat = pSeat
End Property
Public Property Let ItemIndicator(value As String)
pItemIndicator = value
End Property
Public Property Get ItemIndicator() As String
ItemIndicator = Me.ItemType.GetItemIndicator
End Property

Public Property Let ItemType(value As Variant)
Set pItemType = value
End Property
Public Property Get ItemType() As Variant
Set ItemType = pItemType
End Property

Public Property Let ParentID(value As Integer)
pParentID = value
End Property
Public Property Get ParentID() As Integer
ParentID = pParentID
End Property

Public Property Let CollID(value As Integer)
pCollID = value
End Property
Public Property Get CollID() As Integer
CollID = pCollID
End Property

Public Property Let ChildID(value As Integer)
pChildID = value
End Property
Public Property Get ChildID() As Integer
ChildID = pChildID
End Property

Public Property Let Components(value As Collection)
Set pComponents = value
End Property
Public Property Get Components() As Collection
Set Components = pComponents
End Property

Public Property Let Parent(value As bclsParent)
Set pParent = value
End Property
Public Property Get Parent() As bclsParent
Set Parent = pParent
End Property

Public Property Let Children(value As bclsChildren)
Set pChildren = value
End Property
Public Property Get Children() As bclsChildren
Set Children = pChildren
End Property
Public Property Let RequiredComponents(value As Dictionary)

Set pRequiredComponents = value
End Property
Public Property Get RequiredComponents() As Dictionary

Set RequiredComponents = pRequiredComponents
End Property

Public Property Let OrderRank(value As Variant)
Set pOrderRank = value
End Property
Public Property Get OrderRank() As Variant
Set OrderRank = pOrderRank
End Property

Public Sub Initialize(ItemID As Integer)
Dim ItemType As Variant
Set ItemType = GetItemType(CStr(GetItemClassCode(ItemID)))
ItemType.ItemInitialize ItemID



ThisItem.ItemType = ItemType
Set ThisItem = ThisItem.CreateNew(ThisItem.ItemType.ItemID)
Set ItemType = Nothing
End Sub
Public Function CreateNew(ItemID As Integer) As aclsItem
Dim item As New aclsItem
Set item = PopItem(ItemID, item)
item.ItemType = ThisItem.ItemType
Set CreateNew = item
Set item = Nothing
End Function

Public Function GetRequiredComponents() As Dictionary
Dim dict As New Dictionary
If Not Me.Req1 = "" Then
    dict.Add Me.Req1, Me.Req1
End If
If Not Me.Req2 = "" Then
    dict.Add Me.Req2, Me.Req2
End If
Set GetRequiredComponents = dict
End Function

Public Function ValidateRequirements(item As aclsItem, Requirement As String) As Boolean
Dim member As aclsItem
For Each member In CItem
    If member.Parent.ID = item.CollID Then
        If member.Family = Requirement Then
            ValidateRequirements = True
            Exit Function
        End If
    End If
Next member
ValidateRequirements = False
End Function

Public Function UnassignChild(ID As Integer)
Dim child As bclsChild
For Each child In Me.Children.coll
    If child.ID = ID Then
        Me.Children.coll.Remove CStr(ID)
    End If
Next child
End Function

Public Function AssignOrderRank() As Variant
If CItem.Count > 0 Then
    Dim iChild As New bclsChild
    Set AssignOrderRank = iChild
    Set iChild = Nothing
    Exit Function
End If

Dim iParent As New bclsParent
Set AssignOrderRank = iParent
Set iParent = Nothing
End Function

