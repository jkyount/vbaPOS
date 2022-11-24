VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsPizzaTopping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Implements zclsItem

Private pItemID As Integer


Public Property Let ItemID(value As Integer)
pItemID = value
End Property
Public Property Get ItemID() As Integer
ItemID = pItemID
End Property


Public Sub ItemInitialize(ItemID As Integer)
Me.ItemID = ItemID

End Sub
Public Sub SpclConfig(item As aclsItem)
ThisItem.Price = GetPrice
If TypeName(CItem("1").ItemType) = "zclsPizza" Then
    If CItem("1").ItemType.ToppingArea = 1 Then
        ThisItem.ItemName = "1/2 " & ThisItem.ItemName
    End If
End If

End Sub

Public Sub InheritParentPrintParams(item As aclsItem)
End Sub

Private Function GetPrice() As Currency
If Not TypeName(CItem("1").ItemType) = "zclsPizza" Then
    GetPrice = 2
    Exit Function
End If

If CItem("1").ItemName = "Large" Then
        If CItem("1").ItemType.ToppingArea = 0 Then
            GetPrice = 2
        End If
        If CItem("1").ItemType.ToppingArea = 1 Then
            GetPrice = 1.25
        End If
End If
If CItem("1").ItemName = "Small" Then
        If CItem("1").ItemType.ToppingArea = 0 Then
            GetPrice = 1.5
        End If
        If CItem("1").ItemType.ToppingArea = 1 Then
            GetPrice = 1
        End If
End If
If ThisItem.ItemName = "Plain" Then
    GetPrice = 0
End If
End Function



Public Sub zclsItem_RefreshPreviewWindow()
RefreshPreviewWindow
End Sub

Public Sub zclsItem_UpdateGUI()
End Sub

Public Function GetItemIndicator() As String
GetItemIndicator = ""
End Function
