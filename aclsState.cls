VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get check() As String
check = GetStateValue("State_currentcheck")
End Property
Public Property Get seat() As String
seat = GetStateValue("State_currentseat")
End Property
Public Property Get EmployeeID() As Long
EmployeeID = GetStateValue("State_IDNumber")
End Property
Public Property Get OrderName() As String
OrderName = GetStateValue("State_OrderName")
End Property
Public Property Get Table() As String
Table = GetStateValue("State_ParentTable")
End Property


Private Function GetStateValue(value As String) As Variant
GetStateValue = Sheet5.range(value).value
End Function
Public Sub CaptureState()
Dim sheet As Worksheet
Set sheet = Sheet5
With sheet
    .range("State_currentcheck").value = currentcheck
    .range("State_currentseat").value = GetCurrentSeat
    .range("State_IDNumber").value = ThisEmployee.IDNumber
    .range("State_OrderName").value = ThisOrder.ValueDict("OrderName")
    .range("State_ParentTable").value = ThisOrder.ValueDict("Table")
End With
Set sheet = Nothing
End Sub

Public Sub RestoreState()
Dim sheet As Worksheet
Set sheet = Sheet5
With sheet
currentcheck = .range("State_currentcheck").value
SetCurrentSeat .range("State_currentseat").value
ThisEmployee.IDNumber = .range("State_IDNumber").value
ThisTable.Table = .range("State_OrderName").value
ThisTable.ParentTable = .range("State_ParentTable").value
End With
Set sheet = Nothing



End Sub




