VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsCheckLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private pRow As Integer
Private pPrimary As Boolean
Private pData() As Variant
Private pPrintKitchen As Boolean
Private pPrintPantry As Boolean
Private pSeat As Integer
Private pGuiRow As Integer
Private pCashPayment As Double
Private pChargePayment As Double
Private pGiftCertPayment As Double
Private pLocalGroup As Integer
Private pCoolerFlag As Boolean
Private pItemID As Integer
Private pEntityGroup As Integer



Public dbCheck As New ADODB.Connection
Public rsCheck As New ADODB.RecordSet


Public Property Get Db() As String
Db = GetDb
End Property
Public Property Let row(value As Integer)
pRow = value
End Property
Public Property Get row() As Integer
row = pRow
End Property
Public Property Let GuiRow(value As Integer)
pGuiRow = value
End Property
Public Property Get GuiRow() As Integer
GuiRow = pGuiRow
End Property
Public Property Let seat(value As Integer)
pSeat = value
End Property
Public Property Get seat() As Integer
seat = pSeat
End Property
Public Property Let primary(value As Boolean)
pPrimary = value
End Property
Public Property Get primary() As Boolean
primary = pPrimary
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
Public Property Let Data(value As Variant)
pData = value
End Property
Public Property Get Data() As Variant
Data = pData
End Property
Public Property Let LocalGroup(value As Integer)
pLocalGroup = value
End Property
Public Property Get LocalGroup() As Integer
LocalGroup = pLocalGroup
End Property
Public Property Let CoolerFlag(value As Boolean)
pCoolerFlag = value
End Property
Public Property Get CoolerFlag() As Boolean
CoolerFlag = pCoolerFlag
End Property
Public Property Let ItemID(value As Integer)
pItemID = value
End Property
Public Property Get ItemID() As Integer
ItemID = pItemID
End Property
Public Property Let EntityGroup(value As Integer)
pEntityGroup = value
End Property
Public Property Get EntityGroup() As Integer
EntityGroup = pEntityGroup
End Property


Public Function Wrap(obj As aclsDataObject) As aclsDataObject
Dim iDataObj As New aclsDataObject
Set iDataObj = obj
iDataObj.rs = Me.GetRs
iDataObj.Conn = Me.GetConn
iDataObj.Db = Me.GetDb
iDataObj.DbFile = Me.GetDbFile
Set Wrap = iDataObj
Set iDataObj = Nothing
End Function

Public Function GetDb() As String
GetDb = "DailyCheckDetail"
End Function
Public Function GetDbFile() As String
GetDbFile = "CheckDb"
End Function

Public Function GetConn() As ADODB.Connection
Set GetConn = dbCheck
End Function

Public Function GetRs() As ADODB.RecordSet
Set GetRs = rsCheck
End Function

Public Sub OpenDbs()
Dim iDataObject As New aclsDataObject
Set iDataObject = Wrap(iDataObject)
iDataObject.OpenDbs iDataObject
Set iDataObject = Nothing
End Sub

Public Sub CloseDbs()
Dim iDataObject As New aclsDataObject
Set iDataObject = Wrap(iDataObject)
iDataObject.CloseDbs iDataObject
Set iDataObject = Nothing
End Sub



'==========================================================================


Public Function DefineWriteLines(coll As Collection, WriteLines As Variant) As Collection
Set DefineWriteLines = pDefineWriteLines(coll, WriteLines)
End Function
Public Function pDefineWriteLines(coll As Collection, WriteLines As Variant) As Collection

Dim i As Integer
Dim arr() As Variant
Dim arr2() As Variant
Dim member As Object
For Each member In coll
    arr2 = member.Data
    ReDim arr(0 To UBound(WriteLines))
    For i = 0 To UBound(WriteLines)
        arr(i) = arr2(WriteLines(i))
    Next i
    member.Data = arr
Next member

Set pDefineWriteLines = coll
End Function

Public Function CreateNew(row As Integer, ValueDict As Dictionary) As zclsCheckLines
Dim x As New zclsCheckLines
Dim TempArray(1 To 4) As Variant
  x.row = row
  x.ItemID = ValueDict("ItemID")
 TempArray(1) = ValueDict("ItemIndicator")
 TempArray(2) = ValueDict("ItemName")
 TempArray(3) = ValueDict("Price")
 TempArray(4) = ValueDict("ItemID")
 x.PrintKitchen = ValueDict("PrintKitchen")
 x.PrintPantry = ValueDict("PrintPantry")
 x.seat = ValueDict("Seat")
 x.LocalGroup = ValueDict("LocalGroup")
 x.primary = ValueDict("IsPrimaryItem")
 x.EntityGroup = ValueDict("EntityGroup")
 x.Data = TempArray
Set CreateNew = x
End Function
Public Function GetCheckLines(check As String, ValueDict As Collection) As Collection
Set GetCheckLines = pGetCheckLines(check, ValueDict)
End Function
Public Function pGetCheckLines(check As String, ValueDict As Collection) As Collection
'If Not IsEmpty(CheckArray) Then

Dim line As New zclsCheckLines
Dim coll As New Collection
If Not ValueDict(1).Count = 0 Then
Dim i As Integer
For i = 1 To ValueDict.Count
     Set line = line.CreateNew(coll.Count + 1, ValueDict(i))
     coll.Add line, ("Line" & i)
Next i
End If
Set pGetCheckLines = coll
Set coll = Nothing
Set ValueDict = Nothing
End Function

Public Sub WriteCheckLines(range As range, coll As Collection)
Dim i As Integer, SeatLines As Integer, GuiRow As Integer, seat As Integer
Dim seatarray(1 To 3) As String
If coll.Count = 0 Then
    range.value = ""
    'MsgBox "Attempted to write an empty collection."
    Exit Sub
End If
range.value = ""
SeatLines = 1
seat = coll("Line1").seat
GuiRow = 2
seatarray(1) = "- - - - - - -"
seatarray(2) = "- - - - - - - Seat " & seat & "- - - - - - - - -"
seatarray(3) = "- - - - - - -"
range.Rows(1).value = seatarray()
For i = 1 To coll.Count
    If Not coll("Line" & i).seat = seat Then
        seat = coll("Line" & i).seat
        seatarray(2) = "- - - - - - - Seat " & seat & "- - - - - - - - -"
        range.Rows(i + SeatLines).value = seatarray()
        SeatLines = SeatLines + 1
    End If
    range.Rows(i + SeatLines).value = coll("Line" & i).Data()
Next i

End Sub



