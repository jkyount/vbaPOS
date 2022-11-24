VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public Sub PrintCheck(obj As Variant)

AssignRangeAttributeValues GetCheckAttributes(obj.check), Sheet13
obj.InitializePrintRange (obj)
obj.WritePrintLines obj.GetPrintColl(obj)
obj.AddCheckFooter obj.check

obj.FormatPrintRange
ExecutePrint obj
End Sub

Public Sub ExecutePrint(obj As Variant)
Dim PrintRg As range
Set PrintRg = obj.GetPrintArea(obj)
'PrintRg.Interior.color = rgbYellow
PrintRg.PrintOut ActivePrinter:="POS-80C (copy2)"
End Sub








