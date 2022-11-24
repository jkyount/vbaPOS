Attribute VB_Name = "xGlobalVars"
Public CItem As New Collection
Public currentcheck As String
Public ItemCell As range
Public ThisOrder As New aclsOrder
Public ThisTable As New zclsTable
Public ThisEmployee As New zclsEmployee
Public EndOfCheck As Integer
Public collCheckData As New Collection
'Public collArchivedCheckData As New Collection
Public LastEntry As New Collection
Public ThisItem As aclsItem
Public CurrentParent As Integer




'------------------Implement GetArchive for all database objects
'Module option for invisibility to other projects
'Config options for quickmods

'Restore state to include restoring collCheckData
'Convert checknumbers to integers


'To add new item property:
'   Databases to be modified:
'        AllItems
'        DailyCheckDetail
'        ArchiveCheckDetail
'        TempCheck
'   Functions to be modified:
'        PopItem
'        zclsCheckLines.CreateNew
'        zclsDailyCheckDetail.AddItem
        
        













Public Sub SetPickupTotalCells()
Set PaySubtotalB = Worksheets("Payment").range("SubTotal")
Set PayTax2 = Worksheets("Payment").range("Tax")
Set PayCheckNumber = Worksheets("Payment").range("checknumbercell")
End Sub



Public Function GetSplitCheck() As String
GetSplitCheck = GetNextCheck
End Function

Public Function SetQty(value As Integer)
Qty = value
End Function




