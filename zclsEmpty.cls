VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsEmpty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Private pFamily As String
'
'Public Property Let Family(value As String)
'pFamily = value
'End Property
'Public Property Get Family() As String
'Family = pFamily
'End Property
'
'Public Sub PopulateCItem()
'Dim arr As Variant
'Dim iFamily As New zclsFamily
'Dim iDataObj As aclsDataObject
'Set iDataObj = iFamily.Wrap(GetNewMatchObj("FamilyGroup", "Component"))
'
'arr = FilteredMatch(iDataObj, "Family")
'Dim EmptyObj As zclsEmpty
'Dim key As Variant
'Dim i As Integer
'For i = 1 To UBound(arr)
'    Set EmptyObj = New zclsEmpty
'    EmptyObj.Family = "Empty"
'    CItem.Add EmptyObj, arr(i)(0, 0)
'Next i
'Set EmptyObj = Nothing
'End Sub
