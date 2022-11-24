Attribute VB_Name = "i_CheckLines"
Option Explicit

Public Sub WriteCheckLines(rg As range, coll As Collection)
Dim iCheckLines As New zclsCheckLines
iCheckLines.WriteCheckLines rg, coll
Set iCheckLines = Nothing
End Sub



Public Function GetCheckLines(check As String, ValueDict As Collection) As Collection
Dim iCheckLines As New zclsCheckLines
Set GetCheckLines = iCheckLines.GetCheckLines(check, ValueDict)
Set iCheckLines = Nothing
End Function

Public Function DefineWriteLines(coll As Collection, ParamArray WriteLines() As Variant) As Collection
Dim iCheckLines As New zclsCheckLines
Dim arr() As Variant
arr = WriteLines
Set DefineWriteLines = iCheckLines.DefineWriteLines(coll, arr)
Set iCheckLines = Nothing
End Function


