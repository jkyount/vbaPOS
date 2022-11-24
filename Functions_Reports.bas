Attribute VB_Name = "Functions_Reports"
Option Explicit

Public FormatDict As New Dictionary

Public Sub PopFormatDict(item As Variant, ItemName As String)
FormatDict.Add ItemName, item
End Sub

Public Sub RemoveFormatDict(key As String)
FormatDict.Remove key
End Sub

Public Function GetFormatDict() As Dictionary
Set GetFormatDict = FormatDict
End Function

Public Sub ClearFormatDict()
FormatDict.RemoveAll
End Sub

Public Function GetDailyDataObj() As dclsDailyData
Set GetDailyDataObj = New dclsDailyData
End Function
Public Function GetArchiveDataObj() As dclsArchiveData
Set GetArchiveDataObj = New dclsArchiveData
End Function

Public Function ValidateRequiredParams(RequiredParams As Dictionary, SelectedParams As Dictionary)
Dim key As Variant
For Each key In RequiredParams.Keys
    If Not SelectedParams.Exists(key) Then
        MsgBox "Missing" & key
        Exit Function
    End If
Next key
End Function


