Attribute VB_Name = "i_Family"
Option Explicit

Public Function GetFamilyObj(Family As String) As zclsFamily
Dim iFamily As zclsFamily
Set iFamily = New zclsFamily
iFamily.Family = Family
Set GetFamilyObj = iFamily
Set iFamily = Nothing
End Function

'Public Function GetFamilyButtonObj(Family As String) As zclsFamily
'Dim iFamily As zclsFamily
'Set iFamily = New zclsFamily
'Set GetFamilyButtonObj = iFamily.GetFamilyButtonObj(Family)
'Set iFamily = Nothing
'End Function


