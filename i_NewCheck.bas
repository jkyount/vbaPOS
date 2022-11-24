Attribute VB_Name = "i_NewCheck"
Option Explicit

Public Function GetNextCheck() As String
Dim iNewCheck As New zclsNewCheck
GetNextCheck = iNewCheck.GetNextCheck
Set iNewCheck = Nothing
End Function

Public Function PeekNextCheck() As String
Dim iNewCheck As New zclsNewCheck
PeekNextCheck = iNewCheck.PeekNextCheck
Set iNewCheck = Nothing
End Function

Public Function InitializeCheck(check As String) As String
Dim iNewCheck As New zclsNewCheck
iNewCheck.InitializeCheck (check)
InitializeCheck = check
Set iNewCheck = Nothing
End Function

Public Function SetCheckInUse(check As String)
Dim iNewCheck As New zclsNewCheck
iNewCheck.SetCheckInUse check
Set iNewCheck = Nothing
End Function

Public Function SetCheckUnused(check As String)
Dim iNewCheck As New zclsNewCheck
iNewCheck.SetCheckUnused check
Set iNewCheck = Nothing
 
End Function


