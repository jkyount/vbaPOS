Attribute VB_Name = "i_State"
Option Explicit

Public Sub CaptureState()
Dim iState As New aclsState
iState.CaptureState
Set iState = Nothing
End Sub

Public Sub RestoreState()
If ThisEmployee.IDNumber = 0 Then
    Dim iState As New aclsState
    iState.RestoreState
    Set iState = Nothing
End If
SyncValues
End Sub


