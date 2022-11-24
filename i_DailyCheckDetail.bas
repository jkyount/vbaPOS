Attribute VB_Name = "i_DailyCheckDetail"
Option Explicit

Public Function GetDailyCheckDetailObj() As zclsDailyCheckDetail
Set GetDailyCheckDetailObj = New zclsDailyCheckDetail
End Function

Public Function IsValidCheck(check As String) As Boolean

Dim iDailyCheckDetail As New zclsDailyCheckDetail
IsValidCheck = iDailyCheckDetail.IsValidCheck(check)
Set iDailyCheckDetail = Nothing

End Function


Public Sub SendCurrentItems(check As String)

Dim iDailyCheckDetail As New zclsDailyCheckDetail
iDailyCheckDetail.SendCurrentItems check
Set iDailyCheckDetail = Nothing

End Sub


Public Function GetNextLocalGroup(check As String) As Integer
Dim iDailyCheckDetail As New zclsDailyCheckDetail
GetNextLocalGroup = iDailyCheckDetail.GetNextLocalGroup(check)
Set iDailyCheckDetail = Nothing
End Function

Public Function GetNextEntityGroup(check As String) As Integer
Dim iDailyCheckDetail As New zclsDailyCheckDetail
GetNextEntityGroup = iDailyCheckDetail.GetNextEntityGroup(check)
Set iDailyCheckDetail = Nothing
End Function

Public Sub RemoveItem(check As String, EntityGroup As Integer)
Dim iDailyCheckDetail As New zclsDailyCheckDetail
iDailyCheckDetail.RemoveItem check, EntityGroup
Set iDailyCheckDetail = Nothing
End Sub

Public Sub CopyToTemp(check As String)
Dim iDailyCheckDetail As New zclsDailyCheckDetail
iDailyCheckDetail.CopyToTemp check
Set iDailyCheckDetail = Nothing
End Sub

Public Sub AppendToTemp(check As String)
Dim iDailyCheckDetail As New zclsDailyCheckDetail
iDailyCheckDetail.AppendToTemp check
Set iDailyCheckDetail = Nothing
End Sub
Public Sub SubmitChanges(check As String)
Dim iDailyCheckDetail As New zclsDailyCheckDetail
iDailyCheckDetail.SubmitChanges check
Set iDailyCheckDetail = Nothing
End Sub








