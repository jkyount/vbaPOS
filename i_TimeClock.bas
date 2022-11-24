Attribute VB_Name = "i_TimeClock"
Option Explicit

Public Sub TimeClock_ClockIn(ID As Long)
Dim iTimeClock As New aclsTimeClock
iTimeClock.ClockIn ID
Set iTimeClock = Nothing
End Sub

Public Sub TimeClock_ClockOut(ID As Long)
Dim iTimeClock As New aclsTimeClock
iTimeClock.ClockOut ID
Set iTimeClock = Nothing
End Sub
