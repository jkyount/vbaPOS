Attribute VB_Name = "i_Carryout"
Option Explicit

Public Function GetNewCarryoutObj() As oclsCarryout
Dim iCarryout As New oclsCarryout
Set GetNewCarryoutObj = iCarryout
End Function
