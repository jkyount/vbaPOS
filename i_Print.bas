Attribute VB_Name = "i_Print"
Option Explicit


Public Sub PrintGuestCheck(check As String)

Dim iPrint As New aclsPrint
iPrint.PrintCheck GetGuestCheckObject(check)
Set iPrint = Nothing

End Sub

Public Sub PrintPrepCheck(check As String)

Dim iPrint As New aclsPrint
Dim iKitchen As New pclsKitchenCheck
iPrint.PrintCheck GetPrepCheckObject(check, iKitchen)
Set iPrint = Nothing

End Sub

Public Sub PrintPantryCheck(check As String)

Dim iPrint As New aclsPrint
Dim iPantry As New pclsPantryCheck
iPrint.PrintCheck GetPrepCheckObject(check, iPantry)
Set iPrint = Nothing

End Sub


Public Function GetGuestCheckObject(check As String) As aclsGuestCheck
Dim iGuestCheck As New aclsGuestCheck
Set GetGuestCheckObject = iGuestCheck.NewPrintObject(check)
Set iGuestCheck = Nothing
End Function

Public Function GetPrepCheckObject(check As String, PrintType As Variant) As aclsPrepCheck
Dim iPrepCheck As New aclsPrepCheck
Set GetPrepCheckObject = iPrepCheck.NewPrintObject(check, PrintType)
Set iPrepCheck = Nothing
End Function
