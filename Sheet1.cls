VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit






Private Sub Worksheet_Activate()

Sheet1.ScrollArea = "A1:K30"
End Sub

' Private Sub SpecialInstructionText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        If Not SpecialInstructionText.value = "" Then
'            Ssfub
'        End If
'    End If
'End Sub
'
'
' Private Sub CustomItemName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        If Not CustomItemName.value = "" And Not CustomItemPrice.value = "" Then
'            Ssfub
'        End If
'    End If
'End Sub



