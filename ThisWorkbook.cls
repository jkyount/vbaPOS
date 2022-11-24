VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit





Private Sub Workbook_Open()
MsgBox "Integrated Mods into menu button framework.  All MenuStyles are now integrated.  " _
            & "Removed zclsItem implementation, to be replaced by bclsParent and bclsChild for ordering behavior." _
            & "zclsPrimary, NotPrimary, etc to be used for item parameters, not ordering behavior."
End Sub
