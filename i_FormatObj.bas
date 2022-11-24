Attribute VB_Name = "i_FormatObj"
Option Explicit

Public Function GetNewFormatObj(Optional FmtBold As Boolean = False, Optional FmtFontSize As Integer = 0, Optional FmtAlign As Variant = "", Optional FmtBorderWeight As Integer = 0, Optional FmtBorderStyle As Variant = "") As aclsFormatObj
Dim iFormat As New aclsFormatObj
Set GetNewFormatObj = iFormat.GetNewFormatObj(FmtBold, FmtFontSize, FmtAlign, FmtBorderWeight, FmtBorderStyle)
Set iFormat = Nothing
End Function

