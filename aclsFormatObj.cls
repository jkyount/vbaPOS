VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aclsFormatObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pFmtBold As Boolean
Private pFmtFontSize As Integer
Private pFmtAlign As Variant
Private pFmtBorderWeight As Integer
Private pFmtBorderStyle As Variant
Private pData As Variant


Public Property Let FmtBold(value As Boolean)
pFmtBold = value
End Property
Public Property Get FmtBold() As Boolean
FmtBold = pFmtBold
End Property

Public Property Let FmtFontSize(value As Integer)
pFmtFontSize = value
End Property
Public Property Get FmtFontSize() As Integer
FmtFontSize = pFmtFontSize
End Property

Public Property Let FmtAlign(value As Variant)
pFmtAlign = value
End Property
Public Property Get FmtAlign() As Variant
FmtAlign = pFmtAlign
End Property

Public Property Let FmtBorderWeight(value As Integer)
pFmtBorderWeight = value
End Property
Public Property Get FmtBorderWeight() As Integer
FmtBorderWeight = pFmtBorderWeight
End Property


Public Property Let FmtBorderStyle(value As Variant)
pFmtBorderStyle = value
End Property
Public Property Get FmtBorderStyle() As Variant
FmtBorderStyle = pFmtBorderStyle
End Property

Public Property Let Data(value As Variant)
pData = value
End Property
Public Property Get Data() As Variant
Data = pData
End Property

Public Function GetNewFormatObj(FmtBold As Boolean, FmtFontSize As Integer, FmtAlign As Variant, FmtBorderWeight As Integer, FmtBorderStyle As Variant) As aclsFormatObj

Dim x As New aclsFormatObj
x.FmtBold = FmtBold
x.FmtFontSize = FmtFontSize
x.FmtAlign = FmtAlign
x.FmtBorderWeight = FmtBorderWeight
x.FmtBorderStyle = FmtBorderStyle
Set GetNewFormatObj = x
Set x = Nothing
End Function



