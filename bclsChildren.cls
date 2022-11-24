VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bclsChildren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColl As New Collection


Public Property Let coll(value As Collection)
Set pColl = value
End Property
Public Property Get coll() As Collection
Attribute coll.VB_UserMemId = 0
Set coll = pColl
End Property

Public Function GetNew() As bclsChildren
Dim x As New bclsChildren
Set GetNew = x
Set x = Nothing
End Function



