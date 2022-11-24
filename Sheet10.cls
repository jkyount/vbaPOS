VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()

Sheet10.ScrollArea = "A1:R41"
End Sub

Public Sub Worksheet_SelectionChange(ByVal Target As range)
 Dim LoopInterrupt As Boolean
 'If LoopInterrupt = True Then
'    Exit Sub
'End If
 'SetClickCount 0
Dim RangeLine As Integer, CollLine As Integer
RangeLine = Target.row - GetRangeOffset
Dim InORange As Boolean, InSRange As Boolean

Dim ORange As range, SRange As range
Set ORange = Sheet10.range("OriginalCheckRange")
Set SRange = Sheet10.range("SplitCheckRange")

Dim coll As New Collection
If Not Intersect(ORange, Target) Is Nothing Then
    InORange = True
    Set coll = GetcollOriginal
End If
If Not Intersect(SRange, Target) Is Nothing Then
    InSRange = True
    Set coll = GetcollSplit
End If
        
If GetClickCount = 0 Then
    If InORange = False And InSRange = False Then
        Exit Sub
        Set coll = Nothing
        Set ORange = Nothing
        Set SRange = Nothing
    End If
    If coll.Count = 0 Then
        Set coll = Nothing
        Set ORange = Nothing
        Set SRange = Nothing
        Exit Sub
    End If
    
    If RangeLine > coll(coll.Count).GuiRow Then
        GoTo EH
        Exit Sub
    End If
    
   
    
    CollLine = GetCollRow(coll, RangeLine)

    If CollLine = -1 Then
        Exit Sub
    End If
    SetClickCount 1
     SetClickRange Target
    ORange.Interior.color = RGB(255, 255, 255)
    SRange.Interior.color = RGB(255, 255, 255)
    
    HighlightEntityGroup Target, coll
    Exit Sub
End If


'
If GetClickCount > 0 Then
    
    If InSRange = False And InORange = False Then
        GoTo EH:
    End If
    
    If InORange = True _
    And Not Intersect(ORange, GetClickRange) Is Nothing _
    Then

        If RangeLine > coll(coll.Count).GuiRow Then
            GoTo EH:
            Exit Sub
        End If
'**********  SET CLICK RANGE AFTER CHECKING CollLine = -1???
        
        CollLine = GetCollRow(coll, RangeLine)
        If CollLine = -1 Then
            Exit Sub
        End If
        SetClickRange Target
        ORange.Interior.color = RGB(255, 255, 255)
        SRange.Interior.color = RGB(255, 255, 255)
        HighlightEntityGroup Target, coll
        Exit Sub
        
    End If

    If InSRange = True _
    And Not Intersect(SRange, GetClickRange) Is Nothing _
    Then
        If RangeLine > coll(coll.Count).GuiRow Then
            GoTo EH:
            Exit Sub
        End If
        
        
        CollLine = GetCollRow(coll, RangeLine)
        If CollLine = -1 Then
            Exit Sub
        End If
        SetClickRange Target
        ORange.Interior.color = RGB(255, 255, 255)
        SRange.Interior.color = RGB(255, 255, 255)
        HighlightEntityGroup Target, coll
        Exit Sub
    End If


If Not Intersect(ORange, Target) Is Nothing Then
        Set coll = GetcollSplit
        
        SplitItem GetEntityGroup(GetClickRange.row, coll), currentcheck
        ORange.Interior.color = RGB(255, 255, 255)
        SRange.Interior.color = RGB(255, 255, 255)
        SetClickCount 0
  
        Exit Sub
    End If
'
    If Not Intersect(SRange, Target) Is Nothing Then
        Set coll = GetcollOriginal
       
        SplitItem GetEntityGroup(GetClickRange.row, coll), PeekNextCheck
        ORange.Interior.color = RGB(255, 255, 255)
        SRange.Interior.color = RGB(255, 255, 255)
        SetClickCount 0
   
        Exit Sub
    End If
End If


EH:
SetClickCount 0
ORange.Interior.color = RGB(255, 255, 255)
SRange.Interior.color = RGB(255, 255, 255)
Set coll = Nothing
Set ORange = Nothing
Set SRange = Nothing
End Sub


