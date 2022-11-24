Attribute VB_Name = "ut_Collections"
Option Explicit

Public Sub ClearCollection(coll As Collection)
If Not coll.Count = 0 Then
    Dim i As Integer
    For i = 1 To coll.Count
        coll.Remove (1)
    Next i
End If
End Sub

Public Function DuplicateCollection(coll As Collection) As Collection
Dim i As Integer
Dim NewColl As New Collection
For i = 1 To coll.Count
    NewColl.Add coll(i)
Next i
Set DuplicateCollection = NewColl
If coll.Count = 0 Then
    MsgBox "Attempted to duplicate an empty collection."
End If
End Function
Public Function DuplicateCheckLines(coll As Collection) As Collection
Dim i As Integer
Dim NewColl As New Collection
For i = 1 To coll.Count
    NewColl.Add coll(i), "Line" & i
Next i
Set DuplicateCheckLines = NewColl
If coll.Count = 0 Then
    MsgBox "Attempted to duplicate an empty collection."
End If
End Function

Public Function DuplicateCItem(coll As Collection) As Collection

Dim i As Integer
Dim NewColl As New Collection
For i = 1 To coll.Count
    NewColl.Add coll(i), CStr(coll(i).CollID)
Next i
Set DuplicateCItem = NewColl
End Function

Public Function RecallCheckLines(check As String) As Collection


Dim x As New zclsDailyCheckDetail
Dim OrderBy As String
OrderBy = "ORDER BY Seat ASC, LocalGroup ASC"
Dim iDataObj As New aclsDataObject
Set iDataObj = x.Wrap(GetNewMatchObj(, check))
Set RecallCheckLines = SortCheckLines(GetCheckLines(check, GetValueDict(iDataObj, ConstructOrderedQuery(iDataObj, OrderBy))))

Set x = Nothing
Set iDataObj = Nothing

End Function



Public Function SortCheckLines(coll As Collection) As Collection
Dim collTemp As New Collection
Dim i As Integer
Dim z As zclsCheckLines
Dim k As Integer
k = 1
If Not coll.Count = 0 Then
    For i = 1 To 12
        For Each z In coll
            If z.seat = i Then
                collTemp.Add z, ("Line" & k)
                z.row = k
                k = k + 1
            End If
        Next z
    Next i
    Dim st As Integer, GuiRow As Integer
    st = collTemp("Line1").seat
    GuiRow = 2
    For i = 1 To collTemp.Count
        If Not collTemp("Line" & i).seat = st Then
            st = collTemp("Line" & i).seat
            GuiRow = GuiRow + 1
        End If
        collTemp("Line" & i).GuiRow = GuiRow
        GuiRow = GuiRow + 1
    Next i
End If
EH:
Set SortCheckLines = collTemp
Set coll = Nothing
Set collTemp = Nothing
End Function

Public Function NullToZero(arr As Variant) As Variant
Dim i As Integer
For i = 1 To UBound(arr)
    If IsNull(arr(i)) Then arr(i) = CStr(0)
Next i
NullToZero = arr
End Function

Public Function ReplaceDictValue(dict As Dictionary, key As String, NewVal As Variant)
Dim member As Variant
For Each member In dict.Keys
    If member = key Then
        dict.Remove key
        dict.Add key, NewVal
        Exit Function
    End If
Next member
End Function

Public Function OrderByCollID(coll As Collection) As Collection
If coll.Count = 1 Then
    Set OrderByCollID = coll
    Exit Function
End If
Dim member As aclsItem
Dim NewColl As Collection
Set NewColl = New Collection
Dim i As Integer
Do Until coll.Count = 0
For Each member In coll
    For i = 1 To coll.Count
        If Not member.CollID <= coll(i).CollID Then
            GoTo NextMember
        End If
    Next i
    NewColl.Add member, CStr(member.CollID)
    
    coll.Remove CStr(member.CollID)

NextMember:
Next member
Loop
Set OrderByCollID = NewColl
End Function

Public Function GetNextCollID(coll As Collection) As Integer
If Not coll.Count = 0 Then
GetNextCollID = coll(coll.Count).CollID + 1
Exit Function
End If
GetNextCollID = 1
End Function

Public Function OrderByParent(coll As Collection) As Collection
If coll.Count = 1 Then
    Set OrderByParent = coll
    Exit Function
End If

Dim member As aclsItem
Dim memberB As aclsItem
Dim NewColl As New Collection
NewColl.Add coll(1), CStr(coll(1).CollID)

For Each member In coll
    If member.Parent.ID = coll(1).CollID Then
        NewColl.Add member, CStr(member.CollID)
        AddChildrenToColl NewColl, member
    End If
Next member
Set OrderByParent = NewColl
End Function


Public Sub AddChildrenToColl(coll As Collection, member As aclsItem)
Dim child As bclsChild
Dim item As aclsItem
For Each child In member.Children.coll
    Set item = GetItemByID(child.ID)
    coll.Add item, CStr(item.CollID)
    AddChildrenToColl coll, item
Next child

End Sub

