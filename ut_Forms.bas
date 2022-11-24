Attribute VB_Name = "ut_Forms"
Option Explicit


Public Sub PopDropDown(ctrl As MSForms.ComboBox, arr As Variant)
ctrl.Clear
If Not TypeName(arr) = "Boolean" Then
Dim i As Integer
For i = 1 To UBound(arr)
ctrl.AddItem arr(i)(0, 0)
Next i
If Not ctrl.ListCount = 0 Then
    ctrl.value = ctrl.List(0)
End If
End If
End Sub

Public Function ValidComboBoxValue(ComboBox As MSForms.ComboBox) As Boolean
Dim arr As Variant
arr = ComboBox.List
Dim i As Integer
For i = 0 To UBound(arr)
    If ComboBox.value = arr(i, 0) Then
        ValidComboBoxValue = True
        Exit For
        Exit Function
    End If
Next i
End Function

Public Function GetCtrlDict(Form As MSForms.UserForm, ParamArray Ctrls() As Variant) As Dictionary
Dim arr As Variant
arr = Ctrls
Dim CtrlDict As New Dictionary
Dim i As Integer
For i = 0 To UBound(Ctrls)
    On Error Resume Next
    CtrlDict.Add Ctrls(i), Form.Controls(Ctrls(i))
Next i
On Error GoTo 0
Set GetCtrlDict = CtrlDict
End Function

Public Sub DisplayCtrls(CtrlDict As Dictionary, value As Boolean)
Dim member As Variant
For Each member In CtrlDict.Keys
    CtrlDict(member).Visible = value
Next member
Set member = Nothing
End Sub

Public Sub EnableCtrls(CtrlDict As Dictionary, value As Boolean)
Dim member As Variant
For Each member In CtrlDict.Keys
    CtrlDict(member).Enabled = value
Next member
Set member = Nothing
End Sub

Public Function GetFormValueDict(Frame As MSForms.Control) As Dictionary
Dim dict As New Dictionary
Dim ctrl As Control
For Each ctrl In Frame.Controls
    If Not TypeName(ctrl) = "Label" Then
        If Not TypeName(ctrl) = "Frame" Then
            dict.Add ctrl.name, ctrl.value
        End If
    End If
Next ctrl
Set GetFormValueDict = dict
Set dict = Nothing
Set ctrl = Nothing
End Function

Public Sub PopFormWithValues(Values As Dictionary, Frame As MSForms.Control)
Dim ctrl As Control
For Each ctrl In Frame.Controls
    If Values.Exists(ctrl.name) Then
        ctrl.value = Values(ctrl.name)
    End If
Next ctrl
Set ctrl = Nothing
End Sub

