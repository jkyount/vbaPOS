VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fDatePicker 
   Caption         =   "Choose a date.."
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5235
   OleObjectBlob   =   "fDatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit




Private Sub CommandButton1_Click()
CLICK_Day Me.CommandButton1
End Sub
Private Sub CommandButton2_Click()
CLICK_Day Me.CommandButton2
End Sub
Private Sub CommandButton3_Click()
CLICK_Day Me.CommandButton3
End Sub
Private Sub CommandButton4_Click()
CLICK_Day Me.CommandButton4
End Sub
Private Sub CommandButton5_Click()
CLICK_Day Me.CommandButton5
End Sub
Private Sub CommandButton6_Click()
CLICK_Day Me.CommandButton6
End Sub
Private Sub CommandButton7_Click()
CLICK_Day Me.CommandButton7
End Sub
Private Sub CommandButton8_Click()
CLICK_Day Me.CommandButton8
End Sub
Private Sub CommandButton9_Click()
CLICK_Day Me.CommandButton9
End Sub
Private Sub CommandButton10_Click()
CLICK_Day Me.CommandButton10
End Sub
Private Sub CommandButton11_Click()
CLICK_Day Me.CommandButton11
End Sub
Private Sub CommandButton12_Click()
CLICK_Day Me.CommandButton12
End Sub
Private Sub CommandButton13_Click()
CLICK_Day Me.CommandButton13
End Sub
Private Sub CommandButton14_Click()
CLICK_Day Me.CommandButton14
End Sub
Private Sub CommandButton15_Click()
CLICK_Day Me.CommandButton15
End Sub
Private Sub CommandButton16_Click()
CLICK_Day Me.CommandButton16
End Sub
Private Sub CommandButton17_Click()
CLICK_Day Me.CommandButton17
End Sub
Private Sub CommandButton18_Click()
CLICK_Day Me.CommandButton18
End Sub
Private Sub CommandButton19_Click()
CLICK_Day Me.CommandButton19
End Sub
Private Sub CommandButton20_Click()
CLICK_Day Me.CommandButton20
End Sub
Private Sub CommandButton21_Click()
CLICK_Day Me.CommandButton21
End Sub
Private Sub CommandButton22_Click()
CLICK_Day Me.CommandButton22
End Sub
Private Sub CommandButton23_Click()
CLICK_Day Me.CommandButton23
End Sub
Private Sub CommandButton24_Click()
CLICK_Day Me.CommandButton24
End Sub
Private Sub CommandButton25_Click()
CLICK_Day Me.CommandButton25
End Sub
Private Sub CommandButton26_Click()
CLICK_Day Me.CommandButton26
End Sub
Private Sub CommandButton27_Click()
CLICK_Day Me.CommandButton27
End Sub
Private Sub CommandButton28_Click()
CLICK_Day Me.CommandButton28
End Sub
Private Sub CommandButton29_Click()
CLICK_Day Me.CommandButton29
End Sub
Private Sub CommandButton30_Click()
CLICK_Day Me.CommandButton30
End Sub
Private Sub CommandButton31_Click()
CLICK_Day Me.CommandButton31
End Sub
Private Sub CommandButton32_Click()
CLICK_Day Me.CommandButton32
End Sub
Private Sub CommandButton33_Click()
CLICK_Day Me.CommandButton33
End Sub
Private Sub CommandButton34_Click()
CLICK_Day Me.CommandButton34
End Sub
Private Sub CommandButton35_Click()
CLICK_Day Me.CommandButton35
End Sub
Private Sub CommandButton36_Click()
CLICK_Day Me.CommandButton36
End Sub
Private Sub CommandButton37_Click()
CLICK_Day Me.CommandButton37
End Sub
Private Sub CommandButton38_Click()
CLICK_Day Me.CommandButton38
End Sub
Private Sub CommandButton39_Click()
CLICK_Day Me.CommandButton39
End Sub
Private Sub CommandButton40_Click()
CLICK_Day Me.CommandButton40
End Sub
Private Sub CommandButton41_Click()
CLICK_Day Me.CommandButton41
End Sub
Private Sub CommandButton42_Click()
CLICK_Day Me.CommandButton42
End Sub

Private Sub MonthBox_Change()

If Not Me.MonthBox.value = "" And Not Me.YearBox.value = "" Then

Dim vDate As Date
vDate = "1/" & Me.MonthBox.value & "/" & Me.YearBox.value
SetDays Month(vDate), Year(vDate)
Me.MonthLabel.value = Me.MonthBox.value

End If


End Sub

Public Sub CLICK_Day(caller As MSForms.CommandButton)



Dim tDate As Date
tDate = fDatePicker.MonthBox.value & " / " & CInt(caller.Caption) & " / " & fDatePicker.YearBox.value

fReports.SetDateValue (tDate)

Me.Hide




End Sub







Private Sub UserForm_Activate()
ShowDatePicker
End Sub

Private Sub UserForm_Initialize()



Dim i As Integer
For i = 1 To 12
    Me.MonthBox.AddItem (Format(i & "/1", "mmmm"))
Next i

For i = 10 To 1 Step -1
    Me.YearBox.AddItem (Format(Now, "yyyy") - (i - 1))
Next i
For i = 1 To 10
    Me.YearBox.AddItem (Format(Now, "yyyy") + i)
Next i



End Sub

Public Sub SetDays(vMonth As Integer, vYear As Integer)

Dim i As Integer
Dim StartDay As Integer
Dim EndDay As Integer
Dim FirstDay As Integer


    
For i = 1 To 42
    fDatePicker.Controls("CommandButton" & i).Caption = ""
    fDatePicker.Controls("CommandButton" & i).BackColor = RGB(255, 255, 255)
    fDatePicker.Controls("CommandButton" & i).Enabled = True
Next i


StartDay = Weekday(vMonth & "/1/" & vYear)

EndDay = Day(DateSerial(vYear, vMonth + 1, 1) - 1)
Dim j As Integer
j = 1
For i = StartDay To 42
    If Not j > EndDay Then
        fDatePicker.Controls("CommandButton" & i).Caption = j
    End If
    If j > EndDay Then
        fDatePicker.Controls("CommandButton" & i).Enabled = False
        fDatePicker.Controls("CommandButton" & i).BackColor = RGB(180, 180, 180)
        
    End If
    
  
    j = j + 1
Next i

For i = 1 To (StartDay - 1)
    fDatePicker.Controls("CommandButton" & i).Enabled = False
    fDatePicker.Controls("CommandButton" & i).BackColor = RGB(180, 180, 180)
Next i
If YearBox.value = Format(Now, "yyyy") Then
    If MonthBox.value = Format(Now, "mmmm") Then
        For i = 1 To 42
            If fDatePicker.Controls("CommandButton" & i).Caption = Format(Now, "d") Then
            fDatePicker.Controls("CommandButton" & i).BackColor = RGB(240, 206, 102)
            End If
        Next i
    End If
End If


End Sub

Public Sub ShowDatePicker()


Dim vDate As Date
vDate = Now
Me.YearBox.value = Format(Now, "yyyy")
Me.MonthBox.value = Format(Now, "mmmm")
Me.MonthLabel.value = Me.MonthBox.value

SetDays Month(vDate), Year(vDate)


End Sub



Private Sub YearBox_Change()


    
If Not Me.MonthBox.value = "" And Not Me.YearBox.value = "" Then

Dim vDate As Date
vDate = "1/" & Me.MonthBox.value & "/" & Me.YearBox.value
SetDays Month(vDate), Year(vDate)

End If

End Sub
