Attribute VB_Name = "gui_Combine"
Option Explicit

Public Sub InitializeRanges()
Sheet12.range("OriginalCheckRange").value = ""
Sheet12.range("CombineCheckRange").value = ""

Sheet12.range("TargetCheckNumber").value = ""
Sheet12.range("TargetServerName").value = ""
Sheet12.range("TargetOrderName").value = ""
Sheet12.range("TargetPhone").value = ""
PopGuiWithCheckAttributes GetCheckAttributes(currentcheck), Sheet12
WriteCheckLines Sheet12.range("OriginalCheckRange"), collCheckData
End Sub

