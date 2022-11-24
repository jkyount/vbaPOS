VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dclsArchiveData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public Function CfgDataSource(iDataObj As aclsDataObject) As aclsDataObject
iDataObj.Db = iDataObj.Archive
iDataObj.DbFile = iDataObj.ArchiveDbFile
Set CfgDataSource = iDataObj
End Function

Public Function DateClause(iReportObj As aclsReport) As String
DateClause = "Between " & Format(iReportObj.StartDate, "mmddyy") & " AND " & Format(iReportObj.EndDate, "mmddyy")
End Function
