Attribute VB_Name = "openModule"
Option Explicit
Dim oExcel As Excel.Application
Dim oWB As Workbook

Sub OpenWorkbook()
'this code show how to open a workbook, with visible=True
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If Err Then
      Set oExcel = New Excel.Application
  End If
  On Error GoTo ErrorHandler
  
  Set oWB = Workbooks.Open(ThisWorkbook.Path & "\workingMacros.xlsm")
  oWB.Worksheets("Sheet1").Visible = True
  oWB.Windows(1).Visible = True
Cleanup:

Exit Sub
ErrorHandler:
  MsgBox "Error"
Resume Cleanup
End Sub
