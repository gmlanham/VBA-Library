Attribute VB_Name = "AutomationModule"
Option Explicit

Public Sub browseExcelVBA()
  StatusBar = "Browse Excel VBA"
  OpenExcel
  proceduresData
  
  'hide the Excel VBAUtility Control Panel
  Run "AutomationModule.hideControlPanel"
End Sub

Sub OpenExcel()
Initialization:
  Const procedureName = "OpenExcel"
  Dim workbookPath As String
  StatusBar = procedureName
  excelFlag = True
  
Body:
  On Error Resume Next
  Set OExcel = GetObject(, "Excel.Application")
  If Err Then
    Set OExcel = New Excel.Application
  End If
  workbookPath = ThisDocument.path & "\VBAUtility.xlsm"

  Set OWB = OExcel.Workbooks.Open(workbookPath)
  OExcel.Visible = False
'  With Excel.Application
'    .DisplayAlerts = False
'    .MacroOptions Macro:="showControlPanel", HasShortcutKey:=True, ShortcutKey:="p"
'    .WindowState = xlMaximized
'  End With
  
Cleanup:
 
 Exit Sub
 
ErrorHandler:
  closeExcel
  Debug.Print "An error was thrown by " & procedureName & _
  vbCr & Err.Number & ": " & Err.Description
  Resume Cleanup:
End Sub

Sub proceduresData(Optional ByVal proceduresCount As Long)
'transfer procedures from Excel to this Word document
  On Error GoTo ErrorHandler:
  Const procedureName = "proceduresData"
  StatusBar = procedureName

  Dim I As Long
  Dim dataArray() As Variant
  
  'run Excel VBAUtility to count procedures
  proceduresCount = Run("AutomationModule.proceduresCount")
  FileToGet = Range("A1").Value
  'dimenstion variables with count
  ReDim dataArray(proceduresCount, 2)
  Range("A3:B" & proceduresCount + 2).Name = "proceduresTable"
  dataArray = Range("proceduresTable")
  
  For I = 1 To UBound(dataArray)
    If dataArray(I, 1) <> vbNullString Then
      ControlPanel.ProcedureListBox.AddItem dataArray(I, 1) & ", Lines: " & dataArray(I, 2)
    End If
  Next I
  
Cleanup:
  'Kill dataArray
Exit Sub
  
ErrorHandler:
  Debug.Print "An error was thrown by " & procedureName & _
  vbCr & Err.Number & ": " & Err.Description
  Resume Cleanup:
End Sub
Sub closeExcel()
  On Error GoTo ErrorHandler:
  Const procedureName = "closeExcel"
  StatusBar = procedureName
  excelFlag = False
  'close workbook
  OWB.Close SaveChanges:=False
  
Cleanup:
  'release all variables, set objects to nothing, strings to vbnullstring
  OExcel.Quit
  Set OExcel = Nothing
  Set OWB = Nothing
  'the sendkeys closes the VBE window, if this macro is called from the VBE window
  'this command close the calling app, i.e. this Word document, not Excel
  'SendKeys "%{F4}", True

Exit Sub
ErrorHandler:
  Debug.Print "An error was thrown by " & procedureName & _
  vbCr & Err.Number & ": " & Err.Description
Resume Cleanup
End Sub


