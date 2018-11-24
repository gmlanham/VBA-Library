Attribute VB_Name = "recoveryModule"
Option Explicit

'\for this macro to run you need to establish a reference to the
 '\Microsoft Excel 8.0 Object Library
 '\also, if you get a File Open error message, hit Debug, then Continue
 
Sub Recover_Excel_VBA_modules()
 
Dim XL As Excel.Application
 Dim XLVBE As Object
 Dim i As Integer, j As Integer
 
Set XL = New Excel.Application
 
XL.Workbooks.Open fileName:=ThisDocument.Path & "\TakeoffUtility4.xlsm"
 
Set XLVBE = XL.VBE
 
j = XLVBE.VBProjects(1).VBComponents.Count
 
For i = 1 To j
 Debug.Print XLVBE.VBProjects(1).VBComponents(i).Name
 XLVBE.VBProjects(1).VBComponents(i).Export _
 fileName:=ThisDocument.Path & "\vbe_" & (100 + i) & ".txt"
 Next
 
XL.Quit
 Set XL = Nothing
 
End Sub

Sub ExportAllVBAWord(Optional fileName As Variant)
'Exports all Modules, etc. to folder named the same as the Workbook.
'code source- http://us.generation-nt.com/answer/recover-vba-code-corrupted-xls-file-help-23926372.html
Dim VBComp As VBIDE.VBComponent
Dim PartPath As String
Dim NextPartPath As String
Dim TotalPath As String
Dim Sfx As String
Dim d As Integer
 
If IsMissing(fileName) Then
  fileName = ActiveDocument.Name
Else
 fileName = CStr(fileName)
End If

PartPath = Left(ThisDocument.Path, InStrRev(ThisDocument.Path, "\"))
Documents.Open PartPath & "\VBAWord\" & fileName, Visible:=False
NextPartPath = Left$(fileName, Len(fileName) - 5)
 
On Error Resume Next
 TotalPath = PartPath & "\VBAModules\VBAModulesWord\" & NextPartPath
 ChDir TotalPath
 If Err.Number = 76 Then MkDir TotalPath
 On Error GoTo ErrExport
 
For Each VBComp In Documents(fileName).VBProject.VBComponents
 Select Case VBComp.Type
 Case vbext_ct_ClassModule, vbext_ct_Document
 Sfx = ".cls"
 Case vbext_ct_MSForm
 Sfx = ".frm"
 Case vbext_ct_StdModule
 Sfx = ".bas"
 Case Else
 Sfx = ""
 End Select
 If Sfx <> "" Then
  VBComp.Export fileName:=PartPath & "\VBAModules\VBAModulesWord\" & NextPartPath & "\" & VBComp.Name & Sfx
 End If
Next VBComp
 
Exit Sub
 
ErrExport:
 
MsgBox "The reason for this message:" & vbCrLf & "The Error Number" _
 & " is: " & Err.Number _
 & vbCrLf & "The Error Description is: " & Err.Description
 
End Sub
 
Sub ExportAllVBAExcel(Optional fileName As Variant)
'Exports all Modules, etc. to folder named the same as the Workbook.
'code source- http://us.generation-nt.com/answer/recover-vba-code-corrupted-xls-file-help-23926372.html
Dim VBComp As VBIDE.VBComponent
Dim PartPath As String
Dim NextPartPath As String
Dim TotalPath As String
Dim Sfx As String
Dim d As Integer
Dim oExcel As Excel.Application
Dim oWB As Workbook
If IsMissing(fileName) Then
  fileName = ActiveWorkbook.Name
Else
 fileName = CStr(fileName)
End If
PartPath = Left(ThisDocument.Path, InStrRev(ThisDocument.Path, "\"))
OpenWorkbook:
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If Err Then
      Set oExcel = New Excel.Application
  End If
  On Error GoTo ErrorHandler
  'ChDir PartPath
  Set oWB = oExcel.Workbooks.Open(PartPath & "\VBAExcel\" & fileName)
  oWB.Windows(1).Visible = True
  
Main:
  NextPartPath = Left$(fileName, Len(fileName) - 5)
   
  On Error Resume Next
   TotalPath = PartPath & "\VBAModules\VBAModulesExcel\" & NextPartPath
   ChDir TotalPath
   If Err.Number = 76 Then MkDir TotalPath
   On Error GoTo ErrorHandler
   
  For Each VBComp In oWB.VBProject.VBComponents
   Select Case VBComp.Type
   Case vbext_ct_ClassModule, vbext_ct_Document
   Sfx = ".cls"
   Case vbext_ct_MSForm
   Sfx = ".frm"
   Case vbext_ct_StdModule
   Sfx = ".bas"
   Case Else
   Sfx = ""
   End Select
   If Sfx <> "" Then
    VBComp.Export fileName:=PartPath & "\VBAModules\VBAModulesExcel\" & NextPartPath & "\" & VBComp.Name & Sfx
   End If
  Next VBComp
 
Cleanup:
  Set oExcel = Nothing
  Set oWB = Nothing
Exit Sub

ErrorHandler:
  MsgBox "The reason for this message:" & vbCrLf & "The Error Number" _
   & " is: " & Err.Number _
   & vbCrLf & "The Error Description is: " & Err.Description
  Resume Cleanup
End Sub
 
 
