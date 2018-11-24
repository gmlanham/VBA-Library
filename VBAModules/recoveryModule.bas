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
 
XL.Workbooks.Open FileName:=ThisDocument.Path & "\TakeoffUtility4.xlsm"
 
Set XLVBE = XL.VBE
 
j = XLVBE.VBProjects(1).VBComponents.Count
 
For i = 1 To j
 Debug.Print XLVBE.VBProjects(1).VBComponents(i).Name
 XLVBE.VBProjects(1).VBComponents(i).Export _
 FileName:=ThisDocument.Path & "\vbe_" & (100 + i) & ".txt"
 Next
 
XL.Quit
 Set XL = Nothing
 
End Sub

Sub ExportAllVBA(Optional varName As Variant)
'Exports all Modules, etc. to folder named the same as the Workbook.
'code source- http://us.generation-nt.com/answer/recover-vba-code-corrupted-xls-file-help-23926372.html
Dim VBComp As VBIDE.VBComponent
Dim PartPath As String
Dim NextPartPath As String
Dim TotalPath As String
Dim Sfx As String
Dim d As Integer
 
If IsMissing(varName) Then
  varName = ActiveWorkbook.Name
Else
 varName = CStr(varName)
End If

PartPath = "C:\Money Files\Computer Helpers\Modules\"
NextPartPath = Left$(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4)
 
On Error Resume Next
 TotalPath = PartPath & NextPartPath
 ChDir TotalPath
 If Err.Number = 76 Then MkDir TotalPath
 On Error GoTo ErrExport
 
For Each VBComp In ActiveWorkbook.VBProject.VBComponents
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
  VBComp.Export FileName:=PartPath & NextPartPath & "\" & VBComp.Name & Sfx
 End If
Next VBComp
 
Exit Sub
 
ErrExport:
 
MsgBox "The reason for this message:" & vbCrLf & "The Error Number" _
 & " is: " & Err.Number _
 & vbCrLf & "The Error Description is: " & Err.Description
 
End Sub
 
 
