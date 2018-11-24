Attribute VB_Name = "ExportModule"
Option Explicit

Public Sub ExportVBComponent(Optional VBComp As VBIDE.VBComponent, _
               Optional folderName As String, Optional procedureName As String, _
               Optional fileName As String, _
               Optional OverwriteExisting As Boolean = True)
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' This function exports the code module of a VBComponent to a Word document.
   ' If FileName is missing, the code will be exported to
   ' a file with the same name as the VBComponent followed by the
   ' appropriate extension.
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  On Error Resume Next
  Dim Extension As String
   Dim FName As String
   Extension = GetFileExtension(VBComp:=VBComp)
   If Trim(fileName) = vbNullString Then
       FName = VBComp.Name & Extension
   Else
       FName = fileName
       If InStr(1, FName, ".", vbBinaryCompare) = 0 Then
           FName = FName & Extension
       End If
   End If
   
   If StrComp(Right(folderName, 1), "\", vbBinaryCompare) = 0 Then
       FName = folderName & FName
   Else
       FName = folderName & "\" & FName
   End If
   
   If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
       If OverwriteExisting = True Then
           Kill FName
       Else
           'ExportVBComponent = False
           Exit Sub
       End If
   End If
   
   VBComp.Export fileName:=FName
   Documents.Open FName
   Call formatReport(documentName:=FName)
   
   Dim PI As ProcInfo
   PI = ProcedureInfo(ProcName:=procedureName, ProcKind:=vbext_pk_Proc, CodeMod:=VBComp.CodeModule)
   With Selection
     .Find.Text = PI.ProcDeclaration
     .Find.Execute
   End With

Cleanup:
   If excelFlag Then closeExcel
   Exit Sub
ErrorHandler:
   MsgBox prompt:="An error occured while processing " & "ExportModule.ExportVBComponent."
   Resume Cleanup
End Sub

Public Function GetFileExtension(VBComp As VBIDE.VBComponent) As String
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' This returns the appropriate file extension based on the Type of
  ' the VBComponent.
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Select Case VBComp.Type
    Case vbext_ct_ClassModule
        GetFileExtension = ".cls"
    Case vbext_ct_Document
        GetFileExtension = ".cls"
    Case vbext_ct_MSForm
        GetFileExtension = ".frm"
    Case vbext_ct_StdModule
        GetFileExtension = ".bas"
    Case Else
        GetFileExtension = ".bas"
  End Select
      
End Function


