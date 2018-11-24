Attribute VB_Name = "FileSystemModule"
Option Explicit

Sub saveProcedureList()
  On Error Resume Next
  
  Dim objFSO As Variant
  Dim fldr As Variant
  Dim txtFile As Variant
  Dim i As Integer
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set fldr = objFSO.CreateFolder("C:\MyTest")
  Set txtFile = objFSO.CreateTextFile("C:\MyTest\testfile.txt", True)
  
  For i = 0 To ControlPanel.ProcedureListBox.ListCount - 1
      txtFile.Write (ControlPanel.ProcedureListBox.List(i)) & vbCrLf
  Next
  
  txtFile.Close
End Sub
