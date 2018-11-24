Attribute VB_Name = "BrowseModule"
Option Explicit
'  Public FileToGet As String
'  Public newPath As String
'  Public oWord As Word.Application
'  Public oDoc As Document
Private Sub test()
    ThisDocument.Windows(1).Visible = True
    ControlPanel.Show vbModeless
End Sub

'use FileDialogOpen popup to open a file to get procedures
Function browseFiles(Optional ByVal fileType As String) As String
  Const procedureName = "browseFiles"
  Const msoFileDialogOpen = 1
'  Dim oWord As Word.Application
  'Dim excelFlag As Boolean
  'Set oWord = CreateObject("Word.Application")
  excelFlag = False
  On Error Resume Next
  Set OWord = GetObject(, "Word.Application")
  If Err Then
      Set OWord = New Word.Application
      'WordWasNotRunning = True
  End If
  On Error GoTo ErrorHandler
  newPath = Documents("VBAUtility.docm").path & "\VBAWord"
  If InStr(fileType, ".bas") Then newPath = Documents("VBAUtility.docm").path & "\VBAModules"
  'oWord.ChangeFileOpenDirectory newPath
  With OWord.FileDialog(msoFileDialogOpen)
    .InitialFileName = newPath
    .Filters.Clear
    .Filters.Add "Word files", "*.do*, filetype"
    .Title = "Select File"
    .AllowMultiSelect = False
  End With
  
  If OWord.FileDialog(msoFileDialogOpen).Show = -1 Then
    FileToGet = OWord.FileDialog(msoFileDialogOpen).SelectedItems.Item(1)
    'Debug.Print FileToGet ', Application.FileSearch.FileType
    If InStr(FileToGet, "docm") Then
      excelFlag = False
      Documents.Open FileToGet, Visible:=False
    ElseIf InStr(FileToGet, "xlsm") Then
      excelFlag = True
      Set OWB = Workbooks.Open(FileToGet)
      OWB.Windows(1).Visible = False
    Else
      'Documents.Open FileToGet, Visible:=True
    End If
    browseFiles = FileToGet
  End If
  Documents("VBAUtility.docm").Bookmarks("\StartOfDoc").Select
Cleanup:
  'Documents(FileToGet).Windows(1).WindowState = wdWindowStateNormal
  Exit Function
  
ErrorHandler:
    LogErrorToFile
    MsgBox prompt:=Err.Description, Buttons:=vbOKOnly + vbCritical, Title:="BrowseFiles"
    Resume Cleanup
 End Function


