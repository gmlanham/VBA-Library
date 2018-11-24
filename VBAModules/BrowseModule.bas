Attribute VB_Name = "BrowseModule"
Option Explicit
  Public FileToGet As String
  Public newPath As String

'use FileDialogOpen popup to open a file to get procedures
Function browseFiles()
On Error GoTo ErrorHandler
  Const ProcedureName = "browseFiles"
  Const msoFileDialogOpen = 1
  Dim oWord As Word.Application
  Dim excelFlag As Boolean
  'Set oWord = CreateObject("Word.Application")
  
  On Error Resume Next
  Set oWord = GetObject(, "Word.Application")
  If Err Then
      Set oWord = New Word.Application
      'WordWasNotRunning = True
  End If

  
  
  newPath = "C:\Users\" & Application.UserName & "\Documents"
  
  oWord.ChangeFileOpenDirectory newPath
  With oWord.FileDialog(msoFileDialogOpen)
    .InitialFileName = newPath
    .Filters.Clear
    .Filters.Add "Word files", "*.do*"
    .Title = "Select File"
    .AllowMultiSelect = False
  End With
  
  If oWord.FileDialog(msoFileDialogOpen).Show = -1 Then
    FileToGet = oWord.FileDialog(msoFileDialogOpen).SelectedItems.Item(1)
    Debug.Print FileToGet ', Application.FileSearch.FileType
    If InStr(FileToGet, "docm") Then
      excelFlag = False
      Documents.Open FileToGet, Visible:=False
    Else
      excelFlag = True
      Workbooks.Open (FileToGet)
      Excel.Application.Visible = False
    End If
    browseFiles = FileToGet
  End If
  ActiveDocument.Bookmarks("\StartOfDoc").Select
Cleanup:
  'Documents(FileToGet).Windows(1).WindowState = wdWindowStateNormal
  
  Exit Function
  
ErrorHandler:
    LogErrorToFile
    MsgBox prompt:=Err.Description, Buttons:=vbOKOnly + vbCritical, Title:="BrowseFiles"
    Resume Cleanup
 End Function


