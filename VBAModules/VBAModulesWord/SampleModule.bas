Attribute VB_Name = "SampleModule"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" _
Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As _
Long) As Long

Public Function ComputerName() As String
  Dim sBuffer As String
  
  Dim lAns As Long
 
  sBuffer = Space$(255)
  lAns = GetComputerName(sBuffer, 255)
  If lAns <> 0 Then
        'read from beginning of string to null-terminator
        ComputerName = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
   Else
        Err.Raise Err.LastDllError, , _
          "A system call returned an error code of " _
           & Err.LastDllError
   End If

End Function

'use FileDialogOpen popup to open a Quote to save
Sub browseQuote()
  TestForm.hide
  Const procedureName = "browseQuote"
  Const msoFileDialogOpen = 1
  Dim newPath As String
  Dim QuoteToParse As String
  'AbbeyFlag = False
  'EagleFlag = False
  'PegasusFlag = False
  Dim oWord As Word.Application
  Set oWord = CreateObject("Word.Application")
  
  'call function to determine ComputerName
  If ComputerName = "MIKE2" Then
    newPath = ThisDocument.Path
    Else
    newPath = "M:\Estimating and Invoicing\Estimating and Invoicing 2012\Estimating 2012"
  End If
  
  oWord.ChangeFileOpenDirectory newPath
  With oWord.FileDialog(msoFileDialogOpen)
    .Title = "Select Quote Template"
    .AllowMultiSelect = False
  End With
  Selection.WholeStory
  
  If oWord.FileDialog(msoFileDialogOpen).Show = -1 Then
    QuoteToParse = oWord.FileDialog(msoFileDialogOpen).SelectedItems.Item(1)
    Documents.Open (QuoteToParse)
    Selection.WholeStory
    Selection.Copy
    ActiveDocument.Close SaveChanges:=False
    
  'this is a nice bit here. The activedocument was the Quote selected in the Dialog,
  'the activedocument is copied, then closed, then the quoteParsingUtility is activated
  'and the clip board containing the Quote selected is pasted into the Utility doc. Sweet,
    ThisDocument.Activate
    Selection.Paste
  End If
  ActiveDocument.Bookmarks("\StartOfDoc").Select
Cleanup:
  oWord.Quit
  Set oWord = Nothing
  Exit Sub
  
ErrorHandler:
    MsgBox prompt:=Err.Description, Buttons:=vbOKOnly + vbCritical, Title:="Browse Quote"
    Resume Cleanup
 End Sub



