VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ControlPanel 
   Caption         =   "Control Panel"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8820
   OleObjectBlob   =   "ControlPanel.frx":0000
End
Attribute VB_Name = "ControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub UserForm_Deactivate()
  'Debug.Print "Userform_Deactivate"
  Set OWord = Nothing
  Set ODoc = Nothing
End Sub

Private Sub UserForm_Initialize()
  With Me
    .StartUpPosition = 0
    .Top = Application.Top + Application.Height - (.Height + 30)
    .Left = Application.Left + Application.Width - (.Width + 30)
  End With
End Sub

Private Sub btnBrowse_Click()
  btnClear_Click
  Dim cancelFlag As Boolean
  cancelFlag = True
  FileToGet = browseFiles
  If FileToGet = vbNullString Then Exit Sub 'no selection in browse window
  
  'cancelFlag=True, then procedures not typetext'd to page
  cancelFlag = VBIDEModule.getProceduresWord(FileToGet)
  'do not format cancelled
  If cancelFlag = False Then formatModule.formatSection
End Sub

Private Sub btnBrowseExcel_Click()
  btnClear_Click
  AutomationModule.browseExcelVBA
End Sub

Private Sub btnBrowseBAS_Click()
  Dim ModuleToGet As String
  ModuleToGet = BrowseModule.browseFiles(fileType:=".bas")
  If ModuleToGet = vbNullString Then Exit Sub 'no selection in browse window
  Documents.Open ModuleToGet
  Word.Application.Options.CheckSpellingAsYouType = False
  Call formatReport(ModuleToGet)
End Sub

Public Sub btnClear_Click()
  With Selection
    .WholeStory
    .Delete
  End With
  ControlPanel.ProcedureListBox.Clear
  For Each ODoc In Documents
    If ODoc.Name <> "VBAUtility.docm" Then Documents(ODoc.Name).Close SaveChanges:=False
  Next ODoc
End Sub

Private Sub btnClose_Click()
  UserForm_Deactivate
  Unload Me
End Sub

Private Sub ProcedureListBox_Click()
  If excelFlag Then
    Call VBIDEModule.viewSelectedItemExcel
  Else
    Call VBIDEModule.viewSelectedItemWord
  End If
End Sub


