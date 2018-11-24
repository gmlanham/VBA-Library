VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ControlPanel 
   Caption         =   "Control Panel"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6720
   OleObjectBlob   =   "ControlPanel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
  Unload Me
End Sub

Private Sub btnList_Click()
  VBIDEModule.getProcedures
End Sub

Private Sub btnSave_Click()
  FileSystemModule.saveProcedureList
End Sub

