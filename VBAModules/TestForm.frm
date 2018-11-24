VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestForm 
   Caption         =   "Test Form"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2475
   OleObjectBlob   =   "TestForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub testButton_Click()
  SampleModule.browseQuote
End Sub
