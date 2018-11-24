VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} categoriesForm 
   Caption         =   "Categories Setup"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6885
   OleObjectBlob   =   "categoriesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "categoriesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'initialize fires before activate
Private Sub UserForm_Initialize()
    categoriesModule.categoriesList2
    categoriesModule.standardCategories
End Sub
