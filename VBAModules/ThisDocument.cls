VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
 
 Private Sub Document_Open()
  On Error Resume Next
  Application.DisplayAlerts = False
  menuModule.AddContextMenuItems
  'TakeoffDataGetModule.OpenExcel
  'formatModule.setMargins
  ActiveDocument.Bookmarks("\StartOfDoc").Select
  Selection.MoveLeft Unit:=wdCharacter, Count:=1
  opencloseModule.showControlPanel
End Sub

Private Sub Document_Close()
  menuModule.RemoveContextMenuItems
  'TakeoffDataGetModule.closeExcel
  ThisDocument.Bookmarks("\StartOfDoc").Select

  With Selection
    .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    .Copy  'copy one character to ensure clipboard is nearly empty
  End With
End Sub


