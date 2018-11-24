Attribute VB_Name = "menuModule"
Option Explicit
Dim oPopUp As CommandBarPopup
Dim oCtr As CommandBarControl

Public Sub AddContextMenuItems()
'Source- http://gregmaxey.mvps.org/word_tip_pages/customize_shortcut_menu.html
  'Prevent double customization
  
  RemoveContextMenuItems
  Dim oBtn As CommandBarButton
  Set oPopUp = CommandBars.FindControl(Tag:="custPopup")
  If Not oPopUp Is Nothing Then GoTo Add_Individual
  'Add PopUp menu control to the top of the "Text" short-cut menu
  Set oPopUp = CommandBars("Text").Controls.Add(msoControlPopup, , , 1)
  With oPopUp
   .Caption = "Quote Control Panel"
   .Tag = "custPopup"
   .BeginGroup = True
  End With
  'Add controls to the PopUp menu
  Set oBtn = oPopUp.Controls.Add(msoControlButton)
  With oBtn
    .Caption = "Show Control &Panel"
    '.FaceId = 71
    .Style = msoButtonIconAndCaption
    'Identify the module and procedure to run
   .OnAction = "opencloseModule.showControlPanel"
  End With
  Set oBtn = oPopUp.Controls.Add(msoControlButton)
  With oBtn
    .Caption = "Select Quote Template"
    '.FaceId = 71
    .Style = msoButtonIconAndCaption
    'Identify the module and procedure to run
   .OnAction = "BrowseModule.selectQuoteTemplate"
  End With
  Set oBtn = oPopUp.Controls.Add(msoControlButton)
  With oBtn
    .Caption = "Quick Quote"
    '.FaceId = 71
    .Style = msoButtonIconAndCaption
    'Identify the module and procedure to run
   .OnAction = "AutomationModule.selectTakeoffBuildQuote"
  End With

'  Set oBtn = Nothing
'  'Add a Builtin command using ID 1589 (Co&mments)
'  Set oBtn = oPopUp.Controls.Add(msoControlButton, 1589)
'  Set oBtn = Nothing
'  'Add the third button
'  Set oBtn = oPopUp.Controls.Add(msoControlButton)
'  With oBtn
'    .Caption = "AutoText Complete"
'    .FaceId = 940
'    .Style = msoButtonIconAndCaption
'    .OnAction = "MySCMacros.MyInsertAutoText"
'  End With
  Set oBtn = Nothing
Add_Individual:
  'Or add individual commands directly to menu
'  Set oBtn = CommandBars.FindControl(Tag:="custCmdBtn")
'  If Not oBtn Is Nothing Then Exit Sub
'  'Add control using built-in ID 758 (Boo&kmarks...)
'  Set oBtn = Application.CommandBars("Text").Controls.Add(msoControlButton, 758, , 2)
'  oBtn.Tag = "custCmdBtn"
'  If MsgBox("This action caused a change to your Normal template." _
'     & vbCr + vbCr & "Recommend you save those changes now.", vbInformation + vbOKCancel, _
'     "Save Changes") = vbOK Then
    NormalTemplate.Save
    ThisDocument.Save
  'End If
  Set oPopUp = Nothing
  Set oBtn = Nothing
lbl_Exit:
  Exit Sub
End Sub
Sub RemoveContextMenuItems()
  'Make command bar changes in Normal.dotm
  CustomizationContext = NormalTemplate
  On Error Resume Next
  'Set oPopUp = CommandBars.FindControl(Tag:="custPopup")

  Set oPopUp = CommandBars("Text").Controls("Quote Control Panel")
  'Delete individual commands on the PopUp menu
  For Each oCtr In oPopUp.Controls
    oCtr.Delete
  Next
  Set oPopUp = CommandBars("Text").Controls("Quote Control Panel")
  'Delete individual commands on the PopUp menu
  For Each oCtr In oPopUp.Controls
    oCtr.Delete
  Next
  Set oPopUp = CommandBars("Text").Controls("Quote Control Panel")
  'Delete individual commands on the PopUp menu
  For Each oCtr In oPopUp.Controls
    oCtr.Delete
  Next
  'Delete the PopUp itself
  oPopUp.Delete
  'Delete individual custom commands on the Text menu
  For Each oCtr In Application.CommandBars("Text").Controls
    If oCtr.Caption = "Boo&kmark..." Then
      oCtr.Delete
      Exit For
    End If
  Next oCtr
'  If MsgBox("This action caused a change to your Normal template." _
'      & vbCr + vbCr & "Recommend you save those changes now.", vbInformation + vbOKCancel, _
'      "Save Changes") = vbOK Then
    NormalTemplate.Save
    ThisDocument.Save
'  End If
  Set oPopUp = Nothing
  Set oCtr = Nothing
  Exit Sub
Err_Handler:
End Sub
