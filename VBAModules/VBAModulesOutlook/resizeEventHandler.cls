VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "resizeEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this Class Module contains the event handlers for the resizeControls
Option Explicit
Public WithEvents resizeControl As MSForms.Label
Attribute resizeControl.VB_VarHelpID = -1
Public indexActiveControl As Integer
Dim newPosition() As Variant
Dim mX As Single
Dim mY As Single
 
Private Sub resizeControl_MouseDown(ByVal Button As Integer, _
                                   ByVal Shift As Integer, _
                                   ByVal X As Single, _
                                   ByVal Y As Single)
  'save initial cursor position
  mX = X
  mY = Y
  'test if active control is a Task or the cursor
  If greenBoard.ActiveControl.Name Like "*cursor*" Or greenBoard.ActiveControl.Name = "scheduleButton" Then
    MsgBox prompt:="Please left-click on a Task before resizing.", Buttons:=vbOKOnly + vbExclamation, Title:="Warning!"
    Exit Sub
  End If
  moveResizeControl
End Sub

'move resizeControl with its associated taskControl
Private Sub moveResizeControl(Optional ByVal I As Integer)
  Dim indexResizeControl As Integer
  On Error Resume Next    'prevent code stopping when the GreenBoard is clicked
  indexActiveControl = Right(greenBoard.ActiveControl.Name, InStr(greenBoard.ActiveControl.Name, "Task"))
  indexResizeControl = Right(resizeControl.Name, InStr(resizeControl.Name, "Resize"))
  If indexActiveControl = indexResizeControl Then
    resizeControl.Left = greenBoard.ActiveControl.Left + greenBoard.ActiveControl.Width - 5
    resizeControl.Top = greenBoard.ActiveControl.Top + greenBoard.ActiveControl.Height - 5
  End If
End Sub
Private Sub resizeControl_MouseMove(ByVal Button As Integer, _
                                   ByVal Shift As Integer, _
                                   ByVal X As Single, _
                                   ByVal Y As Single)
Dim myX As Single
Dim myY As Single
  'Shift= 1 is the shift key, so Right-Click + Sft triggers resize, but first give the label focus with left-click
  If Button = 1 And Shift = 1 Then
    myX = X - mX
    myY = Y - mY
    greenBoard.resizeTaskControlwithResizeControl myX, myY, indexActiveControl
  End If
End Sub


