VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "taskEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this Class Module contains the event handlers for the taskControls
Option Explicit
Public WithEvents taskControl As MSForms.TextBox
Attribute taskControl.VB_VarHelpID = -1
Dim newPosition() As Variant
Dim mX As Double
Dim mY As Double
 
Public Property Set TextBox(ByVal taskTextBox As TextBox)
  Set taskControl = taskTextBox
End Property

'this performs the DragandDrop, a core functionality of this app.
Private Sub taskControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
ByVal X As Single, ByVal Y As Single)
  If Button = 1 Then
    newPosition = draganddropModule.DragAndDrop(Button, Shift, X, Y, mX, mY)
    taskControl.Move newPosition(1, 1), newPosition(1, 2)
    
    'updateTime converts a position to a date/time and updates ControlTips and YellowLabel
    draganddropModule.updateTime (newPosition(1, 1))
    
    'keep task and resize controls together
    Dim indexActiveControl As Integer
    If Not greenBoard.ActiveControl.Name Like "cursor*" Then
    'Debug.Print greenBoard.ActiveControl.Name
      indexActiveControl = Right(greenBoard.ActiveControl.Name, InStr(greenBoard.ActiveControl.Name, "Task"))
      With greenBoard
        .Controls("Resize" & indexActiveControl).Left = .ActiveControl.Left + .ActiveControl.Width - 5
        .Controls("Resize" & indexActiveControl).Top = .ActiveControl.Top + .ActiveControl.Height - 5
      End With
    End If
  End If
End Sub

'this code sets the mouse offsets/positions mX, mY
Private Sub taskControl_MouseDown(ByVal Button As Integer, _
ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  mX = X
  mY = Y
  
  'keep task and resize controls together
  Dim indexActiveControl As Integer
  If Not greenBoard.ActiveControl.Name Like "cursor*" Then
  'Debug.Print greenBoard.ActiveControl.Name
    indexActiveControl = Right(greenBoard.ActiveControl.Name, InStr(greenBoard.ActiveControl.Name, "Task"))
    With greenBoard
      .Controls("Resize" & indexActiveControl).Left = .ActiveControl.Left + .ActiveControl.Width - 5
      .Controls("Resize" & indexActiveControl).Top = .ActiveControl.Top + .ActiveControl.Height - 5
    End With
  End If
End Sub

'this code is for manual rescheduling of the Tasks, Schedule
Private Sub taskControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
ByVal X As Single, ByVal Y As Single)
  Dim timeDate As Date
  Dim tempStr As String
  Dim manualEntry As Double
  'Dim newPosition As Variant
  Dim newX As Double
  
  'when this function is called the public variable newTime is calculated
  updateTime2 (greenBoard.ActiveControl.Left)
  
  If Button = 1 Then
    mX = X
    mY = Y
  Else
    On Error Resume Next
    If newTime <> 0 Then
      With taskAppointmentForm
        .startTextBox.Value = newTime
        .durationTextBox.Value = Format(taskControl.Width * 6.25, "#")
      End With
    End If
    With UpdateTaskForm
      .startTimeTextBox.Value = newTime 'updateTime(mX)  'test time calculation here
      .durationTextBox.Value = Format(taskControl.Width * 6.25, "#")
      .Caption = "Update " & greenBoard.ActiveControl.ControlTipText

      .Show
      If .cancelButton.Cancel Then Exit Sub
      tempStr = .startTimeTextBox.Value
    End With
    If tempStr = "12:00:00 AM" Or tempStr = vbNullString Then GoTo ExitSub
           
    'convert string to date
    timeDate = CDate(tempStr)
    
    'xPosition converts a date to a position
    newX = draganddropModule.xPosition(timeDate)
    
    'manualMove takes the position calaculated by xPosition to determine taskControl.Left
    newPosition = draganddropModule.DragAndDrop(Button, Shift, newX, Y, mX, mY)
    
    'move the taskControl
    taskControl.Move newPosition(1, 1), newPosition(1, 2)
    
    'updateTime converts a position to a date/time and updates ControlTips and YellowLabel
    draganddropModule.updateTime (newX)
    
    'find taskControl index number and store updated time in an array
    Dim indexNumber
    indexNumber = Right(taskControl.Name, InStr(taskControl.Name, "Task"))
    ReDim updatedTimesArray(greenBoard.Controls.Count, 2)
    updatedTimesArray(indexNumber, 1) = taskControl.Name
    updatedTimesArray(indexNumber, 2) = newTime
    Dim updatedDuration As Long
    updatedDuration = UpdateTaskForm.durationTextBox.Value
    
    'call the schedule update sub, pass the index of the Control and the new time
    jobTasksSetModule.updateTaskSchedule I:=indexNumber, newTime:=timeDate, taskWidth:=updatedDuration / 6.25
  
  End If
  Exit Sub
  
ExitSub:
  MsgBox prompt:="Please left-click on a task to select it, before right-clicking.", Buttons:=vbOKOnly + vbExclamation, Title:="Warning"
End Sub


