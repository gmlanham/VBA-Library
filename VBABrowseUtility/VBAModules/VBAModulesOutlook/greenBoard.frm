VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} greenBoard 
   Caption         =   "Schedule Board (Green Board)"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   375
   ClientWidth     =   15255
   OleObjectBlob   =   "greenBoard.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "greenBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'this is the Schedule Board, the "Green Board"
'the purpose of this form is to View and Organize the job task Schedule/Calendar.
Option Explicit
Dim newArray() As Variant
Dim resizeControlsArray(100) As Variant
Private taskCollectionOfEventHandlers As Collection
Private taskControls As New Collection

Private resizeCollectionOfEventHandlers As New Collection
Private resizeControls As New Collection

Dim m_sngLeftPos As Single  'need these for cursor
Dim m_sngTopPos As Single   'needed for cursor

Private Sub showControlPanel_Click()
    controlPanel.Show vbModeless
End Sub

Private Sub UserForm_Initialize()
    activeTaskTimeForm.Show vbModeless
    controlPanel.Show vbModeless
End Sub

Public Sub UserForm_Terminate()
    activeTaskTimeForm.Hide
    controlPanel.Hide
    Me.Hide
End Sub

Private Sub UserForm_Activate()
  resizableModule.MakeFormResizable Me
  greenBoard.ScrollTop = 0
End Sub

Private Sub scheduleButton_Click()
    removeJobTasks
    addJobTasks
    addTasksHandler
    addResizeHandler
End Sub

'this procedure cleans the Board before populating it with the scheduled tasks
Sub removeJobTasks()
    Dim I As Integer
    On Error GoTo ExitSub   'handles error when there are no tasks to remove
    For I = 1 To UBound(newArray(), 1)
        Me.Controls.Remove "Task" & I
        Me.Controls.Remove "Resize" & I
    Next I
ExitSub:
End Sub

'this uses the Class Module taskEventHandler to minimize repeating code for each added TaskControl
Sub addTasksHandler()
    Set taskCollectionOfEventHandlers = New Collection
    Dim I As Integer
    Dim oControl As Control
    Dim oEventHandler As taskEventHandler
 
    For Each oControl In Me.Controls    'taskControls, textbox
        If TypeName(oControl) = "TextBox" Then
            Set oEventHandler = New taskEventHandler
            Set oEventHandler.taskControl = oControl
            taskCollectionOfEventHandlers.Add oEventHandler
        End If
    Next oControl
End Sub

Sub addResizeHandler()
    Set resizeCollectionOfEventHandlers = New Collection
    Dim I As Integer
    Dim oControl As Control
    Dim oEventHandler As resizeEventHandler
 
    For Each oControl In Me.Controls    'resizeControls, Buttons
        If TypeName(oControl) = "Label" Then
            Set oEventHandler = New resizeEventHandler
            Set oEventHandler.resizeControl = oControl
            resizeCollectionOfEventHandlers.Add oEventHandler
        End If
    Next oControl
End Sub

'add a textbox control to the Board for each scheduled Task, and add resize Image controls
Sub addJobTasks(Optional ByRef Cancelled As Boolean)
On Error GoTo ExitSub
    Dim taskControl As MSForms.TextBox
    Dim resizeControl As MSForms.Label
    Dim I As Integer
    Dim priorEnd As Integer
    Dim priorTop As Integer
    priorEnd = 0
    
    'the expanded array function adds the color code to the jobTasksArray
    'the jobTasksArray function gets Calendar item data to put into the jobTasksArray
    'the function gets JobTasks,
    'stores the items in newArray, so that don't have to go back to the array function
    newArray = jobTasksGetModule.jobtasksArray
    'If Cancelled Then Exit Sub
    'synch Day Labels with start date, dteStart
   If dteStart = 0 Then Exit Sub
    synchDayLabels
    'add a textbox for each JobTask
    For I = 1 To UBound(newArray(), 1)

        'add the textbox controls
        Set taskControl = Me.Controls.Add("forms.textbox.1", "Task" & I, True)
         
        'add Task# control to the taskControls array
        If TypeName(taskControl) = "TextBox" Then taskControls.Add Me.Controls("Task" & I)
              
        'set the attributes of the new Task control, using newArray as the data source
        With taskControl
                greenBoard.Caption = "Schedule Board " & dateRangeCaption
                .Enabled = True
                .Height = 24
                If newArray(I, 9) = "" Then MsgBox newArray(I, 1) & " Color is null."
                .BackColor = newArray(I, 11)
                If .BackColor = 0 Then .BackColor = newArray(I - 1, 11)
                .BorderStyle = 1
                .TextAlign = fmTextAlignLeft
                .Font.Size = 7
                .Font.Bold = True
                .WordWrap = True
                .MultiLine = True
                .ControlTipText = newArray(I, 1) '.Subject=1
                .Top = 42 + 18 * I  'priorTop
                .text = newArray(I, 10) & vbCr & newArray(I, 5) 'jobNumer=8 '.Categories=5
                .Width = newArray(I, 4) / 6.25 '.Duration
                If .Width < 4 Then .Width = 4
                .Left = newArray(I, 9) * 96 + newArray(I, 8) / 6.25 'offset from 7 am for the day the item is schedule for
                If newArray(I, 8) = 0 Then .Left = .Left + 2
                If newArray(I, 8) = 96 Then .Left = .Left - 1
                priorEnd = newArray(I, 8) + newArray(I, 4)
                priorTop = .Top
        End With
        
        'add a resizeControl for each taskControl
        Set resizeControl = Me.Controls.Add("forms.Label.1", "Resize" & I, True)
        'add Task# control to the taskControls array
        If TypeName(resizeControl) = "Label" Then resizeControls.Add Me.Controls("Resize" & I)
        
        'If TypeName(resizeControl) = "Image" Then Debug.Print Me.controls("Resize" & I).Name
        With resizeControl
            .Left = Me.Controls("Task" & I).Left + Me.Controls("Task" & I).Width - 1
            .Top = Me.Controls("Task" & I).Top + Me.Controls("Task" & I).Height - 1
            .Width = 5
            .Height = 5
            .MousePointer = 8
            .BackStyle = 1
            .BackColor = vbRed
            resizeControlsArray(I) = .Name
        End With
    Next I
ExitSub:
End Sub

'this keeps the resizeControl and its associated taskControl together
'on this code behind page is the only place that the control arrays are recognized.
'variable defined in this code behind are not Public
Sub moveResizeControlwithTaskControl(Optional ByVal I As Integer)

    resizeControls(I).Left = taskControls(I).Left + taskControls(I).Width - 5
    resizeControls(I).Top = taskControls(I).Top + taskControls(I).Height - 5
End Sub

'resize the taskControl using the resizeControl
Sub resizeTaskControlwithResizeControl(ByVal myX As Single, ByVal myY As Single, ByVal I As Integer)

    With greenBoard
      
      'make sure textbox won't get too small
      'change textbox width
      On Error Resume Next
      .ActiveControl.Width = .ActiveControl.Width + (myX)
      If .ActiveControl.Width < 4 Then .ActiveControl.Width = 4
      
      'make sure the resize handle remains fixed at
      'the left end of the textbox
      resizeControls(I).Left = taskControls(I).Left + taskControls(I).Width - 5
    End With
  
End Sub

Private Sub cursorTextBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    greenBoard.ScrollTop = 0
    greenBoard.cursorTextBox.Top = 42
End Sub

Private Sub cursorTextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    greenBoard.cursorTextBox.Left = cursorTextBox.Left + 1
End Sub

'this code nudges the cursor using the cursorHandle, about 30min
Private Sub cursorHandleLabel_Click()
    cursorTextBox.Move cursorTextBox.Left + 4
End Sub

'this code nudges the cursor using the cursorHandle, about 30min
Private Sub cursorHandleLabel2_Click()
    cursorTextBox.Move cursorTextBox.Left + 4
End Sub

Private Sub cursorHandleLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cursorTextBox.Move cursorTextBox.Left + 1
End Sub

Private Sub cursorHandleLabe2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cursorTextBox.Move cursorTextBox.Left + 1
End Sub

'this changes the captions on the Day Labels to be in synch with the filter start date entered by the user, dteStart
Sub synchDayLabels()
    Dim I As Integer
    Dim dayLabel As Label
        dayLabel0.Caption = dteStart + 0 & vbCr & WeekdayName(weekDay(dteStart + 0))
        dayLabel1.Caption = dteStart + 1 & vbCr & WeekdayName(weekDay(dteStart + 1))
        dayLabel2.Caption = dteStart + 2 & vbCr & WeekdayName(weekDay(dteStart + 2))
        dayLabel3.Caption = dteStart + 3 & vbCr & WeekdayName(weekDay(dteStart + 3))
        dayLabel4.Caption = dteStart + 4 & vbCr & WeekdayName(weekDay(dteStart + 4))
        dayLabel5.Caption = dteStart + 5 & vbCr & WeekdayName(weekDay(dteStart + 5))
        dayLabel6.Caption = dteStart + 6 & vbCr & WeekdayName(weekDay(dteStart + 6))
        dayLabel7.Caption = dteStart + 7 & vbCr & WeekdayName(weekDay(dteStart + 7))
        dayLabel8.Caption = dteStart + 8 & vbCr & WeekdayName(weekDay(dteStart + 8))
        dayLabel9.Caption = dteStart + 9 & vbCr & WeekdayName(weekDay(dteStart + 9))
        dayLabel10.Caption = dteStart + 10 & vbCr & WeekdayName(weekDay(dteStart + 10))
        dayLabel11.Caption = dteStart + 11 & vbCr & WeekdayName(weekDay(dteStart + 11))
        dayLabel12.Caption = dteStart + 12 & vbCr & WeekdayName(weekDay(dteStart + 12))
        dayLabel13.Caption = dteStart + 13 & vbCr & WeekdayName(weekDay(dteStart + 13))
        dayLabel14.Caption = dteStart + 14 & vbCr & WeekdayName(weekDay(dteStart + 14))
End Sub

'this code sets the X,Y position of the mouse
Function cursorhandleLabelPosition(ByVal Button As Integer, _
ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        m_sngLeftPos = X
        m_sngTopPos = 38
        cursorhandleLabelPosition = X
    End If
End Function

'this code drags the cursorHandle control with the mouse,
Function cursorhandleLabelMove(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim newTime As Date
    
    If Button = 1 Then

        With greenBoard.cursorHandleLabel
            sngLeft = .Left + X - m_sngLeftPos
            sngTop = Y
            newTime = DateAdd("n", (sngLeft * 14.1) + 100, dteStart)
        End With
    End If
 
End Function
