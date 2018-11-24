Attribute VB_Name = "draganddropModule"
'this is the DragAndDrop Module; the procedures are for moving the Scheduled Tasks
'there are two ways to move a Task, by drag and drop and by manually entering a new time
Option Explicit
Public taskCollectionOfEventHandlers As Collection
Public taskControls As New Collection

Public resizeCollectionOfEventHandlers As New Collection
Public resizeControls As New Collection
Public newTime As Date     'this is needed in updateTime

Dim m_sngLeftPos As Double
Dim m_sngTopPos As Double
Dim X As Double
Dim Y As Double
Dim sngLeft As Double
Dim sngTop As Double

'this function drags the ActiveControl and limits drag to within the backBoard
Function DragAndDrop(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal X As Double, _
    ByVal Y As Double, _
    ByVal mX As Double, _
    ByVal mY As Double) As Variant
    Dim Position(1, 2) As Variant
    If Button = 2 Then
        greenBoard.ActiveControl.Left = 0
        mX = 0
    End If
    'If Button = 1 Then
        With greenBoard.ActiveControl
            sngLeft = .Left + X - mX
            If sngLeft < greenBoard.backBoard.Left Then sngLeft = greenBoard.backBoard.Left
                If (sngLeft + .Width) > (greenBoard.backBoard.Left + greenBoard.backBoard.Width) Then
                    sngLeft = greenBoard.backBoard.Left + greenBoard.backBoard.Width - .Width
                End If
                
                sngTop = .Top + Y - mY
                If sngTop < greenBoard.backBoard.Top Then sngTop = greenBoard.backBoard.Top
                If (sngTop + .Height) > (greenBoard.backBoard.Top + greenBoard.backBoard.Height) Then
                    sngTop = greenBoard.backBoard.Top + greenBoard.backBoard.Height - .Height
            End If
            Position(1, 1) = sngLeft
            Position(1, 2) = sngTop
            
            'set return values
            DragAndDrop = Position
            'Debug.Print Position(1, 1), Position(1, 2)
            '.Move sngLeft, sngTop
            'newTime = DateAdd("n", (sngLeft * 14.1) + 100, jobTasksGetModule.findSunday(Date))
            'returnedTime = Module1.updatedTime(sngLeft)
            'greenBoard.ActiveControl.ControlTipText = newTime
            'Debug.Print sngLeft; sngTop; newTime
        End With
    'End If
    
    'set the DragAndDrop return array values from the Position array
    DragAndDrop = Position
    'this procedure converts control position to time
    'updateTime (sngLeft)
    
End Function
'this function moves the ActiveControl and limits drag to within the backBoard
'Function manualMove(ByVal Button As Integer, _
'    ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double, ByVal mX As Double, mY As Double) As Variant
'    Dim Position(1, 2) As Variant
''    If Button = 1 Then
'        With greenBoard.ActiveControl
'            'Debug.Print "Manual Move ", .Name
'            sngLeft = X - mX    'm_sngLeftPos  '+ .Left
'            'If sngLeft < greenBoard.backBoard.Left Then sngLeft = greenBoard.backBoard.Left
'            '    If (sngLeft + .Width) > (greenBoard.backBoard.Left + greenBoard.backBoard.Width) Then
'             '       sngLeft = greenBoard.backBoard.Left + greenBoard.backBoard.Width - .Width
'            '    End If
'
'            sngTop = .Top + Y - mY  'm_sngTopPos
'            '    If sngTop < greenBoard.backBoard.Top Then sngTop = greenBoard.backBoard.Top
'            '    If (sngTop + .Height) > (greenBoard.backBoard.Top + greenBoard.backBoard.Height) Then
'            '        sngTop = greenBoard.backBoard.Top + greenBoard.backBoard.Height - .Height
'            'End If
'            Position(1, 1) = sngLeft
'            Position(1, 2) = sngTop
'            '.Move sngLeft, sngTop
'            newTime = DateAdd("n", (sngLeft * 14.1) + 100, jobTasksGetModule.findSunday(Date))
'            'returnedTime = Module1.updatedTime(sngLeft)
'            'greenBoard.ActiveControl.ControlTipText = newTime
'            'Debug.Print sngLeft; sngTop; newTime
'        End With
''    End If
'
'    'set the DragAndDrop return array values from the Position array
'    manualMove = Position
'    'this procedure converts control position to time
'    updateTime (sngLeft)
'End Function

Private Sub testUpdateTime()
'  Dim temp As Double
'  Dim tStr(1, 2) 'array with one row and two columns
'  temp = greenBoard.ActiveControl.Left
  updateTime2 (greenBoard.ActiveControl.Left)
  Debug.Print greenBoard.ActiveControl.Left, newTime
End Sub
'for Button=2, this procedure converts control position to time, takes double, X position and returns a Date
'scale time and accomidate the workday breaks with offset between days
Function updateTime2(ByVal X As Double) As Variant
    Dim cursorPosition As Double
    Dim controlPosition As Double
    cursorPosition = X
    controlPosition = X
    'Dim newArray() As Variant
    
    'get JobTasks, store the elements in newArray, so that don't have to go back to the array function
    'newArray = jobTasksGetModule.itemArray

    Dim I As Integer
        With greenBoard
            For I = 1 To 24  'the limit equals the number of hours in a day, forthe offsets
                'the constant offset, 134.4, skips the 14 hrs of night between 5pm and 7am.
                If .cursorTextBox.Left > 96 * I Then cursorPosition = cursorPosition + 134.4
                If .ActiveControl.Left > 96 * I Then controlPosition = controlPosition + 134.4
            Next I
           
            'this is a linear fit to calculate the datetime given the cursor position and start time
            newTime = DateAdd("n", 6.25 * controlPosition + 420, dteStart)
            
            'yellow label at top of GreenBoard is the activeControlTimeLabel
            .activeControlTimeLabel.Caption = newTime & vbCr & .ActiveControl.ControlTipText
            activeTaskTimeForm.activeTaskLabel.Caption = newTime & vbCr & .ActiveControl.ControlTipText
            
        End With
    'update the Calendar by setting the newTime in the extendedArray
    'find taskControl index number and store updated time in an array
    If greenBoard.ActiveControl.Name Like "*cursor*" Then GoTo ExitFunction
    'Dim indexNumber
    'indexNumber = Right(greenBoard.ActiveControl.Name, InStr(greenBoard.ActiveControl.Name, "Task"))
    ReDim updatedTimesArray(greenBoard.Controls.Count, 2)
    updatedTimesArray(1, 1) = greenBoard.ActiveControl.Name
    updatedTimesArray(1, 2) = newTime
    'this updates the ExpandedArray
    'jobTasksGetModule.updateExpandedArray indexNumber, newTime
    'Debug.Print updatedTimesArray(indexNumber, 2)
ExitFunction:
End Function
'this procedure converts control position to time, takes double, X position and returns a Date
'scale time and accomidate the workday breaks with offset between days
Function updateTime(ByVal X As Double) As Variant
    Dim cursorPosition As Double
    Dim controlPosition As Double
    cursorPosition = sngLeft
    controlPosition = sngLeft
    
    'Dim newArray() As Variant
    
    'get JobTasks, store the elements in newArray, so that don't have to go back to the array function
    'newArray = jobTasksGetModule.itemArray

    Dim I As Integer
        With greenBoard
            For I = 1 To 24  'the limit equals the number of hours in a day, forthe offsets
                'the constant offset, 134.4, skips the 14 hrs of night between 5pm and 7am.
                If .cursorTextBox.Left > 96 * I Then cursorPosition = cursorPosition + 134.4
                If .ActiveControl.Left > 96 * I Then controlPosition = controlPosition + 134.4
            Next I
           
            'this is a linear fit to calculate the datetime given the cursor position and start time
            newTime = DateAdd("n", 6.25 * controlPosition + 420, dteStart)
            
            'yellow label at top of GreenBoard is the activeControlTimeLabel
            .activeControlTimeLabel.Caption = newTime & vbCr & .ActiveControl.ControlTipText
            activeTaskTimeForm.activeTaskLabel.Caption = newTime & vbCr & .ActiveControl.ControlTipText
            
            'Debug.Print newArray(1, 1); newArray(1, 5)
            'On Error Resume Next
            '.ActiveControl.Value = vbNullString
            '.ActiveControl.Value = .ActiveControl.ControlTipText
            
            If .ActiveControl.Name Like "*cursor*" Then
                 .cursorTextBox.ControlTipText = newTime
                 .cursorHandleLabel.ControlTipText = newTime
                 .cursorHandleLabel2.ControlTipText = newTime
                 
                 'move cursorHandle with cursor
                 .cursorHandleLabel.Move .cursorTextBox.Left - 4
                 .cursorHandleLabel2.Move .cursorTextBox.Left - 4
                 .cursorHandleLabel.ControlTipText = newTime
                 .cursorHandleLabel2.ControlTipText = newTime
                
                 Dim handlePosition As Long
                 Dim handleMove As Long
                 handlePosition = .cursorhandleLabelPosition(1, 1, X, Y)
                 handleMove = .cursorhandleLabelMove(1, 1, handlePosition, Y)
                 .cursorTextBox.Value = vbNullString
                 .cursorTextBox.text = vbNullString
            End If
        End With
    'update the Calendar by setting the newTime in the extendedArray
    'find taskControl index number and store updated time in an array
    If greenBoard.ActiveControl.Name Like "*cursor*" Then GoTo ExitFunction
    'Dim indexNumber
    'indexNumber = Right(greenBoard.ActiveControl.Name, InStr(greenBoard.ActiveControl.Name, "Task"))
    ReDim updatedTimesArray(greenBoard.Controls.Count, 2)
    updatedTimesArray(1, 1) = greenBoard.ActiveControl.Name
    updatedTimesArray(1, 2) = newTime
    updateTime = updatedTimesArray
    'this updates the ExpandedArray
    'jobTasksGetModule.updateExpandedArray indexNumber, newTime
    'Debug.Print updatedTimesArray(indexNumber, 2)
ExitFunction:
End Function

'this code converts a manually entered date into an xPosition
Function xPosition(ByVal enteredDate As Date) As Double
    Dim sngLeft As Double
    Dim positionOffset As Double
    Dim timeOffset As Double
    Dim negativeOffsetFlag As Boolean
    Dim oldTime As Date
    Dim totalOffhoursOffset As Double
    Dim I As Integer
    
    oldTime = newTime
    timeOffset = dateDiff("n", oldTime, enteredDate)
    sngLeft = (timeOffset) / 6.25
    
    'determine direction of move, forward or back in time(right or left)
    If sngLeft < 0 Then negativeOffsetFlag = True
    
    'determine the number of off hour offsets
    Do Until Abs(sngLeft) < 134.4
        I = I + 1
        If negativeOffsetFlag Then
            sngLeft = sngLeft + 230.4
        Else
            sngLeft = sngLeft - 230.4
        End If
    Loop
    totalOffhoursOffset = I * 96
    If negativeOffsetFlag Then totalOffhoursOffset = -totalOffhoursOffset
    
    'the return value of this function, TaskControl.Left
    xPosition = greenBoard.ActiveControl.Left + sngLeft + totalOffhoursOffset
    
    greenBoard.activeControlTimeLabel.Caption = enteredDate & vbCr & _
    greenBoard.ActiveControl.ControlTipText
    activeTaskTimeForm.activeTaskLabel.Caption = enteredDate & vbCr & _
    greenBoard.ActiveControl.ControlTipText
End Function



