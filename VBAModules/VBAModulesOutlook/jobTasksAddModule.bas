Attribute VB_Name = "jobTasksAddModule"
'jobTasksAddModule, code to add or edit tasks
Option Explicit

'add new appointment/jobTask
'http://www.excelforum.com/excel-programming/698020-create-outlook-appointment-from-vba.html

Sub fillinForm(Optional I As Variant)   'Optional parameter must be Variant
On Error GoTo ExitSub
With taskAppointmentForm
    If IsMissing(I) = False Then
        .indexLabel.Caption = I
        '.indexLabel.BackColor = combinedArray(I, 11)
        '.indexLabel.BackStyle = 1
        .taskTextBox.text = combinedArray(I, 1) ' "#111-2011 Demo Appointment"
        .startTextBox.Value = combinedArray(I, 2) ' DateSerial(2011, 8, 25) + TimeSerial(11, 30, 0)
        '.endTextBox.Value = combinedArray(I, 3) ' .End + "00:30"
        .durationTextBox.text = combinedArray(I, 4) ' .Duration
        .categoriesComboBox.text = combinedArray(I, 5) ' "Meet Client"
        .contactsComboBox = combinedArray(I, 6)
        .notesTextBox = combinedArray(I, 7)
        .enterButton.Caption = combinedArray(I, 10)
        '.enterButton.BackColor = combinedArray(I, 11)
        .colorChoserLabel.BackColor = combinedArray(I, 11)
      
    '   .locationTextbox.Text = combinedArray(I, 1) ' "Red Deer"
    Else
        'initialize textboxes
        .indexLabel.Caption = vbNullString
        .indexLabel.BackStyle = 1
        .taskTextBox.text = "12-0-xxxx Test Job Meet Client" ' "12-0-2222 Demo Appointment"
        .startTextBox.Value = vbNullString ' DateSerial(2011, 8, 25) + TimeSerial(11, 30, 0)
    End If
End With

ExitSub:
taskAppointmentForm.Show vbModeless
End Sub

Public Sub addTaskAppointment()
Dim OLApp As Outlook.Application
Dim OLAI As Outlook.AppointmentItem
Set OLApp = New Outlook.Application
Set OLAI = OLApp.CreateItem(olAppointmentItem)
With OLAI
    .Subject = taskAppointmentForm.taskTextBox.text ' "Demo Appointment"
    If .Subject = "" Then GoTo ErrorHandler
    .Start = taskAppointmentForm.startTextBox.Value ' DateSerial(2011, 8, 25) + TimeSerial(11, 30, 0)
    '.End = taskAppointmentForm.endTextBox.Value ' .Start + "00:30"
    .Duration = taskAppointmentForm.durationTextBox.text ' .Start + "00:30"
    .Categories = taskAppointmentForm.categoriesComboBox.text ' "Meet Client"
    On Error Resume Next
    .Body = taskAppointmentForm.notesTextBox.text ' "Here is Meeting Request"
    .Location = taskAppointmentForm.locationTextBox.text ' "Red Deer"
    .RequiredAttendees = taskAppointmentForm.contactsComboBox.text ' "m_lanham@hotmail.com;mike@marshallconstruction.ca"
    .ReminderSet = False
    .MeetingStatus = olMeeting
    .ForceUpdateToAllAttendees = True
    
    'prevent corrupt items being added
    If ExceltoCalendarModule.countItems > 1024 Then
        MsgBox prompt:="Item.Count= " & countItems, _
        Buttons:=vbOKOnly + vbCritical, _
        Title:="Add Tasks"
        GoTo ErrorHandler
    End If
    .Send
ErrorHandler:
End With
Set OLAI = Nothing
Set OLApp = Nothing
End Sub
