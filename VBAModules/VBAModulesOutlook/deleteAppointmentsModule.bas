Attribute VB_Name = "deleteAppointmentsModule"
Option Explicit
Sub deleteTaskDeletedOnline()

    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objFolder As Outlook.MAPIFolder
    Dim objAppointment As Outlook.AppointmentItem
    Dim lngDeletedAppointements As Long
    Dim blnRestart As Boolean
    
    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)
       
Here:
    blnRestart = False
    For Each objAppointment In objFolder.Items
        DoEvents
        On Error Resume Next
        If objAppointment.Categories = "Deleted" Then
            objAppointment.Delete
            lngDeletedAppointements = lngDeletedAppointements + 1
            blnRestart = True
        End If
    Next
    If blnRestart = True Then GoTo Here     'thi goes back to top of code
    If lngDeletedAppointements <> 0 Then
        MsgBox prompt:=lngDeletedAppointements & " deleted!", _
        Buttons:=vbOKOnly + vbInformation, _
        Title:="Delete Appointments"
    End If
ExitSub:
End Sub

Sub deleteCorruptAppointments()
Dim objOutlook As Outlook.Application
Dim objNamespace As Outlook.NameSpace
Dim objFolder As Outlook.MAPIFolder
Dim objAppointment As Outlook.AppointmentItem
Dim lngDeletedAppointements As Long
Dim corruptStart As String
Dim defaultStart As String
Dim blnRestart As Boolean
    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)
    
    defaultStart = Now
    corruptStart = InputBox("Input Start Date of Corrupt Appointments", , defaultStart)
    
Here:
    blnRestart = False

    'If corruptStart = "" Then Exit Sub
    'Debug.Print objFolder.Items.Count
        For Each objAppointment In objFolder.Items
            DoEvents
            On Error Resume Next
            If objAppointment.Subject = "" Then
                objAppointment.Delete
                lngDeletedAppointements = lngDeletedAppointements + 1
                blnRestart = True
            End If
            'Debug.Print objAppointment.Start
            'delete bad dates, that is dates with no date, just 12:00 AM.
            If objAppointment.Start = corruptStart Or objAppointment.Start = "12:00 AM" Then
                objAppointment.Delete
                lngDeletedAppointements = lngDeletedAppointements + 1
                blnRestart = True
                'Debug.Print lngDeletedAppointements & " delete!"
            End If
        Next
    If blnRestart = True Then GoTo Here     'thi goes back to top of code

        MsgBox prompt:=lngDeletedAppointements & " done!", _
        Buttons:=vbOKOnly + vbInformation, _
        Title:="Delete Corrupted Appointments"
        'Debug.Print lngDeletedAppointements & " Done!"
ExitSub:
End Sub

Sub deleteJob()
Dim objOutlook As Outlook.Application
Dim objNamespace As Outlook.NameSpace
Dim objFolder As Outlook.MAPIFolder
Dim objAppointment As Outlook.AppointmentItem
Dim objAttachment As Outlook.Attachment
Dim objNetwork As Object
Dim lngDeletedAppointements As Long
Dim lngCleanedAppointements As Long
Dim lngCleanedAttachments As Long
Dim blnRestart As Boolean
Dim intDateDiff As Integer
Dim tasksJobNumber As String
Dim jobNumber As String
Dim oldDifference As Integer

Set objOutlook = Application
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)

jobNumber = InputBox("Delete all tasks for this job," & vbCr & "Enter Job#:", "Delete Job", "12-0-xxxx")
'oldDifference = dateDiff("d", startDate, Now)

Here:
    blnRestart = False
    For Each objAppointment In objFolder.Items
        DoEvents
        'On Error Resume Next
        'intDateDiff = dateDiff("d", objAppointment.Start, Now)
        
        'parse .Description for Job#
        tasksJobNumber = Left(objAppointment.Subject, 9)
        
        'delete bad dates, that is dates with no date, just 12:00 AM.
        'If objAppointment.Start = "12:00 AM" Then objAppointment.Delete
        
        ' Delete task for Job#
        If tasksJobNumber = jobNumber And objAppointment.RecurrenceState = olApptNotRecurring Then
        
        'confirm delete
        'Dim Response As VbMsgBoxResult
        'Response = MsgBox(objAppointment.Subject & ", matches the input Job: " & jobNumber & vbCr & "Delete this task?", vbQuestion + vbYesNo)
        'If Response = vbNo Then Exit Sub
        
        objAppointment.Delete
        lngDeletedAppointements = lngDeletedAppointements + 1
        blnRestart = True
        
        ' Delete attachments from 6-month-old appointments.
        ElseIf intDateDiff > 180 And objAppointment.RecurrenceState = olApptNotRecurring Then
        
        'confirm delete
        'Response = MsgBox("Attachments from 6-month-old found, delete?", vbQuestion + vbYesNo)
        'If Response = vbNo Then Exit Sub
        
        
        If objAppointment.Attachments.Count > 0 Then
            While objAppointment.Attachments.Count > 0
                objAppointment.Attachments.Remove 1
            Wend
            lngCleanedAppointements = lngCleanedAppointements + 1
        End If
        
        ' Delete large attachments from 60-day-old appointments.
        ElseIf intDateDiff > 60 Then
        
        'confirm delete
        'Response = MsgBox("Large attachments from 60-day-old appointments found, delete?", vbQuestion + vbYesNo)
        'If Response = vbNo Then Exit Sub
        
            If objAppointment.Attachments.Count > 0 Then
                For Each objAttachment In objAppointment.Attachments
                    If objAttachment.Size > 500000 Then
                        objAttachment.Delete
                        lngCleanedAttachments = lngCleanedAttachments + 1
                    End If
                Next
            End If
        End If
    Next
    If blnRestart = True Then GoTo Here     'thi goes back to top of code
    
    MsgBox prompt:="Deleted " & lngDeletedAppointements & " appointment(s)." & vbCrLf & _
    "Cleaned " & lngCleanedAppointements & " appointment(s)." & vbCrLf & _
    "Deleted " & lngCleanedAttachments & _
    " attachment(s).", _
    Buttons:=vbOKOnly + vbInformation, _
    Title:="Delete Job"
    
    'Run Scheduler after deleting Job to show updated Schedule
        With greenBoard
        .removeJobTasks
        .addJobTasks
        .addTasksHandler
        .addResizeHandler
        .Show vbModeless
    End With

End Sub

Sub deleteOldAppointments()
Dim objOutlook As Outlook.Application
Dim objNamespace As Outlook.NameSpace
Dim objFolder As Outlook.MAPIFolder
Dim objAppointment As Outlook.AppointmentItem
Dim objAttachment As Outlook.Attachment
Dim objNetwork As Object
Dim lngDeletedAppointements As Long
Dim lngCleanedAppointements As Long
Dim lngCleanedAttachments As Long
Dim blnRestart As Boolean
Dim intDateDiff As Integer
Dim startDate As String
Dim oldDifference As Integer

Set objOutlook = Application
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)

startDate = InputBox("Delete appointments older than," & vbCr & "Enter Date:", "Delete Old Appointments")
On Error GoTo ExitSub
oldDifference = dateDiff("d", startDate, Now)

Here:
    blnRestart = False
    For Each objAppointment In objFolder.Items
        DoEvents
        On Error Resume Next
        intDateDiff = dateDiff("d", objAppointment.Start, Now)
        
        'delete bad dates, that is dates with no date, just 12:00 AM.
        If objAppointment.Start = "12:00 AM" Then objAppointment.Delete
        
        ' Delete year-old appointments.
        If intDateDiff > oldDifference And objAppointment.RecurrenceState = olApptNotRecurring Then
        
        'confirm delete
        'Dim Response As VbMsgBoxResult
        'Response = MsgBox(intDateDiff & " Day-old appointments, " & objAppointment.Subject & " found, delete?", vbQuestion + vbYesNo)
        'If Response = vbNo Then Exit Sub
        
        objAppointment.Delete
        lngDeletedAppointements = lngDeletedAppointements + 1
        blnRestart = True
        
        ' Delete attachments from 6-month-old appointments.
        ElseIf intDateDiff > 180 And objAppointment.RecurrenceState = olApptNotRecurring Then
        
        'confirm delete
        'Response = MsgBox("Attachments from 6-month-old found, delete?", vbQuestion + vbYesNo)
        'If Response = vbNo Then Exit Sub
        
        
        If objAppointment.Attachments.Count > 0 Then
            While objAppointment.Attachments.Count > 0
                objAppointment.Attachments.Remove 1
            Wend
            lngCleanedAppointements = lngCleanedAppointements + 1
        End If
        
        ' Delete large attachments from 60-day-old appointments.
        ElseIf intDateDiff > 60 Then
        
        'confirm delete
        'Response = MsgBox("Large attachments from 60-day-old appointments found, delete?", vbQuestion + vbYesNo)
        'If Response = vbNo Then Exit Sub
        
            If objAppointment.Attachments.Count > 0 Then
                For Each objAttachment In objAppointment.Attachments
                    If objAttachment.Size > 500000 Then
                        objAttachment.Delete
                        lngCleanedAttachments = lngCleanedAttachments + 1
                    End If
                Next
            End If
        End If
    Next
    If blnRestart = True Then GoTo Here     'thi goes back to top of code
    
    MsgBox prompt:="Deleted " & lngDeletedAppointements & " appointment(s)." & vbCrLf & _
    "Cleaned " & lngCleanedAppointements & " appointment(s)." & vbCrLf & _
    "Deleted " & lngCleanedAttachments & _
    " attachment(s).", _
    Buttons:=vbOKOnly + vbInformation, _
    Title:="Delete Old Jobs"
ExitSub:
End Sub

Sub deleteAllTasks()
Dim objOutlook As Outlook.Application
Dim objNamespace As Outlook.NameSpace
Dim objFolder As Outlook.MAPIFolder
Dim objAppointment As Outlook.AppointmentItem
Dim objAttachment As Outlook.Attachment
'Dim objNetwork As Object
Dim lngDeletedAppointements As Long
Dim lngCleanedAppointements As Long
Dim lngCleanedAttachments As Long
Dim blnRestart As Boolean
Dim intDateDiff As Integer
Dim startDate As String
Dim oldDifference As Integer

Set objOutlook = Application
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)

Here:
    blnRestart = False
    'Debug.Print objFolder.Items.Count
    For Each objAppointment In objFolder.Items
        DoEvents
        On Error Resume Next
        objAppointment.Delete
        lngDeletedAppointements = lngDeletedAppointements + 1
        blnRestart = True
        
        If objAppointment.Attachments.Count > 0 Then
            While objAppointment.Attachments.Count > 0
                objAppointment.Attachments.Remove 1
            Wend
        '    lngCleanedAppointements = lngCleanedAppointements + 1
        End If
        

    Next
    If blnRestart = True Then GoTo Here     'this goes back to top of code
    'Debug.Print "Deleted " & lngDeletedAppointements & " appointment(s)."
    MsgBox "Deleted " & lngDeletedAppointements & " appointment(s)."
ExitSub:
End Sub


