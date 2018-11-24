VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} controlPanel 
   Caption         =   "Control Panel"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   7980
   ClientWidth     =   5055
   OleObjectBlob   =   "controlPanel.frx":0000
End
Attribute VB_Name = "controlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub categoriesButton_Click()
    categoriesModule.addStandardCategories
End Sub

Private Sub downloadButton_Click()
    ExceltoCalendarModule.ftpDownload
    controlPanel.importFromExcelButton.BackColor = &HFF00FF
End Sub

Private Sub ftpButton_Click()
    CalendarToExcelModule.showFTPForm   '.transferCalendarToWeb
End Sub

Private Sub runButton_Click()
    'colorChanged = False
    With greenBoard
        .removeJobTasks
        .addJobTasks
        .addTasksHandler
        .addResizeHandler
        .Show vbModeless
    End With
End Sub

Private Sub reportJobTasksButton_Click()
    jobTasksReportModule.reportAppointments
End Sub

Private Sub editJobTasksButton_Click()
    Dim indexNumber
    On Error Resume Next
    With greenBoard
        If .ActiveControl.Name = "scheduleButton" Or .ActiveControl.Name = "cursorTextBox" Then
            MsgBox "A task must be selected before editing. Run the Scheduler." & vbCr _
            & "Click on a Green Board, Task to select it for editing."
            displayModule.runScheduler
        Else
            taskAppointmentForm.Show vbModeless
            indexNumber = Right(.ActiveControl.Name, InStr(.ActiveControl.Name, "Task"))
            jobTasksAddModule.fillinForm indexNumber
        End If
    End With
End Sub

Private Sub addJobTasksButton_Click()
    jobTasksAddModule.fillinForm
End Sub

Private Sub deleteJobTasksButton_Click()
    displayModule.deleteJobTasks
End Sub

Private Sub exportToExcelButton_Click()
    CalendarToExcelModule.SaveCalendarToExcel
End Sub

Private Sub importFromExcelButton_Click()
    ExceltoCalendarModule.ImportAppointments
End Sub

Private Sub newJobButton_Click()
    newJobForm.Show vbModeless
    newJobForm.modifyCheckBox.value = False
    newJobForm.modifyCheckBox.Visible = False
    newJobForm.checkboxLabel.Visible = False
End Sub

Private Sub deleteOldJobsButton_Click()
    deleteAppointmentsModule.deleteOldAppointments
End Sub

Private Sub deleteTaskJobButton_Click()
    deleteAppointmentsModule.deleteJob
End Sub
