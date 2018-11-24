Attribute VB_Name = "ExceltoCalendarModule"
Option Explicit

Sub ftpDownload()
    Dim thisDownload As downloadClassLibrary.downloadComClass
    Set thisDownload = New downloadClassLibrary.downloadComClass
    
    CreateObject("WScript.Shell").PopUp "The download might take a minute. Please wait until" & vbCr & _
    "the 'Download Complete' message before" & vbCr & "starting the Import.", _
        2, "FTP Download", vbOKOnly + vbExclamation
        
    'the download assembly requires two parameters path and name
    'Path is found using the relativePath process; Name is hard coded.
    thisDownload.ftpDownloadWebRequest relativePathModule.myPath, "\My Projects\Project.Scheduler\calendar.xls"
End Sub

Sub callCountItems()
    MsgBox prompt:="There are " & countItems & " items on this calendar.", _
      Buttons:=vbOKOnly + vbInformation, _
      Title:="Count Appointments"
End Sub
Function countItems() As Integer
    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objFolder As Outlook.MAPIFolder
    Dim objAppointment As Outlook.AppointmentItem
    Dim lngFoundAppointements As Long
    Dim blnRestart As Boolean
    Dim tempStr As String
    
    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)
    countItems = 0
    countItems = objFolder.Items.Count
    'MsgBox "There are " & countItems & " items on this calendar."
ExitSub:
End Function

Sub ImportAppointments()
'http://www.outlookcode.com/codedetail.aspx?id=788
Dim startTime As Double
Dim endTime As Double
startTime = Timer
On Error Resume Next
connectToExcel:
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim calendarData As String
  calendarData = relativePathModule.myPath & "\My Projects\Project.Scheduler\calendar.xls"
  Set db = OpenDatabase(Name:=calendarData, Options:=False, ReadOnly:=False, Connect:="Excel 8.0")
  Set rs = db.OpenRecordset("SELECT * FROM `jobSchedule` WHERE Subject <> null")
  Dim totalCount As Long
  rs.MoveLast
  totalCount = rs.recordCount
  rs.MoveFirst
  
createItems:
    Dim itmAppt As Outlook.AppointmentItem
    Dim objFolder As Outlook.MAPIFolder
    Dim aptPtrn As Outlook.RecurrencePattern
    Dim fso As FileSystemObject
    Dim fl As File
    Dim iRecord As Long
    Dim tmpItm As Outlook.Link
    Dim mpiFolder As MAPIFolder
    Dim oNs As NameSpace
    Set oNs = Outlook.GetNamespace("MAPI")
    Set mpiFolder = oNs.GetDefaultFolder(9)
    Dim tempSubject As String
    Dim lastRecord As String
    While Not rs.EOF
        Set itmAppt = Outlook.CreateItem(olAppointmentItem)
        On Error Resume Next
        tempSubject = rs.Fields(6).Value 'Subject text
        Call findTask(tempSubject)  'check that task is not already on the calendar, do not duplicate

        itmAppt.Start = rs.Fields(3).Value
        itmAppt.End = rs.Fields(4).Value
        itmAppt.Duration = rs.Fields(5).Value
        
        itmAppt.Subject = rs.Fields(6).Value
        If itmAppt.Subject = "" Then
            MsgBox prompt:="Subject is absent.", _
            Buttons:=vbOKOnly + vbCritical, _
            Title:="Import Appointments"
            GoTo ErrorHandler
        End If

        itmAppt.Location = rs.Fields(7).Value
        itmAppt.Categories = rs.Fields(8).Value
        itmAppt.Body = rs.Fields(9).Value
        itmAppt.RequiredAttendees = rs.Fields(10).Value
        
        itmAppt.AllDayEvent = False
        itmAppt.ReminderSet = False
        itmAppt.ReminderMinutesBeforeStart = 30
        If itmAppt.Subject = lastRecord Then
            MsgBox prompt:=itmAppt.Subject & " Duplicate Task!", _
              Buttons:=vbOKOnly + vbCritical, _
              Title:="Import Appointments"
            GoTo ErrorHandler
        End If
    If countItems > totalCount Then
        GoTo ErrorHandler
    End If
        itmAppt.Save

        iRecord = iRecord + 1
        If iRecord = totalCount + 1 Then GoTo ErrorHandler
        lastRecord = itmAppt.Subject
        rs.MoveNext
    Wend
    endTime = Timer
    Debug.Print Format(endTime - startTime, "##.00000")
    deleteAppointmentsModule.deleteTaskDeletedOnline
    MsgBox prompt:="Import Complete: " & countItems, _
      Buttons:=vbOKOnly + vbInformation, _
      Title:="Import Appointments"
      
ErrorHandlerExit:
   GoTo ExitSub
ErrorHandler:
  MsgBox prompt:="Error No: " & Err.Number & "; Description: ", _
    Buttons:=vbOKOnly + vbCritical, _
    Title:="Import Appointments"

  Resume ErrorHandlerExit
ExitSub:
  rs.Close
  db.Close
  Set rs = Nothing
  Set db = Nothing
End Sub


Function findTask(ByVal taskSubject As String) As Boolean
'check to see if task is on calendar, if found- delete it
    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objFolder As Outlook.MAPIFolder
    Dim objAppointment As Outlook.AppointmentItem
    
    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)
    
        For Each objAppointment In objFolder.Items
            DoEvents
            On Error Resume Next
            If objAppointment.Subject = taskSubject Then
                'delete task
                jobTasksDeleteModule.deleteTask taskSubject
                Exit Function
            End If
        Next
ExitSub:
End Function

Function findJob(ByVal jobNumber As String) As Boolean
Dim objOutlook As Outlook.Application
Dim objNamespace As Outlook.NameSpace
Dim objFolder As Outlook.MAPIFolder
Dim objAppointment As Outlook.AppointmentItem
Dim lngFoundAppointements As Long
Dim blnRestart As Boolean
Dim tempStr As String

Set objOutlook = Application
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)

    For Each objAppointment In objFolder.Items
        DoEvents
        On Error Resume Next
        tempStr = Left(objAppointment.Subject, 9)
        If tempStr = jobNumber Then
            'MsgBox "Found match " & objAppointment.Subject
            findJob = True
            Exit Function
        End If
    Next
ExitSub:
End Function
Private Sub testModifyJob()
  Dim jobNumber As String
  Dim clientName As String
  Dim startDate As Date
  jobNumber = "12-0-0506"
  clientName = "Test ClientName"
  startDate = #5/6/2012#
  Call ModifyNewJob(jobNumber, clientName, startDate)
End Sub
Sub ModifyNewJob(ByVal jobNumber As String, ByVal clientName As String, ByVal startDate As Date)
'http://www.outlookcode.com/codedetail.aspx?id=788
Dim exlApp As excel.Application
Dim exlWkb As Workbook
Dim exlSht As Worksheet
Dim rng As Range

Dim itmAppt As Outlook.AppointmentItem
Dim aptPtrn As Outlook.RecurrencePattern
Dim fso As FileSystemObject
Dim fl As File
Dim strFilePath As String

'before adding NewJob to Outlook check to see if it is already on the board
If findJob(jobNumber) Then
    MsgBox prompt:="Tasks with " & jobNumber & " are already Scheduled." & vbCr & _
    "Please use another Job Number or delete the Job with " & jobNumber & ".", _
    Buttons:=vbOKOnly + vbExclamation, _
    Title:="Edit New Job Template"
    Exit Sub
End If

'get open instance of Excel, if not Open, then Set New instance
On Error Resume Next
Set exlApp = GetObject(, "Excel.Application")
If Err Then
    Set exlApp = New excel.Application
    'ExcelWasNotRunning = True
End If

'this opens dialog box to access the NewJob.xlsm file
'these are the hard-coded paths on Marshall and Mike for the NewJob file.
    'strFilePath = "C:\Users\mike.MARSHALLBROTHER\Documents\My Projects\Project.Scheduler\newJob.xlsm"
    'C:\Documents and Settings\Mike Lanham\My Documents\My Projects\Project.Scheduler
    strFilePath = relativePathModule.myPath & "\My Projects\Project.Scheduler\newJob.xlsm"
'strFilePath = exlApp.GetOpenFilename
If strFilePath = "" Then
    exlApp.Quit
    Set exlApp = Nothing
    Exit Sub
End If

Set exlSht = excel.Application.Workbooks.Open(strFilePath).Worksheets(1)

exlApp.Visible = True
Dim iRow As Integer
Dim iCol As Integer

Dim tmpItm As Outlook.Link
Dim mpiFolder As MAPIFolder
Dim oNs As NameSpace
Set oNs = Outlook.GetNamespace("MAPI")
Set mpiFolder = oNs.GetDefaultFolder(9)
iRow = 3
iCol = 1
With exlSht
    .Range("D1").Value = jobNumber
    .Range("E1").Value = clientName
    .Range("F1").Value = startDate
    .Activate
End With
With ActiveWorkbook.Worksheets(2)
    .Range("D1").Value = jobNumber
    .Range("E1").Value = clientName
    .Range("F1").Value = startDate
    .Activate
End With

excel.Application.Run "newJobModule.messageBox"

'Dim itmAppt As Outlook.AppointmentItem
'While sheet1.Cells(iRow, 1) <> ""
'    Dim cnct As ContactItem
'    Set itmAppt = Outlook.CreateItem(olAppointmentItem)
'    itmAppt.Start = exlSht.Cells(iRow, 1)
'    itmAppt.End = exlSht.Cells(iRow, 2)
'    itmAppt.Subject = exlSht.Cells(iRow, 4)
'    itmAppt.Categories = exlSht.Cells(iRow, 6)
'    itmAppt.AllDayEvent = False
'    itmAppt.ReminderSet = True
'    itmAppt.ReminderMinutesBeforeStart = 30
'    itmAppt.Save
 '   iRow = iRow + 1
'Wend
'Excel.Application.Workbooks.Close
'exlApp.Quit
Set exlSht = Nothing
Set exlWkb = Nothing
Set exlApp = Nothing
'newJobForm.Show vbModeless

End Sub
Sub testDAO()
  ImportNewJob "12-0-0521", "test dao", #5/21/2012#
End Sub
Sub ImportNewJobDAO(ByVal jobNumber As String, ByVal clientName As String, ByVal startDate As Date)
Dim startTime As Double
Dim endTime As Double
startTime = Timer
On Error Resume Next
connectToExcel:
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim calendarData As String
  calendarData = relativePathModule.myPath & "\My Projects\Project.Scheduler\newJob.xls"
  Set db = OpenDatabase(Name:=calendarData, Options:=False, ReadOnly:=False, Connect:="Excel 8.0")
  Set rs = db.OpenRecordset("SELECT * FROM `standardTasks` WHERE Subject <> null")
  Dim totalCount As Long
  rs.MoveLast
  totalCount = rs.recordCount
  rs.MoveFirst
  Dim iRecord As Long
  Dim itmAppt As Outlook.AppointmentItem
  Dim objFolder As Outlook.MAPIFolder
  
  Dim aptPtrn As Outlook.RecurrencePattern
  Dim fso As FileSystemObject
  Dim fl As File
  Dim strFilePath As String

'before adding NewJob to Outlook check to see if it is already on the board
If findJob(jobNumber) Then
    MsgBox prompt:="Tasks with " & jobNumber & " are already Scheduled." & vbCr & _
    "Please use another Job Number or delete the Job with " & jobNumber & ".", _
    Buttons:=vbOKOnly + vbExclamation, _
    Title:="Add New Job"
    Exit Sub
End If

'get open instance of Excel, if not Open, then Set New instance
'On Error Resume Next
'Set exlApp = GetObject(, "Excel.Application")
'If Err Then
'    Set exlApp = New Excel.Application
'    'ExcelWasNotRunning = True
'End If

'strFilePath = relativePathModule.myPath & "\My Projects\Project.Scheduler\newJob.xlsm"
'Debug.Print strFilePath
'strFilePath = "C:\Documents and Settings\Mike Lanham\My Documents" & "\My Projects\Project.Scheduler\newJob.xlsm"
'Debug.Print strFilePath

'If strFilePath = "" Then
'    exlApp.GetOpenFilename
'End If
'
'Set exlSht = Excel.Application.Workbooks.Open(strFilePath).Worksheets(1)
'exlApp.Visible = False

'Dim iRow As Integer
'Dim iCol As Integer
Dim tmpItm As Outlook.Link
Dim mpiFolder As MAPIFolder
Dim oNs As NameSpace
Set oNs = Outlook.GetNamespace("MAPI")
Set mpiFolder = oNs.GetDefaultFolder(9)
Dim cnct As ContactItem
'iRow = 3
'iCol = 1
'With exlSht
'    .Range("D1").Value = jobNumber
'    .Range("E1").Value = clientName
'    .Range("F1").Value = startDate
'End With
'With ActiveWorkbook.Worksheets(2)
'    .Range("D1").Value = jobNumber
'    .Range("E1").Value = clientName
'    .Range("F1").Value = startDate
'End With

'Excel.Application.Run "newJobModule.calculateDates"

'While exlSht.Cells(iRow, 1) <> ""
  While Not rs.EOF
  'Fields(index) is zero based
    Set itmAppt = Outlook.CreateItem(olAppointmentItem)
    itmAppt.Start = rs.Fields(0).Value  'exlSht.Cells(iRow, 1)
    itmAppt.End = rs.Fields(1).Value  'exlSht.Cells(iRow, 2)
    'itmAppt.Duration = rs.Fields(4).Value  'exlSht.Cells(iRow, 4)
    itmAppt.Subject = rs.Fields(3).Value  'exlSht.Cells(iRow, 4)
    If itmAppt.Subject = "" Then
      MsgBox "Subject is absent."
      GoTo ErrorHandler
    End If
    'Set cnct = mpiFolder.Items.Find("[FullName] = " & exlSht.Cells(iRow, 8))
    'If cnct Is Nothing Then
    '    Set cnct = Outlook.CreateItem(olContactItem)
    '    cnct.FullName = exlSht.Cells(iRow, 8)
    '    cnct.Save
    'End If
    
    itmAppt.Location = rs.Fields(4).Value  'exlSht.Cells(iRow, 5)
    itmAppt.Categories = rs.Fields(5).Value  'exlSht.Cells(iRow, 6)
    'itmAppt.Body = rs.Fields(8).Value  'exlSht.Cells(iRow, 8)
    itmAppt.RequiredAttendees = rs.Fields(7).Value  'exlSht.Cells(iRow, 7)
    
    itmAppt.AllDayEvent = False
    'itmAppt.Links.Add cnct
    
    'Set aptPtrn = itmAppt.GetRecurrencePattern
    'aptPtrn.startTime = exlSht.Cells(iRow, 7)
    'aptPtrn.EndTime = exlSht.Cells(iRow, 6)
    'aptPtrn.RecurrenceType = olRecursYearly
    'aptPtrn.NoEndDate = True
    
    'If aptPtrn.Duration > 1440 Then aptPtrn.Duration = aptPtrn.Duration - 1440
    'Select Case exlSht.Cells(iRow, 7)
    'Case "No Reminder"
    '    itmAppt.ReminderSet = False
    'Case "30 minutes"
      itmAppt.ReminderSet = False
      itmAppt.ReminderMinutesBeforeStart = 30
    'Case "1 day"
    '    itmAppt.ReminderSet = True
    '    itmAppt.ReminderMinutesBeforeStart = 1440
    'Case "2 days"
    '    itmAppt.ReminderSet = True
    '    itmAppt.ReminderMinutesBeforeStart = 2880
    'Case "1 week"
    '    itmAppt.ReminderSet = True
    '    itmAppt.ReminderMinutesBeforeStart = 10080
    'End Select
    If iRecord > totalCount Then
      MsgBox "Count= " & iRecord
      GoTo ErrorHandler
    End If
    itmAppt.Save
    iRecord = iRecord + 1
    rs.MoveNext
    Wend
    endTime = Timer
    Debug.Print Format(endTime - startTime, "##.00000")
    MsgBox prompt:="Import Complete: " & countItems, _
    Buttons:=vbOKOnly + vbInformation, _
    Title:="Add New Job"
      
ErrorHandlerExit:
   GoTo ExitSub
ErrorHandler:
  MsgBox prompt:="Error No: " & Err.Number & "; Description: ", Buttons:=vbOKOnly + vbCritical
  Resume ErrorHandlerExit
ExitSub:
  rs.Close
  db.Close
  Set rs = Nothing
  Set db = Nothing
End Sub
Sub ImportNewJob(ByVal jobNumber As String, ByVal clientName As String, ByVal startDate As Date)
'http://www.outlookcode.com/codedetail.aspx?id=788
Dim exlApp As excel.Application
'Dim exlWkb As Workbook
Dim exlSht As Worksheet
Dim rng As Range
Dim itmAppt As Outlook.AppointmentItem
Dim objFolder As Outlook.MAPIFolder

Dim aptPtrn As Outlook.RecurrencePattern
Dim fso As FileSystemObject
Dim fl As File
Dim strFilePath As String

'before adding NewJob to Outlook check to see if it is already on the board
If findJob(jobNumber) Then
    MsgBox prompt:="Tasks with " & jobNumber & " are already Scheduled." & vbCr & _
    "Please use another Job Number or delete the Job with " & jobNumber & ".", _
    Buttons:=vbOKOnly + vbExclamation, _
    Title:="Add New Job"
    Exit Sub
End If

'get open instance of Excel, if not Open, then Set New instance
On Error Resume Next
Set exlApp = GetObject(, "Excel.Application")
If Err Then
    Set exlApp = New excel.Application
    'ExcelWasNotRunning = True
End If

strFilePath = relativePathModule.myPath & "\My Projects\Project.Scheduler\newJob.xlsm"
'Debug.Print strFilePath
'strFilePath = "C:\Documents and Settings\Mike Lanham\My Documents" & "\My Projects\Project.Scheduler\newJob.xlsm"
'Debug.Print strFilePath

If strFilePath = "" Then
    exlApp.GetOpenFilename
End If

Set exlSht = excel.Application.Workbooks.Open(strFilePath).Worksheets(1)
exlApp.Visible = False

Dim iRow As Integer
Dim iCol As Integer
Dim tmpItm As Outlook.Link
Dim mpiFolder As MAPIFolder
Dim oNs As NameSpace
Set oNs = Outlook.GetNamespace("MAPI")
Set mpiFolder = oNs.GetDefaultFolder(9)
iRow = 3
iCol = 1
With exlSht
    .Range("D1").Value = jobNumber
    .Range("E1").Value = clientName
    .Range("F1").Value = startDate
End With
With ActiveWorkbook.Worksheets(2)
    .Range("D1").Value = jobNumber
    .Range("E1").Value = clientName
    .Range("F1").Value = startDate
End With

excel.Application.Run "newJobModule.calculateDates"

While exlSht.Cells(iRow, 1) <> ""
    Dim cnct As ContactItem
    Set itmAppt = Outlook.CreateItem(olAppointmentItem)
    itmAppt.Start = exlSht.Cells(iRow, 1)
    itmAppt.End = exlSht.Cells(iRow, 2)
    'itmAppt.Duration = exlSht.Cells(iRow, 4)
    itmAppt.Subject = exlSht.Cells(iRow, 4)
    If itmAppt.Subject = "" Then
        MsgBox "Subject is absent."
        GoTo ErrorHandler
    End If
    'Set cnct = mpiFolder.Items.Find("[FullName] = " & exlSht.Cells(iRow, 8))
    'If cnct Is Nothing Then
    '    Set cnct = Outlook.CreateItem(olContactItem)
    '    cnct.FullName = exlSht.Cells(iRow, 8)
    '    cnct.Save
    'End If
    
    itmAppt.Location = exlSht.Cells(iRow, 5)
    itmAppt.Categories = exlSht.Cells(iRow, 6)
    'itmAppt.Body = exlSht.Cells(iRow, 8)
    itmAppt.RequiredAttendees = exlSht.Cells(iRow, 7)

    itmAppt.AllDayEvent = False
    'itmAppt.Links.Add cnct
    
    'Set aptPtrn = itmAppt.GetRecurrencePattern
    'aptPtrn.startTime = exlSht.Cells(iRow, 7)
    'aptPtrn.EndTime = exlSht.Cells(iRow, 6)
    'aptPtrn.RecurrenceType = olRecursYearly
    'aptPtrn.NoEndDate = True
    
    'If aptPtrn.Duration > 1440 Then aptPtrn.Duration = aptPtrn.Duration - 1440
    'Select Case exlSht.Cells(iRow, 7)
    'Case "No Reminder"
    '    itmAppt.ReminderSet = False
    'Case "30 minutes"
        itmAppt.ReminderSet = False
        itmAppt.ReminderMinutesBeforeStart = 30
    'Case "1 day"
    '    itmAppt.ReminderSet = True
    '    itmAppt.ReminderMinutesBeforeStart = 1440
    'Case "2 days"
    '    itmAppt.ReminderSet = True
    '    itmAppt.ReminderMinutesBeforeStart = 2880
    'Case "1 week"
    '    itmAppt.ReminderSet = True
    '    itmAppt.ReminderMinutesBeforeStart = 10080
    'End Select
    If iRow > 20 Then
        MsgBox "Count= " & iRow
        GoTo ErrorHandler
    End If
    itmAppt.Save
    iRow = iRow + 1
Wend
'Excel.Application.DisplayAlerts = False
'    Excel.Application.Workbooks.Close
'Excel.Application.DisplayAlerts = True
'
'exlApp.Quit
'Set exlSht = Nothing
''Set exlWkb = Nothing
'Set exlApp = Nothing

ErrorHandlerExit:
   GoTo ExitSub

ErrorHandler:
   If Err.Number = 429 Then
      'Application object is not set by GetObject; use CreateObject instead
      If exlApp Is Nothing Then
         Set exlApp = CreateObject("Excel.Application")
         Resume Next
      End If
   Else
      MsgBox prompt:="Error No: " & Err.Number & "; Description: ", _
      Buttons:=vbOKOnly + vbCritical, _
      Title:="Add New Job"
      Resume ErrorHandlerExit
   End If
ExitSub:
    With exlApp
    excel.Application.DisplayAlerts = False
        .Workbooks.Close
    excel.Application.DisplayAlerts = True
        .Quit
    End With
    Set exlApp = Nothing
End Sub

Sub sortExcelTable()
    Dim exlApp As excel.Application
    Dim exlSht As Worksheet
    Dim strFilePath As String
    Set exlApp = New excel.Application
    strFilePath = relativePathModule.myPath & "\My Projects\Project.Scheduler\calendar.xls"
    Set exlSht = excel.Application.Workbooks.Open(strFilePath).Worksheets(1)
    exlApp.Visible = True
    
    'need to sort out empty rows, need contiguous block of data in table
    exlSht.Range("JobSchedule").Select
        ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("A:A") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("D:D") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("JobSchedule")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    excel.Application.DisplayAlerts = False
        'Excel.Application.Workbooks.Close SaveChanges:=True
            ActiveWorkbook.Close SaveChanges:=True

    excel.Application.DisplayAlerts = True
    
    exlApp.Quit
    Set exlSht = Nothing
    Set exlApp = Nothing

End Sub

Sub ImportAppointmentsOpenExcel()
'http://www.outlookcode.com/codedetail.aspx?id=788
Dim startTime As Double
Dim endTime As Double
startTime = Timer

openExcel:
    Dim exlApp As excel.Application
    Dim exlSht As Worksheet
    Dim strFilePath As String
    
    'get open instance of Excel, if not Open, then Set New instance
    On Error Resume Next
    Set exlApp = GetObject(, "Excel.Application")
    If Err Then
        Set exlApp = New excel.Application
        'ExcelWasNotRunning = True
    End If
    
    strFilePath = relativePathModule.myPath & "\My Projects\Project.Scheduler\calendar.xls"
    Set exlSht = excel.Application.Workbooks.Open(strFilePath).Worksheets(1)
    exlApp.Visible = False

createItems:
    Dim rng As Range
    Dim itmAppt As Outlook.AppointmentItem
    Dim objFolder As Outlook.MAPIFolder
    Dim aptPtrn As Outlook.RecurrencePattern
    Dim fso As FileSystemObject
    Dim fl As File
    'Dim cnct As CalendarItem

    Dim iRow As Integer
    Dim iCol As Integer
    Dim tmpItm As Outlook.Link
    Dim mpiFolder As MAPIFolder
    Dim oNs As NameSpace
    Set oNs = Outlook.GetNamespace("MAPI")
    Set mpiFolder = oNs.GetDefaultFolder(9)
    iRow = 3
    iCol = 1

    'sort out empty rows, to get contiguous block of data in table
    exlSht.Range("JobSchedule").Select
        ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("A:A") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("D:D") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("JobSchedule")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Save
    Dim tempSubject As String
    Dim tempUpdated As Date
    While exlSht.Cells(iRow, 7) <> ""   'Subject required
        tempSubject = exlSht.Cells(iRow, 7) 'Subject text
        'tempUpdated = exlSht.Cells(iRow, 2) 'Updated date
        
        'this calls function, findTask. If match found it deletes the task
        Call findTask(tempSubject)
        
        Set itmAppt = Outlook.CreateItem(olAppointmentItem)
        'Debug.Print "New Task " & exlSht.Cells(iRow, 7)
        'itmAppt.GlobalAppointmentID = exlSht.Cells(iRow, 1)
        'itmAppt.LastModificationTime = exlSht.Cells(iRow, 2)
        'itmAppt.Created = exlSht.Cells(iRow, 3)
        itmAppt.Start = exlSht.Cells(iRow, 4)
        itmAppt.End = exlSht.Cells(iRow, 5)
        itmAppt.Duration = exlSht.Cells(iRow, 6)
        
        itmAppt.Subject = exlSht.Cells(iRow, 7)
        If itmAppt.Subject = "" Then
            MsgBox "Subject is absent."
            GoTo ErrorHandler
        End If

        itmAppt.Location = exlSht.Cells(iRow, 8)
        itmAppt.Categories = exlSht.Cells(iRow, 9)
        itmAppt.Body = exlSht.Cells(iRow, 10)
        itmAppt.RequiredAttendees = exlSht.Cells(iRow, 11)
        'itmAppt.Organizer = exlSht.Cells(iRow, 12)
        'itmAppt.UserField1 = exlSht.Cells(iRow, 13)
        'itmAppt.UserField2 = exlSht.Cells(iRow, 14)
        
        itmAppt.AllDayEvent = False
        'itmAppt.Links.Add cnct
        
        itmAppt.ReminderSet = False
        itmAppt.ReminderMinutesBeforeStart = 30
        If itmAppt.Subject = exlSht.Cells(iRow - 1, 7) Then
            MsgBox itmAppt.Subject & " Duplicate Task!"
            GoTo ErrorHandler
        End If
    If countItems > 1024 Then
        MsgBox "Item.Count= " & countItems
        GoTo ErrorHandler
    End If
        itmAppt.Save
NextRow:
        iRow = iRow + 1
    Wend
    endTime = Timer
    Debug.Print Format(endTime - startTime, "##.00000")

    deleteAppointmentsModule.deleteTaskDeletedOnline
    MsgBox prompt:="Import Complete: " & iRow, _
    Buttons:=vbOKOnly + vbInformation, _
    Title:="Import Appointments"
      
ErrorHandlerExit:
   GoTo ExitSub

ErrorHandler:
   If Err.Number = 429 Then
      'Application object is not set by GetObject; use CreateObject instead
      If exlApp Is Nothing Then
         Set exlApp = CreateObject("Excel.Application")
         Resume Next
      End If
   Else
      MsgBox "Error No: " & Err.Number & "; Description: "
      Resume ErrorHandlerExit
   End If
ExitSub:
    With exlApp
        .Workbooks.Close
        .Quit
    End With
    Set exlApp = Nothing
End Sub

