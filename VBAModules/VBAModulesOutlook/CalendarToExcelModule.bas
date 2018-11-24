Attribute VB_Name = "CalendarToExcelModule"
Option Explicit

'Must declare as Object because folders may contain different
'types of items
Private itm As Object
Private strPrompt As String
Private strTitle As String

Sub showFTPForm()
'this replaces the transferCalendarToWeb process, much shorter!
'note that the relativePathModule.MyPath returns "C:\Documents and Settings\Mike Lanham\My Documents"
    Shell relativePathModule.myPath & _
    "\My Projects\Project.Scheduler" & "\FileTransferApplication.exe", vbNormalFocus
End Sub

Sub SaveContactsToExcel()
  'Created by Helen Feddema 9-17-2004
  'Last modified 20-Jul-2008
  'Demonstrates pushing Contacts data to rows in an Excel worksheet
  
  On Error GoTo ErrorHandler
  
  Dim appExcel As excel.Application
  'Dim appWord As Word.Application
  Dim blnMultiDay As Boolean
  Dim dteEnd As Date
  Dim dteStart As Date
  Dim fld As Outlook.MAPIFolder
  Dim I As Integer
  Dim intReturn As Integer
  Dim itms As Outlook.Items
  Dim J As Integer
  Dim lngCount As Long
  Dim nms As Outlook.NameSpace
  Dim ritms As Outlook.Items
  Dim rng As excel.Range
  Dim strDateRange As String
  Dim strEndDate As String
  'Dim strSheet As String
  Dim strSheetTitle As String
  Dim strStartDate As String
  'Dim strTemplatePath As String
  Dim wkb As excel.Workbook
  Dim wks As excel.Worksheet
  
  'Pick up Template path from the Word Options dialog
  'Set appWord = GetObject(, "Word.Application")
  'strTemplatePath = appWord.Options.DefaultFilePath(wdUserTemplatesPath) & "\"
  'Debug.Print "Templates folder: " & strTemplatePath
  Dim fullPath As String
  Dim folderPath As String
  Dim subfoldersPath As String
  Dim workbookName As String
  Dim myContactsPathName As String
  folderPath = relativePathModule.myPath    '   "C:\Documents and Settings\Mike Lanham\My Documents"
  subfoldersPath = "\My Projects\Project.Scheduler\"
  fullPath = folderPath & subfoldersPath
  workbookName = "contacts.xls"
  myContactsPathName = fullPath & workbookName
  
  'Test for file in the Templates folder
  If TestFileExists(myContactsPathName) = False Then
     strTitle = myContactsPathName & " file not found"
     strPrompt = myContactsPathName & _
        " not found; please copy Contacts.xls to Project.Scheduler folder and try again"
     MsgBox strPrompt, vbCritical + vbOKOnly, strTitle
     GoTo ErrorHandlerExit
  End If
   
   
SelectContactsFolder:
  'hardcode the Contacts rather than pick it.
  Set nms = Application.GetNamespace("MAPI")
  Set fld = nms.GetDefaultFolder(olFolderContacts)
  
  If fld Is Nothing Then
     MsgBox "Please select a Contacts folder"
     'Allow user to select Contacts folder
     Set fld = nms.PickFolder
  End If
  If fld.DefaultItemType <> olContactItem Then
     MsgBox "Please select a Contacts folder"
     GoTo SelectContactsFolder
  End If
   
getItems:
  Set itms = fld.Items
  Set ritms = itms.Restrict("[MessageClass]='IPM.Contact'")

  'Get an accurate count
  lngCount = 0
  For Each itm In ritms
     lngCount = lngCount + 1
  Next itm
CreateWorksheet:
  CreateObject("WScript.Shell").PopUp "The Export process can take a minute." & vbCr & _
  "Please wait until the 'Export Complete'" & vbCr & _
  "message, before clicking around.", _
      2, "Please Wait", 48
  Set appExcel = GetObject(, "Excel.Application")
  appExcel.Workbooks.Open (myContactsPathName)
  Set wkb = appExcel.ActiveWorkbook
  Set wks = wkb.Sheets(1)
  wks.Activate
  appExcel.Application.Visible = False

  'clear sheet of existing data
  'Excel.Application.Run "cleanupModule.clearSheet"

   'Adjust i (row number) to be 1 less than the number of the first body row
   I = 1
   
  'Iterate through contact items in Contacts folder, and export a few fields
  'from each item to a row in the Contacts worksheet
  For Each itm In ritms
    If I = lngCount Then Exit For
    I = I + 1
    'Process item only if it is an contacts item
    If itm.Class = olContact Then
      
      With wks
        .Cells(I, 2).Value = itm.Title
        .Cells(I, 3).Value = itm.FirstName
        .Cells(I, 5).Value = itm.LastName
        .Cells(I, 7).Value = itm.CompanyName
        .Cells(I, 9).Value = itm.JobTitle
        .Cells(I, 10).Value = itm.BusinessAddress
        .Cells(I, 13).Value = itm.BusinessAddressCity
        .Cells(I, 17).Value = itm.HomeAddress
        .Cells(I, 20).Value = itm.HomeAddressCity
        .Cells(I, 32).Value = itm.BusinessFaxNumber
        .Cells(I, 33).Value = itm.BusinessTelephoneNumber
        .Cells(I, 39).Value = itm.HomeTelephoneNumber
        .Cells(I, 42).Value = itm.MobileTelephoneNumber
        .Cells(I, 51).Value = itm.Anniversary
        .Cells(I, 54).Value = itm.Birthday
        .Cells(I, 59).Value = itm.Email1Address
        .Cells(I, 61).Value = itm.Email1DisplayName
      End With
    End If
  Next itm

'  Set rng = wks.Range("A1")
'  rng.value = strSheetTitle
  
  'sort sheet and trigger the File Tranfer Form to popup
  'Excel.Application.Run "cleanupModule.sortTable"
  
ErrorHandlerExit:
   GoTo ExitSub

ErrorHandler:
  If Err.Number = 429 Then
    If appExcel Is Nothing Then
      Set appExcel = CreateObject("Excel.Application")
      Resume Next
    End If
  Else
    MsgBox "Error No: " & Err.Number & "; Description: "
    Resume ErrorHandlerExit
  End If
ExitSub:
  With appExcel
    .DisplayAlerts = False
    .Save
    .Workbooks.Close
    .Quit
    .DisplayAlerts = True
  End With
  Set appExcel = Nothing
  MsgBox prompt:=lngCount & " Contacts Exported", _
  Buttons:=vbOKOnly + vbInformation, _
  Title:="Export Contacts"
End Sub
Sub SaveCalendarToExcel()
'Created by Helen Feddema 9-17-2004
'Last modified 20-Jul-2008
'Demonstrates pushing Calendar data to rows in an Excel worksheet

On Error GoTo ErrorHandler

   Dim appExcel As excel.Application
   'Dim appWord As Word.Application
   Dim blnMultiDay As Boolean
   Dim dteEnd As Date
   Dim dteStart As Date
   Dim fld As Outlook.MAPIFolder
   Dim I As Integer
   Dim intReturn As Integer
   Dim itms As Outlook.Items
   Dim J As Integer
   Dim lngCount As Long
   Dim nms As Outlook.NameSpace
   Dim ritms As Outlook.Items
   Dim rng As excel.Range
   Dim strDateRange As String
   Dim strEndDate As String
   'Dim strSheet As String
   Dim strSheetTitle As String
   Dim strStartDate As String
   'Dim strTemplatePath As String
   Dim wkb As excel.Workbook
   Dim wks As excel.Worksheet
   
   'Pick up Template path from the Word Options dialog
   'Set appWord = GetObject(, "Word.Application")
   'strTemplatePath = appWord.Options.DefaultFilePath(wdUserTemplatesPath) & "\"
   'Debug.Print "Templates folder: " & strTemplatePath
Dim fullPath As String
Dim folderPath As String
Dim subfoldersPath As String
Dim workbookName As String
Dim myCalendarPathName As String
   folderPath = relativePathModule.myPath    '   "C:\Documents and Settings\Mike Lanham\My Documents"
   subfoldersPath = "\My Projects\Project.Scheduler\"
   fullPath = folderPath & subfoldersPath
   workbookName = "calendar.xlsm"
   myCalendarPathName = fullPath & workbookName

   'Test for file in the Templates folder
   If TestFileExists(myCalendarPathName) = False Then
      strTitle = myCalendarPathName & " file not found"
      strPrompt = myCalendarPathName & _
         " not found; please copy Calendar.xls to Project.Scheduler folder and try again"
      MsgBox strPrompt, vbCritical + vbOKOnly, strTitle
      GoTo ErrorHandlerExit
   End If
   
   'Put up input boxes for start date and end date, and
   'create filter string
   strPrompt = "Please enter a start date for filtering appointments"
   strTitle = "Start Date"
   strStartDate = InputBox(strPrompt, strTitle)
   If strStartDate = "" Then
      'Don't use a date range
      strDateRange = ""
      strSheetTitle = "Calendar Items from Excel"
      GoTo SelectCalendarFolder
   Else
      If IsDate(strStartDate) = True Then
         dteStart = CDate(strStartDate)
         GoTo endDate
      Else
         GoTo CreateWorksheet
      End If
   End If
   
endDate:
   strPrompt = "Please enter an end date for filtering appointments"
   strTitle = "End Date"
   strEndDate = InputBox(strPrompt, strTitle)
   
   If IsDate(strEndDate) = True Then
      dteEnd = CDate(strEndDate)
      GoTo CreateFilter
   Else
      dteEnd = Date
   End If
   
CreateFilter:
   'Create date range string
   strStartDate = dteStart & " 12:00 AM"
   strEndDate = dteEnd & " 11:59 PM"
   strDateRange = "[Start] >= """ & strStartDate & _
      """ and [Start] <= """ & strEndDate & """"
   'Debug.Print strDateRange
   strSheetTitle = "Calendar Items from Excel for " _
      & Format(dteStart, "d-mmm-yyyy") & " to " _
      & Format(dteEnd, "d-mmm-yyyy")
   
SelectCalendarFolder:
   'hardcode the Calendar rather than pick it.
    Set nms = Application.GetNamespace("MAPI")
    Set fld = nms.GetDefaultFolder(9)
     
    If fld Is Nothing Then
        MsgBox "Please select a Calendar folder"
        'Allow user to select Calendar folder
        Set fld = nms.PickFolder
    End If
    If fld.DefaultItemType <> olAppointmentItem Then
        MsgBox "Please select a Calendar folder"
        GoTo SelectCalendarFolder
    End If
   
getItems:
   Set itms = fld.Items
   itms.IncludeRecurrences = True
   itms.Sort Property:="[Start]", Descending:=False
   'Debug.Print "Number of items: " & itms.Count
   
   If strDateRange <> "" Then
      Set ritms = itms.Restrict(strDateRange)
      ritms.Sort Property:="[Start]", Descending:=False
      
      'Get an accurate count
      'this is where all the filtered items are accessed
      'they could be read into an array here.
      lngCount = 0
      For Each itm In ritms
         'Debug.Print "Appt. subject: " & itm.Subject _
            & "; Start: " & itm.Start
         lngCount = lngCount + 1
      Next itm
   Else
      Set ritms = itms
   End If
   
   'Debug.Print "Number of restricted items: " & lngCount
   
   'Determine whether there are any multi-day events in the range
   blnMultiDay = False
   
   For Each itm In ritms
      If itm.AllDayEvent = True Then
         blnMultiDay = True
      End If
   Next itm
        
   If blnMultiDay = True Then
      'Ask whether to split all-day multi-day events
      strTitle = "Question"
      strPrompt = "Split all-day multi-day events into separate daily events" _
         & vbCrLf & "so they can be exported correctly?"
      intReturn = MsgBox(prompt:=strPrompt, _
         Buttons:=vbQuestion + vbYesNo, _
         Title:=strTitle)
      
      If intReturn = True Then
         Call SplitMultiDayEvents2(itms)
         'Reset ritms variable after splitting up multi-day events into
         'separate days
         Set ritms = ritms.Restrict(strDateRange)
         ritms.Sort Property:="[Start]", Descending:=False
      End If
   End If
      
CreateWorksheet:
    CreateObject("WScript.Shell").PopUp "The Export process can take a minute." & vbCr & _
    "Please wait until the 'Export Complete'" & vbCr & _
    "message, before clicking around.", _
        2, "Please Wait", 48
   Set appExcel = GetObject(, "Excel.Application")
   'Debug.Print myCalendarPathName
   appExcel.Workbooks.Open (myCalendarPathName)
   Set wkb = appExcel.ActiveWorkbook
   Set wks = wkb.Sheets(1)
   wks.Activate
   appExcel.Application.Visible = False

    'clear sheet of existing data
    excel.Application.Run "cleanupModule.clearSheet"

   'Adjust i (row number) to be 1 less than the number of the first body row
   I = 3
   
  'Iterate through contact items in Calendar folder, and export a few fields
   'from each item to a row in the Calendar worksheet
   For Each itm In ritms
      If itm.Class = olAppointment Then
         'Process item only if it is an appointment item
         I = I + 1
         
         'j is the column number
         J = 1
         
         Set rng = wks.Cells(I, J)
         If itm.GlobalAppointmentID <> "" Then rng.Value = itm.GlobalAppointmentID
         J = J + 1
         
         Set rng = wks.Cells(I, J)
         If itm.LastModificationTime <> "" Then rng.Value = itm.LastModificationTime
         J = J + 1
         
         Set rng = wks.Cells(I, J)
         If itm.CreationTime <> "" Then rng.Value = itm.CreationTime
         J = J + 1
         
         Set rng = wks.Cells(I, J)
         If itm.Start <> "" Then rng.Value = itm.Start
         J = J + 1
         
         Set rng = wks.Cells(I, J)
         If itm.End <> "" Then rng.Value = itm.End
         J = J + 1
        
         Set rng = wks.Cells(I, J)
         If itm.Duration <> "" Then rng.Value = itm.Duration
         J = J + 1
                 
         Set rng = wks.Cells(I, J)
         If itm.Subject <> "" Then rng.Value = itm.Subject
         J = J + 1
         
         Set rng = wks.Cells(I, J)
         If itm.Location <> "" Then rng.Value = itm.Location
         J = J + 1
         
         Set rng = wks.Cells(I, J)
         If itm.Categories <> "" Then rng.Value = itm.Categories
         J = J + 1
        
         Set rng = wks.Cells(I, J)
         If itm.Body <> "" Then rng.Value = itm.Body
         J = J + 1
                 
         Set rng = wks.Cells(I, J)
         If itm.RequiredAttendees <> "" Then rng.Value = itm.RequiredAttendees
         J = J + 1
         
         'Set rng = wks.Cells(I, J)
         'If itm.Originator <> "" Then rng.Value = itm.Originator
         'J = J + 1
         
         Set rng = wks.Cells(I, J)
         On Error Resume Next
         'The next line illustrates the syntax for referencing
         'a custom Outlook field
         If itm.UserProperties("CustomField") <> "" Then
            rng.Value = itm.UserProperties("CustomField")
         End If
         J = J + 1
      End If
      I = I + 1
   Next itm

   Set rng = wks.Range("A1")
   rng.Value = strSheetTitle
   
    'sort sheet and trigger the File Tranfer Form to popup
    excel.Application.Run "cleanupModule.sortTable"
     
ErrorHandlerExit:
   GoTo ExitSub

ErrorHandler:
   If Err.Number = 429 Then
      'Application object is not set by GetObject; use CreateObject instead
      'If appWord Is Nothing Then
     '    Set appWord = CreateObject("Word.Application")
     '    Resume Next
      If appExcel Is Nothing Then
         Set appExcel = CreateObject("Excel.Application")
         Resume Next
      End If
   Else
      MsgBox "Error No: " & Err.Number & "; Description: "
      Resume ErrorHandlerExit
   End If
ExitSub:
    With appExcel
        .Close SaveChanges:=True
        .Quit
    End With
    Set appExcel = Nothing
End Sub

Public Function TestFileExists(strFile As String) As Boolean
'Created by Helen Feddema 9-1-2004
'Last modified 9-1-2004
'Tests for existing of a file, using the FileSystemObject
   
   Dim fso As New Scripting.FileSystemObject
   Dim fil As Scripting.File
   
On Error Resume Next

   Set fil = fso.GetFile(strFile)
   If fil Is Nothing Then
      TestFileExists = False
   Else
      TestFileExists = True
   End If
   
End Function

Public Sub SplitMultiDayEvents2(itmsSet As Outlook.Items)
'Created by Helen Feddema 7-Jun-2007
'Last modified 20-Jul-2008

On Error GoTo ErrorHandler

   Dim dteNewEnd As Date
   Dim itmCopy As Outlook.AppointmentItem
   Dim lngDayCount As Long
   Dim N As Integer
   Dim strApptStart As String
   Dim strApptEnd As String
   Dim strApptRange As String
   Dim strApptSubject As String
   Dim strApptLocation As String
   Dim strApptNotes As String
   Dim strNewDate As String
   
   For Each itm In itmsSet
      strApptStart = Format(itm.Start, "h:mma/p")
      strApptEnd = Format(itm.End, "h:mma/p")
      strApptRange = strApptStart & " - " & strApptEnd & ":"
      strApptSubject = itm.Subject
      strApptLocation = itm.Location
      strApptNotes = itm.Body
      
      If itm.AllDayEvent = True Then
         'Debug.Print "All-day appt. range: " & itm.Start & " to " & itm.End
         
         'Check for multi-day all-day events, and make a separate all-day event
         'for each day in the date range if found
         lngDayCount = dateDiff("d", itm.Start, itm.End)
         If lngDayCount > 1 Then
            'This is a multi-day event; change original event to a single-day event
            dteNewEnd = DateAdd("d", 1, itm.Start)
            strNewDate = dteNewEnd & " 12:00 AM"
            itm.End = strNewDate
            itm.Close (olSave)
            
            'Make copies of this event for the other days in the range
            For N = 1 To lngDayCount - 1
               Set itmCopy = itm.Copy
               itmCopy.Subject = strApptSubject
               itmCopy.Location = strApptLocation
               itmCopy.AllDayEvent = True
               itmCopy.Body = strApptNotes
               itmCopy.Start = strNewDate
               dteNewEnd = DateAdd("d", 1, dteNewEnd)
               strNewDate = dteNewEnd & " 12:00 AM"
               itmCopy.End = strNewDate
               'itmCopy.Display
               itmCopy.Close (olSave)
            Next N
         End If
      End If
   Next itm
        
   strTitle = "Done"
   strPrompt = "Multi-day events split"
   MsgBox prompt:=strPrompt, _
      Buttons:=vbInformation + vbOKOnly, _
      Title:=strTitle
           
ErrorHandlerExit:
   Exit Sub

ErrorHandler:
   MsgBox "Error No: " & Err.Number & "; Description: " & _
      Err.Description
   Resume ErrorHandlerExit

End Sub


Sub transferCalendarToWeb()
'Created by Helen Feddema 9-17-2004
'open the Calendar Excel workbook
'then upload to website

On Error GoTo ErrorHandler

   Dim appExcel As excel.Application
   Dim strSheet As String
   Dim strTemplatePath As String
   Dim wkb As excel.Workbook
   Dim wks As excel.Worksheet
'
    'My Path is  "C:\Documents and Settings\Mike Lanham\My Documents"
    'the relavie path code was lifted from the web. It uses the shell32 library.
    Dim folderPath As String
    folderPath = relativePathModule.myPath
   
   strTemplatePath = folderPath & "\My Projects\Project.Scheduler\"
   strSheet = "calendar.xlsm"
   strSheet = strTemplatePath & strSheet
'
   'Test for file
   If TestFileExists(strSheet) = False Then
      strTitle = "Worksheet file not found"
      strPrompt = strSheet & _
         " not found; please copy Calendar.xls to this folder and try again"
      MsgBox strPrompt, vbCritical + vbOKOnly, strTitle
      GoTo ErrorHandlerExit
   End If

CreateWorksheet:
   Set appExcel = GetObject(, "Excel.Application")
   appExcel.Workbooks.Open (strSheet)
   Set wkb = appExcel.ActiveWorkbook
   Set wks = wkb.Sheets(1)
   wks.Activate
   appExcel.Application.Visible = False

    'FTP Data
    excel.Application.Run "ftpModule.ftpData3"
    excel.Application.Quit
    
ErrorHandlerExit:
   Exit Sub

ErrorHandler:
   If Err.Number = 429 Then
      If appExcel Is Nothing Then
         Set appExcel = CreateObject("Excel.Application")
         Resume Next
      End If
   Else
      MsgBox prompt:="Error No: " & Err.Number & "; Description: ", Buttons:=vbOKOnly + vbCritical
      Resume ErrorHandlerExit
   End If

End Sub



