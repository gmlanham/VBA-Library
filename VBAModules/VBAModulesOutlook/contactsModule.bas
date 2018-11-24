Attribute VB_Name = "contactsModule"
Option Explicit

'this code get the Contacts from Excel
Sub transferContacts()

OpenConnectionToExcel:
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim contactsFile As String
  Dim iRecord As Long
  contactsFile = relativePathModule.myPath & "\My Projects\Project.Scheduler\contacts.xls"
  Set db = OpenDatabase(Name:=contactsFile, Options:=False, ReadOnly:=False, Connect:="Excel 8.0")
  Set rs = db.OpenRecordset("SELECT * FROM `contactsTable` WHERE firstName <> null")
  Dim totalCount As Long
  rs.MoveLast
  totalCount = rs.recordCount
  rs.MoveFirst
  

CreateContactItems:
  On Error Resume Next
  Dim myNamespace As Outlook.NameSpace
  Dim myFolder As Outlook.Folder
  Dim myItem As Outlook.ContactItem
  Dim fullName As String

  Set myNamespace = Application.GetNamespace("MAPI")
  Set myFolder = myNamespace.GetDefaultFolder(olFolderContacts)
  Set myItem = myFolder.Items.Add

  While Not rs.EOF
    Set myItem = Outlook.CreateItem(olContactItem)
  
    fullName = rs.Fields(3).Value & ", " & rs.Fields(1).Value
    Call findName(fullName) 'check for duplicates
  
    myItem.FileAs = rs.Fields(3).Value & ", " & rs.Fields(1).Value
    myItem.LastName = rs.Fields(3).Value
    myItem.FirstName = rs.Fields(1).Value
    myItem.CompanyName = rs.Fields(5).Value
    myItem.BusinessAddress = rs.Fields(8).Value
    myItem.MobileTelephoneNumber = rs.Fields(40).Value
    myItem.BusinessTelephoneNumber = rs.Fields(31).Value
    myItem.HomeTelephoneNumber = rs.Fields(37).Value
    myItem.Email1Address = rs.Fields(57).Value
    myItem.Email1DisplayName = rs.Fields(59).Value
    iRecord = iRecord + 1
    If iRecord > totalCount Then GoTo ErrorHandler  'stop loop if count exceeds total records spreadsheet
    myItem.Save
    rs.MoveNext
  Wend
  MsgBox prompt:="Import Complete: " & iRecord, _
  Buttons:=vbOKOnly + vbInformation, _
  Title:="Import Contacts"
      
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

Function findName(ByVal fullName As String) As Boolean
'check to see if name is in Contacts, if found- delete it
  Dim myNamespace As Outlook.NameSpace
  Dim myFolder As Outlook.Folder
  Dim myItem As Outlook.ContactItem
  'Dim myContacts As Outlook.Items

  Set myNamespace = Application.GetNamespace("MAPI")
  Set myFolder = myNamespace.GetDefaultFolder(olFolderContacts)

'    Dim objOutlook As Outlook.Application
'    Dim objNamespace As Outlook.NameSpace
'    Dim objFolder As Outlook.MAPIFolder
'    Dim objAppointment As Outlook.AppointmentItem
'
'    Set objOutlook = Application
'    Set objNamespace = objOutlook.GetNamespace("MAPI")
'    Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)
    
        For Each myItem In myFolder.Items
            DoEvents
            On Error Resume Next
            If myItem.LastName & ", " & myItem.FirstName = fullName Then
                'delete task
                myItem.Delete
                Exit Function
            End If
        Next
ExitSub:
End Function

Sub contactsList()
    Dim myNamespace As Outlook.NameSpace
    Dim myContacts As Outlook.Items
    Dim myItems As Outlook.Items
    Dim myItem As Object
    
    Set myNamespace = Application.GetNamespace("MAPI")
    Set myContacts = myNamespace.GetDefaultFolder(olFolderContacts).Items
    
    Set myItems = myContacts
    For Each myItem In myItems
        If (myItem.Class = olContact) Then
            taskAppointmentForm.contactsComboBox.AddItem myItem.fullName
        End If
    Next
End Sub

'this macro goes through all contact items to show the ones created date input by user
Sub ContactDateCheck()
    Dim myNamespace As Outlook.NameSpace
    Dim myContacts As Outlook.Items
    Dim myItems As Outlook.Items
    Dim myItem As Object
    
    Set myNamespace = Application.GetNamespace("MAPI")
    Set myContacts = myNamespace.GetDefaultFolder(olFolderContacts).Items
    'ask user for date
    'Dim modificationDate As Date
    'modificationDate = InputBox("What last modification date that you want to see contacts from?")
    'modificationDate = #7/3/2011#
    
    Set myItems = myContacts   '.Restrict("[LastModificationTime] > """ & modificationDate & """")
    For Each myItem In myItems
        If (myItem.Class = olContact) Then
            'MsgBox myItem.FullName & ": " & myItem.LastModificationTime
            Debug.Print myItem.fullName & ": " & myItem.LastModificationTime
            categoriesListForm.contactsComboBox.AddItem myItem.fullName
        End If
    Next
End Sub

