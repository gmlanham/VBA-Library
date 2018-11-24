Attribute VB_Name = "DAOModule"
Option Explicit

Sub transferContacts()

OpenConnectionToExcel:
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim dataSourceFile As String
  dataSourceFile = "C:\Users\mike\Documents\My Projects\Project.Scheduler\contacts.xls"
  Set db = OpenDatabase(Name:=dataSourceFile, Options:=False, ReadOnly:=False, Connect:="Excel 8.0")
  Dim dataTable As String
  Dim dataField As String
  Dim dataQuery As String
  dataTable = "Contactstable"
  dataField = "firstName"
  dataQuery = _
    "SELECT * FROM " & [dataTable] & _
    " WHERE " & [dataField] & " <> null"
  
  Set rs = db.OpenRecordset(Name:=dataQuery)
  Dim totalCount As Long
  rs.MoveLast
  totalCount = rs.RecordCount
  rs.MoveFirst
  
  Debug.Print rs.RecordCount + 1
  rs.Close
  db.Close
  
  Set rs = Nothing
  Set db = Nothing
End Sub
Private Sub testGetData()
  Dim contactsFile As String
  contactsFile = "C:\Users\mike\Documents\My Projects\Project.Scheduler\contacts.xls"
  getExcelData contactsFile
End Sub
Public Sub getExcelData(ByVal FileName As String)
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Set db = OpenDatabase(Name:=FileName, Options:=False, ReadOnly:=False, Connect:="Excel 8.0")
  Set rs = db.OpenRecordset("SELECT * FROM `contactsTable` WHERE firstName <> null")
  
  While Not rs.EOF
    Debug.Print rs.RecordCount, rs.Fields(3).Value
    rs.MoveNext
  Wend
  Debug.Print rs.RecordCount + 1
  rs.Close
  db.Close
  
  Set rs = Nothing
  Set db = Nothing
End Sub

Private Sub test()
'  Dim copyData As String
'  copyData = "C:\Users\mike\Documents\My Projects\Project.Scheduler\contactsCopy.xls"
  Dim calendarData As String
  calendarData = "C:\Users\mike\Documents\My Projects\Project.Scheduler\contacts.xls"
  Dim tableName As String
  tableName = "Contacts"
  Dim fieldName As String
  fieldName = "FirstName"
  Dim targetRange As Range
  Set targetRange = Workbooks(calendarData) '.Worksheets(tableName).Range("C1")
  DAOCopyFromRecordSet DBFullName:=calendarData, _
    tableName:=tableName, _
    fieldName:=fieldName, _
    targetRange:=Worksheets("Sheet1").Range("C1")
End Sub
Sub DAOCopyFromRecordSet(DBFullName As String, tableName As String, _
    fieldName As String, targetRange As Range)
' Example: DAOCopyFromRecordSet "C:\FolderName\DataBaseName.mdb", _
    "TableName", "FieldName", Range("C1")
    
Dim db As Database, rs As Recordset
Dim intColIndex As Integer
    Set targetRange = targetRange.Cells(1, 1)
    Set db = OpenDatabase(DBFullName)
    Set rs = db.OpenRecordset(tableName, dbOpenTable) ' all records
    'Set rs = db.OpenRecordset("SELECT * FROM " & TableName & _
        " WHERE " & FieldName & _
        " = 'MyCriteria'", dbReadOnly) ' filter records
    ' write field names
    For intColIndex = 0 To rs.Fields.Count - 1
        targetRange.Offset(0, intColIndex).Value = rs.Fields(intColIndex).Name
    Next
    ' write recordset
    targetRange.Offset(1, 0).CopyFromRecordset rs
    Set rs = Nothing
    db.Close
    Set db = Nothing
End Sub

