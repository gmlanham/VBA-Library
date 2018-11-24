Attribute VB_Name = "Module1"
Option Explicit


Sub transferContacts()

OpenConnectionToExcel:
  Dim db As DAO.dataBase
  Dim Rs As DAO.Recordset
  Dim dataSourceFile As String
  Dim filePath As String
  filePath = "C:\Users\mike\Documents\My Projects\MacroLibrary\"
  dataSourceFile = filePath & "calendar.xls"
  Set db = OpenDatabase(Name:=dataSourceFile, Options:=False, ReadOnly:=False, Connect:="Excel 8.0")
  Dim dataTable As String
  Dim dataField As String
  Dim dataQuery As String
  
  dataTable = "jobSchedule"
  dataField = "Subject"
  dataQuery = _
    "SELECT * FROM " & [dataTable] _
    & " WHERE " & [dataField] & " <> null"
  
  Set Rs = db.OpenRecordset(Name:=dataQuery)
  Dim totalCount As Long
  Rs.MoveLast
  totalCount = Rs.RecordCount
  Rs.MoveFirst
  While Not Rs.EOF
    Debug.Print Rs.Fields(6).Value
    Rs.MoveNext
  Wend
  Debug.Print Rs.RecordCount + 1
  Rs.Close
  db.Close
  
  Set Rs = Nothing
  Set db = Nothing
End Sub

Private Sub testTableInfo()
  Dim dataSourceFile As Workbook
  Dim filePath As String
  filePath = "C:\Users\mike\Documents\My Projects\MacroLibrary\"
  dataSourceFile = filePath & "calendar.xls"

Dim Wb As Workbook
Set Wb = dataSourceFile
Dim Ws As Worksheet
Set Ws = Wb.Worksheets("Sheet1")

Dim dataTable As String
'dataTable = "jobSchedule"
dataTable = Ws.Name
  TableInfo (dataTable)
End Sub
Function TableInfo(strTableName As String)
On Error GoTo TableInfoErr
   ' Purpose:   Display the field names, types, sizes and descriptions for a table.
   ' Argument:  Name of a table in the current database.
   Dim db As DAO.dataBase
   Dim tdf As DAO.TableDef
   Dim fld As DAO.Field
    Dim dataSourceFile As String
  Dim filePath As String
  filePath = "C:\Users\mike\Documents\My Projects\MacroLibrary\"
  dataSourceFile = filePath & "calendar.xls"

  Set db = OpenDatabase(Name:=dataSourceFile, Options:=False, ReadOnly:=False, Connect:="Excel 8.0")
   Set tdf = db.TableDefs(strTableName)
   Debug.Print "FIELD NAME", "FIELD TYPE", "SIZE", "DESCRIPTION"
   Debug.Print "==========", "==========", "====", "==========="

   For Each fld In tdf.Fields
      Debug.Print fld.Name,
      Debug.Print FieldTypeName(fld),
      Debug.Print fld.Size,
      Debug.Print GetDescrip(fld)
   Next
   Debug.Print "==========", "==========", "====", "==========="

TableInfoExit:
   Set db = Nothing
   Exit Function

TableInfoErr:
   Select Case Err
   Case 3265&  'Table name invalid
      MsgBox strTableName & " table doesn't exist"
   Case Else
      Debug.Print "TableInfo() Error " & Err & ": " & Error
   End Select
   Resume TableInfoExit
End Function


Function GetDescrip(obj As Object) As String
    On Error Resume Next
    GetDescrip = obj.Properties("Description")
End Function


Function FieldTypeName(fld As DAO.Field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn
End Function

'Private Sub ADOTest()
' Dim cn As ADODB.Connection
' Set cn = New ADODB.Connection
'   Dim dataSourceFile As String
'  Dim filePath As String
'  filePath = "C:\Users\mike\Documents\My Projects\MacroLibrary\"
'  dataSourceFile = filePath & "calendar.xls"
'  Dim dataTable As String
'  Dim dataField As String
'  Dim dataQuery As String
'
'  dataTable = "jobSchedule"
'  dataField = "Subject"
'  dataQuery = _
'    "SELECT * FROM  " & [dataTable] & _
'    " WHERE " & [dataField] & " <> null"
' With cn
' .Provider = "MSDASQL"
' .ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & _
' "DBQ=datasourcefile; ReadOnly=False;"
' .Open
' End With
' Dim Rs As ADODB.Recordset
' Set Rs = New ADODB.Recordset
'
'Rs.Open dataQuery, cn, adOpenStatic, , adCmdText
' Rs.Close
' Set Rs = Nothing
' cn.Close
' Set cn = Nothing
' End Sub

