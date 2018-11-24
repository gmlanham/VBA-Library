Attribute VB_Name = "TakeoffDataGetModule"
Option Explicit
Public divisionTitle As String
Dim TakeoffApp As excel.Application
Dim TakeoffWorkbook As excel.Workbook
Dim ExcelWasNotRunning As Boolean
Dim WorkbookPath As String
Dim updatedTakeoff As Worksheet
'Public flagCount As Boolean
'Public rowCount
'Public dataArray()

Sub OpenExcel() '(Optional ByVal showTakeoffCP As Boolean)
  On Error GoTo ErrorHandler:
  Const procedureName = "OpenExcel"
  Dim startTimer As Single
  Dim endTimer As Single
  Dim statusbarString As String
  StatusBar = procedureName
  'StartTime = Timer
  startTimer = Timer - StartTime
'open Takeoff
  On Error Resume Next
  'closeExcel
  Set TakeoffApp = GetObject(, "Excel.Application")
  If Err Then
    Set TakeoffApp = New excel.Application
    ExcelWasNotRunning = True
  End If
  'WorkbookPath = ("C:/Users/Mike/Documents/My Projects/Project.QT/TakeoffUtility4.xlsm")
  WorkbookPath = (ThisDocument.Path & "\TakeoffUtility4.xlsm")
  'WorkbookPath = (ThisDocument.Path & "/" & "takeoffTables.xls")

'open the excel workbook to Word's eyes
  Set TakeoffWorkbook = TakeoffApp.Workbooks.Open(WorkbookPath)
  TakeoffApp.Visible = controlPanel.screenUpdatesCheckBox.Value 'this where we make Takeoff visible or not, much faster when false

'  Dim TakeoffCP As UserForm
'  Set TakeoffCP = excel.Application.TakeoffCP
'  If showTakeoffCP Then TakeoffCP.Show
  
Cleanup:
  endTimer = Timer - StartTime
  statusbarString = procedureName & "... " & vbTab & _
    "Section Timer= " & Format(endTimer - startTimer, "#.0") & " seconds"
  Application.StatusBar = statusbarString
  Debug.Print statusbarString
  startTimer = endTimer
 Exit Sub
ErrorHandler:
  closeExcel
  Debug.Print "An error was thrown by " & procedureName & _
  vbCr & Err.Number & ": " & Err.Description
  Resume Cleanup:
End Sub

Sub closeExcel()
  Const procedureName = "closeExcel"
  Debug.Print procedureName
'Clean-up and close
On Error GoTo ExitSub:
    TakeoffWorkbook.Close SaveChanges:=False
    
    'release all variables, set objects to nothing, strings to vbnullstring
    Set TakeoffApp = Nothing
    Set TakeoffWorkbook = Nothing
    Set updatedTakeoff = Nothing
    WorkbookPath = vbNullString
    excel.Application.Quit
    'the sendkeys closes the VBE window, if this macro is called from the VBE window
    'this command close the calling app, i.e. this Word document, not Excel
    'SendKeys "%{F4}", True
    
ExitSub:
End Sub
'killExcel replicated dozens of instances causing computer hangup.
Sub killExcel()
  Dim sKill As String
  sKill = "TASKKILL /F /IM excel.exe"
  Shell sKill, vbHide
End Sub

'this code opens the Takeoff spreadsheet to get the section data
Function takeoffDataArray(ByVal sectionName As String) As Variant
    ReDim transferArray(60, 6)  'this is necessary to prevent error until the array set in arrayTest
    'openExcel
    takeoffDataArray = takeoffData(sectionName)
    'closeExcel
End Function
Private Sub testTakeoffData()
  Dim startTimer As Single
  Dim stopTimer As Single
  startTimer = Timer
  Dim transferArray() As Variant
  Call OpenExcel
  transferArray = takeoffData("Walls")
  Debug.Print UBound(transferArray, 1), UBound(transferArray, 2)
  Debug.Print transferArray(0, 0), transferArray(1, 0) ', transferArray(2, 0) ', transferArray(3, 0)
  Debug.Print transferArray(0, 1), transferArray(1, 1) ', transferArray(2, 1) ', transferArray(3, 1)
  Debug.Print transferArray(0, 2), transferArray(1, 2) ', transferArray(2, 2) ', transferArray(3, 2)
  stopTimer = Timer
  Debug.Print "Run Time: " & Format(stopTimer - startTimer, "#.000")
exitHandler:
End Sub


Function takeoffData(ByVal sectionName As String) As Variant
    Dim I As Long
    Dim J As Long
    Dim transferArray()
    Dim rowCount As Long
    Dim lastColumn As Long
    Dim startRange As String
    Dim endRange As String
    Dim rangeAddress As String
    
    lastColumn = 6
    sectionName = "selected" & sectionName
    
    'the range with Columns() seems to be unstable, so get the range address as a string first then count
    rangeAddress = Range(sectionName).Columns(2).Address
    rowCount = WorksheetFunction.CountA(Range(rangeAddress))
    If InStr(sectionName, "extras") <> 0 Then rowCount = 61
    'If rowCount = 0 Then Exit Function
    startRange = Range(sectionName).Rows(1).Address
    If rowCount = 0 Then rowCount = 1
    endRange = Range(sectionName).Rows(rowCount).Address
    
    ReDim transferArray(1 To rowCount, 1 To lastColumn)
    transferArray = Range(startRange, endRange)
    divisionTitle = Range(sectionName).Cells(1, 1).Offset(-2, 1).Value
    
    takeoffData = transferArray
End Function
Sub ADORecordset()
    On Error GoTo AdoError
    'using client side cursor
    Dim startTimer As Single
    Dim stopTimer As Single
    startTimer = Timer
    ' Error Handling Variables
    Dim errLoop As Error
    Dim strTmp As String
    Dim Errs1 As Errors
      
    Dim newPath As String
    Dim rsWalls As ADODB.Recordset
    Dim rsOther As ADODB.Recordset
    Dim rsOPI As ADODB.Recordset
    Dim rsMoreOPI As ADODB.Recordset
    Dim rsSummaryOPI As ADODB.Recordset
    Dim rsExcavation As ADODB.Recordset
    Dim rsWater As ADODB.Recordset
    Dim rsFoundation As ADODB.Recordset
    Dim rsSeasonal As ADODB.Recordset
    Dim rsPegasusExcavation As ADODB.Recordset
    Dim rsSpreadEagle As ADODB.Recordset
    Dim rsMaterials As ADODB.Recordset
    Dim rsWaterPegasus As ADODB.Recordset
    Dim rsNewDivision As ADODB.Recordset
    Dim rsClientInfo As ADODB.Recordset
    'Dim calendarData As String
    Dim I As Long
    Dim totalCount As Long
    Dim cn As New ADODB.Connection

connectToExcel:
    If ComputerName = "MIKE2" Then
      newPath = ThisDocument.Path
    Else
      newPath = "M:\Estimating and Invoicing\Estimating and Invoicing 2012\"
    End If
    With cn
      .Provider = "MSDASQL"
      .ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & _
        "DBQ=" & newPath & "\TakeoffTables.xls; "
'      .Provider = "Microsoft.Jet.OLEDB.4.0"
'      .ConnectionString = "Data Source=" & ThisDocument.Path & _
'     "\TakeoffTables.xls;Extended Properties=Excel 8.0;"
      .CursorLocation = CursorLocationEnum.adUseClient
      .Open
    End With

    'calendarData = newPath & "TakeoffTables.xls"
    'only Excel 8.0 files with .xls extensions connect
    'the indexes are Zero based for records and for fields
    'need column headers in the named range
    On Error Resume Next
    Set rsWalls = cn.Execute("SELECT * FROM `selectedWalls` WHERE Count <> null")
    Set rsOther = cn.Execute("SELECT * FROM `selectedOther` WHERE Count <> null")
    Set rsOPI = cn.Execute("SELECT * FROM `selectedOPI` WHERE Count <> null")
    Set rsMoreOPI = cn.Execute("SELECT * FROM `selectedMoreOPI` WHERE Count <> null")
    Set rsSummaryOPI = cn.Execute("SELECT * FROM `selectedSummaryOPI` WHERE Count <> null")
    Set rsExcavation = cn.Execute("SELECT * FROM `selectedExcavagtion` WHERE Count <> null")
    Set rsWater = cn.Execute("SELECT * FROM `selectedWater` WHERE Count <> null")
    Set rsFoundation = cn.Execute("SELECT * FROM `selectedFoundation` WHERE Count <> null")
    Set rsSeasonal = cn.Execute("SELECT * FROM `selectedSeasonal` WHERE Count <> null")
    Set rsPegasusExcavation = cn.Execute("SELECT * FROM `selectedPegasusExcavation` WHERE Count <> null")
    Set rsSpreadEagle = cn.Execute("SELECT * FROM `selectedSpreadEagle` WHERE Count <> null")
    Set rsMaterials = cn.Execute("SELECT * FROM `selectedMaterials` WHERE Count <> null")
    Set rsWaterPegasus = cn.Execute("SELECT * FROM `selectedWaterPegasus` WHERE Count <> null")
    Set rsNewDivision = cn.Execute("SELECT * FROM `selectedNewDivision` WHERE Count <> null")
    Set rsClientInfo = cn.Execute("SELECT * FROM `ClientTable` WHERE [Client Name] <> null")

    StartForm.List1.AddItem ("Tasks: ClientInfo")
    StartForm.List1.AddItem ("--------------------")
    With rsClientInfo
      While Not (.BOF Or .EOF)
        StartForm.List1.AddItem (.Fields(0).Value & ", " & .Fields(1).Value & ", " & .Fields(2).Value & vbCr)
        StartForm.List1.AddItem (.Fields(3).Value & ", " & .Fields(4).Value & ", " & .Fields(5).Value & vbCr)
        StartForm.List1.AddItem (.Fields(6).Value & ", " & .Fields(7).Value & ", " & .Fields(8).Value & vbCr)
        StartForm.List1.AddItem (.Fields(9).Value & ", " & .Fields(10).Value & ", " & .Fields(11).Value & ", " & .Fields(12).Value & vbCr)
        .MoveNext
      Wend
    End With
    StartForm.List1.AddItem (vbTab & vbTab & "--------------------")
    StartForm.List1.AddItem ("Tasks: Walls")
    StartForm.List1.AddItem ("--------------------")
    With rsWalls
      While Not (.BOF Or .EOF)
        StartForm.List1.AddItem (.Fields("count").Value & " " & .Fields("description").Value & " " & .Fields("rate").Value & vbCr)
        .MoveNext
      Wend
    End With
    StartForm.List1.AddItem (vbTab & vbTab & "--------------------")
    StartForm.List1.AddItem ("Tasks: Other")
    StartForm.List1.AddItem ("--------------------")
    With rsOther
      While Not (.BOF Or .EOF)
        StartForm.List1.AddItem (.Fields("count").Value & " " & .Fields("description").Value & " " & .Fields("rate").Value & vbCr)
        .MoveNext
      Wend
    End With
    StartForm.List1.AddItem (vbTab & vbTab & "--------------------")

    rsClientInfo.Close
    rsWalls.Close
    rsOther.Close
    cn.Close
    StartForm.Show vbModeless
    stopTimer = Timer
    MsgBox ("Run Time: " & Format(stopTimer - startTimer, "#.0000") & " seconds")
Cleanup:
    Set rsClientInfo = Nothing
    Set rsWalls = Nothing
    Set rsOther = Nothing

    Set cn = Nothing

    Exit Sub
AdoError:
       I = 1
       On Error Resume Next

       ' Enumerate Errors collection and display properties of
       ' each Error object (if Errors Collection is filled out)
'       Set Errs1 = cn.Errors
'       For Each errLoop In Errs1
'        With errLoop
'           strTmp = strTmp & vbCrLf & "ADO Error # " & I & ":"
'           strTmp = strTmp & vbCrLf & "   ADO Error   # " & .Number
'           strTmp = strTmp & vbCrLf & "   Description   " & .Description
'           strTmp = strTmp & vbCrLf & "   Source        " & .Source
'           I = I + 1
'        End With
'       Next

AdoErrorLite:
       ' Get VB Error Object's information
       strTmp = strTmp & vbCrLf & "VB Error # " & Str(Err.Number)
       strTmp = strTmp & vbCrLf & "   Generated by " & Err.Source
       strTmp = strTmp & vbCrLf & "   Description  " & Err.Description

       MsgBox strTmp

       ' Clean up gracefully without risking infinite loop in error handler
       On Error GoTo 0
       GoTo Cleanup:
   End Sub
Private Sub testRSArray()
  Dim startTimer As Single
  Dim stopTimer As Single
  startTimer = Timer

'variant arrays are necessary to pass parameters to a function that get the recordset
'a second array, the transferArray, is populate by the values returned by the function that gets the recordset
'then array indexes specify the elements to print/use
  Dim sectionName As String
  Dim transferArray() As Variant
  sectionName = "summaryopi"
  transferArray = recordsetArray(sectionName)
  Dim I As Long
  Debug.Print UBound(transferArray, 1), UBound(transferArray, 2)
  Debug.Print transferArray(0, 0), transferArray(1, 0), transferArray(2, 0), transferArray(3, 0)
  Debug.Print transferArray(0, 1), transferArray(1, 1), transferArray(2, 1), transferArray(3, 1)
  Debug.Print transferArray(0, 2), transferArray(1, 2), transferArray(2, 2), transferArray(3, 2)
  stopTimer = Timer
  Debug.Print "Run Time: " & Format(stopTimer - startTimer, "#.000")
exitHandler:
End Sub

Function recordsetArray(ByVal sectionName As String) As Variant
'  Dim startTimer As Single
'  Dim stopTimer As Single
'  startTimer = Timer

  Dim cn As New ADODB.Connection
  Dim rs As ADODB.Recordset
  Dim arrRecordArray() As Variant
  Dim newPath As String
  If ComputerName = "MIKE2" Then
    newPath = ThisDocument.Path
  Else
    newPath = "M:\Estimating and Invoicing\Estimating and Invoicing 2012\"
  End If
  With cn
    .Provider = "MSDASQL"
    .ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & _
      "DBQ=" & newPath & "\TakeoffTables.xls; "
'      .Provider = "Microsoft.Jet.OLEDB.4.0"
'      .ConnectionString = "Data Source=" & ThisDocument.Path & _
'     "\TakeoffTables.xls;Extended Properties=Excel 8.0;"
    .CursorLocation = CursorLocationEnum.adUseClient
    .Open
  End With
    sectionName = "selected" & sectionName
    On Error Resume Next
    Set rs = cn.Execute("SELECT * FROM " & sectionName & " WHERE Count <> null")

'    If rs.EOF = True And rs.BOF = True Then
'      GoTo exitHandler
'    End If
    
  'Get this to Array.
  arrRecordArray = rs.GetRows
  recordsetArray = arrRecordArray
'  stopTimer = Timer
'  Debug.Print Format(stopTimer - startTimer, "#.000")
exitHandler:
End Function
  Sub getTakeoffData2()
    Dim I As Long
    Dim Sections(9) As String
    Sections(1) = "walls"
    Sections(2) = "other"
    Sections(3) = "opi"
    Sections(4) = "moreopi"
    Sections(5) = "summaryopi"
    Sections(6) = "excavation"
    Sections(7) = "water"
    Sections(8) = "foundation"
    Sections(9) = "seasonal"
    ReDim transferArray(14, 6)  'this is necessary to prevent error until the array set in arrayTest
    'openExcel
    For I = 1 To 9
        takeoffData Sections(I)
    Next I
    'closeExcel
End Sub
