VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim closingFlag As Boolean

Private Sub Workbook_Open()
  On Error Resume Next
  closingFlag = False
  With Application
      .DisplayAlerts = False
      .MacroOptions Macro:="showControlPanel", HasShortcutKey:=True, ShortcutKey:="p"
      .WindowState = xlMaximized
  End With
  Range("A1").Select
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  With Application
    .DisplayAlerts = False
    .ScreenUpdating = False
    Call DeleteFromCellMenu
  
    ThisWorkbook.Save
    If closingFlag = False Then
      CopyDataToNewWorkbook
      closingFlag = True
    End If
    .ScreenUpdating = True
    .DisplayAlerts = True
  End With
End Sub

Sub CopyDataToNewWorkbook()
  Application.DisplayAlerts = False
  Dim startTimer As Single
  Dim stopTimer As Single
  startTimer = Timer

    Cells.Select
    Selection.Copy
    Workbooks.Add
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    'ChDir "C:\Users\mike\Documents\My Projects\Project.QT"
    
    'FileFormat:=56 for xlExcel8, 97-2003 format w/o macros
    ActiveWorkbook.SaveAs FileName:= _
        ThisWorkbook.Path & "\TakeoffTables.xls", FileFormat _
        :=56, Password:="", WriteResPassword:="", ReadOnlyRecommended:= _
        False, CreateBackup:=False
        
'Sub setTableRanges()
'  Windows("TakeoffTables.xls").Activate
  'include column headings with range for DAO connection
  Dim rangesTable(1 To 15) As String
  Dim I As Long
  rangesTable(1) = "selectedExcavation"
  rangesTable(2) = "selectedWater"
  rangesTable(3) = "selectedFoundation"
  rangesTable(4) = "selectedSeasonal"
  rangesTable(5) = "selectedOPI"
  rangesTable(6) = "selectedMoreOPI"
  rangesTable(7) = "selectedSummaryOPI"
  rangesTable(8) = "selectedWalls"
  rangesTable(9) = "selectedOther"
  rangesTable(10) = "selectedMaterials"
  rangesTable(11) = "selectedExcavationPegasus"
  rangesTable(12) = "selectedWaterPegasus"
  rangesTable(13) = "selectedSpreadEagle"
  rangesTable(14) = "selectedNewDivision"
  rangesTable(15) = "ClientTable"

  With ActiveWorkbook.Sheets(1)
  .Range("G68:K82").Name = "selectedExcavation"
  .Range("G84:K98").Name = "selectedWater"
  .Range("G100:K114").Name = "selectedFoundation"
  .Range("G116:K130").Name = "selectedSeasonal"
  .Range("G132:K146").Name = "selectedOPI"
  .Range("G148:K162").Name = "selectedMoreOPI"
  .Range("G164:K178").Name = "selectedSummaryOPI"
  .Range("G180:K194").Name = "selectedWalls"
  .Range("G196:K210").Name = "selectedOther"
  .Range("G229:K243").Name = "selectedMaterials"
  .Range("G246:K260").Name = "selectedExcavationPegasus"
  .Range("G263:K277").Name = "selectedWaterPegasus"
  .Range("G280:K294").Name = "selectedSpreadEagle"
  .Range("G297:K311").Name = "selectedNewDivision"
  .Range("T1:AD10").Name = "ClientTable"
  
  For I = 1 To 14
    .Range(rangesTable(I)).Interior.Color = 10092543
    .Range(rangesTable(I)).Cells(1, 1).Offset(-1, 0).Value = rangesTable(I)
    .Range(rangesTable(I)).Cells(1, 1).Value = "Count"
    .Range(rangesTable(I)).Cells(1, 2).Value = "Description"
    .Range(rangesTable(I)).Cells(1, 3).Value = "Rate"
    .Range(rangesTable(I)).Cells(1, 4).Value = "Units"
    .Range(rangesTable(I)).Cells(1, 5).Value = "Price"
  Next I
  'ClientTable
    .Range(rangesTable(I)).Interior.Color = 10092543
    '.Range(rangesTable(I)).Cells(1, 1).Offset(-1, 0).Value = rangesTable(I)
    .Range(rangesTable(I)).Cells(1, 1).Value = "Takeoff Filename"
    .Range(rangesTable(I)).Cells(1, 2).Value = "Client Name"
    .Range(rangesTable(I)).Cells(1, 3).Value = "Address"
    .Range(rangesTable(I)).Cells(1, 4).Value = "City"
    .Range(rangesTable(I)).Cells(1, 5).Value = "Postal Code"
    .Range(rangesTable(I)).Cells(1, 6).Value = "Regards"
    .Range(rangesTable(I)).Cells(1, 7).Value = "Phone:"
    .Range(rangesTable(I)).Cells(1, 8).Value = "Fax:"
    .Range(rangesTable(I)).Cells(1, 9).Value = "Cell:"
    .Range(rangesTable(I)).Cells(1, 10).Value = "Email:"
    .Range(rangesTable(I)).Cells(1, 11).Value = "Track#"
  End With
  On Error Resume Next
  Workbooks("TakeoffTables.xls").Close savechanges:=True
  Application.DisplayAlerts = True
  stopTimer = Timer
  'Debug.Print Format(stopTimer - startTimer, "#.000")
End Sub

'Private Sub Workbook_Activate()
'    Call AddToCellMenu
'End Sub

'Sub CopySheetToOtherWbk()
'  Dim CopyFromBook As Workbook
'  Dim CopyToWbk As Workbook
'  Dim ShToCopy As Worksheet
'
'  Set CopyFromWbk = Workbooks("TakeoffUtility4.xlsm")
'  Set ShToCopy = CopyFromBook.Worksheets("UpdatedTakeoff")
'  Set CopyToWbk = Workbooks("test.xls")
'
'  ShToCopy.Copy After:=CopyToWbk.Sheets("Sheet1")
' End Sub


'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'    Call DeleteFromCellMenu
'    With Application
'      .DisplayAlerts = False
'      .Cursor = xlNorthwestArrow
'      ThisWorkbook.SaveAs FileName:=ThisWorkbook.Path & "\TakeoffTables.xls"
'      .DisplayAlerts = True
'    End With
'    If ThisWorkbook.Name Like "*.xls" Then
'      initializationModule.setTableRanges
'    End If
'End Sub
'Sub DeleteAllCode()
' 'code source- http://www.vbaexpress.com/kb/getarticle.php?kb_id=93
'     'Trust Access To Visual Basics Project must be enabled.
'     'From Excel: Tools | Macro | Security | Trusted Sources
'
'    Dim x               As Integer
'    Dim Proceed         As VbMsgBoxResult
'    Dim Prompt          As String
'    Dim Title           As String
'
'    Prompt = "Are you certain that you want to delete all the VBA Code from " & _
'    ActiveWorkbook.Name & "?"
'    Title = "Verify Procedure"
'
'    Proceed = MsgBox(Prompt, vbYesNo + vbQuestion, Title)
'    If Proceed = vbNo Then
'        MsgBox "Procedure Canceled", vbInformation, "Procedure Aborted"
'        Exit Sub
'    End If
'
'    On Error Resume Next
'    With ActiveWorkbook.VBProject
'        For x = .VBComponents.Count To 1 Step -1
'            .VBComponents.Remove .VBComponents(x)
'        Next x
'        For x = .VBComponents.Count To 1 Step -1
'            .VBComponents(x).CodeModule.DeleteLines _
'            1, .VBComponents(x).CodeModule.CountOfLines
'        Next x
'    End With
'    On Error GoTo 0
'
'End Sub

'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
''  On Error Resume Next
''    updatedSheet.Activate
''    Application.WindowState = xlMaximized
''    'Range("A1").Select
''    'updatedSheet.activateClientTabButton.Select
''    With updatedSheet.confirmationLabel
''        .Caption = Range("clientName").Value
''        .BackColor = &HFF&
''        .Font = "System"
''        .ForeColor = vbBlack
''    End With
'    'updateWordingModule.correctWording
'    'sortModule.opiTags
'
'End Sub

'Private Sub test()
'  Application.DisplayAlerts = False
'  Dim startTimer As Single
'  Dim stopTimer As Single
'  startTimer = Timer
''  Dim oWorkbook As Excel.Workbook
''  Dim oSheet As Excel.Worksheet
''  Set oWorkbook = Workbooks("TakeoffUtility4.xlsm")
''  Set oSheet = oWorkbook.Worksheets("UpdatedTakeoff")
'  Workbooks("TakeoffUtility4.xlsm").Worksheets("UpdatedTakeoff").Cells.Select
''  Cells.Select
'  Selection.Copy
'  Workbooks.Add
'
'  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'  ChDir "C:\Users\mike\Documents\My Projects\Project.QT"
'
'  'FileFormat:=56 for xlExcel8, 97-2003 format w/o macros
'  ActiveWorkbook.SaveAs FileName:= _
'      ThisWorkbook.Path & "\test.xls", FileFormat _
'      :=56, Password:="", WriteResPassword:="", ReadOnlyRecommended:= _
'      False, CreateBackup:=False
'
''Sub setTableRanges()
''  Windows("TakeoffTables.xls").Activate
'  'include column headings with range for DAO connection
'  Dim rangesTable(1 To 15) As String
'  Dim I As Long
'  rangesTable(1) = "selectedExcavation"
'  rangesTable(2) = "selectedWater"
'  rangesTable(3) = "selectedFoundation"
'  rangesTable(4) = "selectedSeasonal"
'  rangesTable(5) = "selectedOPI"
'  rangesTable(6) = "selectedMoreOPI"
'  rangesTable(7) = "selectedSummaryOPI"
'  rangesTable(8) = "selectedWalls"
'  rangesTable(9) = "selectedOther"
'  rangesTable(10) = "selectedMaterials"
'  rangesTable(11) = "selectedExcavationPegasus"
'  rangesTable(12) = "selectedWaterPegasus"
'  rangesTable(13) = "selectedSpreadEagle"
'  rangesTable(14) = "selectedNewDivision"
'  rangesTable(15) = "ClientTable"
'
'  With ActiveWorkbook.Sheets(1)
'  .Range("G68:K82").Name = "selectedExcavation"
'  .Range("G84:K98").Name = "selectedWater"
'  .Range("G100:K114").Name = "selectedFoundation"
'  .Range("G116:K130").Name = "selectedSeasonal"
'  .Range("G132:K146").Name = "selectedOPI"
'  .Range("G148:K162").Name = "selectedMoreOPI"
'  .Range("G164:K178").Name = "selectedSummaryOPI"
'  .Range("G180:K194").Name = "selectedWalls"
'  .Range("G196:K210").Name = "selectedOther"
'  .Range("G229:K243").Name = "selectedMaterials"
'  .Range("G246:K260").Name = "selectedExcavationPegasus"
'  .Range("G263:K277").Name = "selectedWaterPegasus"
'  .Range("G280:K294").Name = "selectedSpreadEagle"
'  .Range("G297:K311").Name = "selectedNewDivision"
'  .Range("T1:AD10").Name = "ClientTable"
'
'  For I = 1 To 14
'    .Range(rangesTable(I)).Interior.Color = 10092543
'    .Range(rangesTable(I)).Cells(1, 1).Offset(-1, 0).Value = rangesTable(I)
'    .Range(rangesTable(I)).Cells(1, 1).Value = "Count"
'    .Range(rangesTable(I)).Cells(1, 2).Value = "Description"
'    .Range(rangesTable(I)).Cells(1, 3).Value = "Rate"
'    .Range(rangesTable(I)).Cells(1, 4).Value = "Units"
'    .Range(rangesTable(I)).Cells(1, 5).Value = "Price"
'  Next I
'  'ClientTable
'    .Range(rangesTable(I)).Interior.Color = 10092543
'    '.Range(rangesTable(I)).Cells(1, 1).Offset(-1, 0).Value = rangesTable(I)
'    .Range(rangesTable(I)).Cells(1, 1).Value = "Takeoff Filename"
'    .Range(rangesTable(I)).Cells(1, 2).Value = "Client Name"
'    .Range(rangesTable(I)).Cells(1, 3).Value = "Address"
'    .Range(rangesTable(I)).Cells(1, 4).Value = "City"
'    .Range(rangesTable(I)).Cells(1, 5).Value = "Postal Code"
'    .Range(rangesTable(I)).Cells(1, 6).Value = "Regards"
'    .Range(rangesTable(I)).Cells(1, 7).Value = "Phone:"
'    .Range(rangesTable(I)).Cells(1, 8).Value = "Fax:"
'    .Range(rangesTable(I)).Cells(1, 9).Value = "Cell:"
'    .Range(rangesTable(I)).Cells(1, 10).Value = "Email:"
'    .Range(rangesTable(I)).Cells(1, 11).Value = "Track#"
'  End With
'ActiveWorkbook.Save
'  Application.DisplayAlerts = True
'  ActiveWindow.Close
'  stopTimer = Timer
'  Debug.Print Format(stopTimer - startTimer, "#.000")
'End Sub


