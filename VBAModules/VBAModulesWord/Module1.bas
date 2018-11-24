Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Dim TakeoffApp As Excel.Application
Dim TakeoffWorkbook As Excel.Workbook
Dim WorkbookPath As String
Dim Extras As Excel.Worksheet

Dim VBAEditor As VBIDE.VBE
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
Dim CodeMod As VBIDE.CodeModule
Dim Count As Integer
Dim extrasListGlob As String
Dim extrasArray() As String

'this procedure opens Takeoff
Sub extrasUpdate()
'This procedure loads the ExtrasListBox with all possible Extras (49) from Takeoff.
Dim extrasAllCount As Integer
Dim I As Integer
Dim ExcelWasNotRunning As Boolean
    Set VBAEditor = Application.VBE
    Set VBProj = VBAEditor.ActiveVBProject
   ' Set VBComp = VBProj.VBComponents("extrasListModule")
    'Set CodeMod = VBComp.CodeModule
    
    DoEvents
    On Error Resume Next
    Count = CodeMod.CountOfLines + 1

'open the Takeoff
    On Error Resume Next
    Set TakeoffApp = GetObject(, "Excel.Application")
    If Err Then
        Set TakeoffApp = New Excel.Application
        ExcelWasNotRunning = True
    End If

    WorkbookPath = ("C:\Users\mike\Documents\Takeoff.xlsm")

    'open the excel workbook to Word's eyes
    Set TakeoffWorkbook = TakeoffApp.Workbooks.Open(WorkbookPath)
    'ThisDocument.Path & "/" & "Takeoff.xlsm"
    'TakeoffApp.Visible = True
    
'Store the total number of rows in the Takeoff Extras data section.
    Set Extras = TakeoffApp.Worksheets("Extras")
    Extras.Activate
    extrasAllCount = 1 + TakeoffApp.Cells(1, 9).Value

'redimension the dynamic array
    ReDim extrasArray(extrasAllCount)
   
   
'switch this statement so that instead of getting the Cells value, it sets them
    For I = 1 To extrasAllCount
        TakeoffApp.Cells(I, 10).Value = extrasArray(I)
    Next 'i
 
   
'Loop through the rows to store the extras in the array
    'For I = 1 To extrasAllCount
    '    extrasArray(I) = TakeoffApp.Cells(I, 10).Value
    'Next 'i

'pass extrasArray to the extrasListbox
'    ExtrasUpdateForm.extrasListBox.List = extrasArray
    
'update the listbox item list with a fresh list
    'create a glob called "listGlob" for the InsertLines method
    extrasListGlob = "" & "Public Sub extrasList" & vbCr
    For I = 1 To extrasAllCount
       extrasListGlob = extrasListGlob & "    " & "ExtrasUpdateForm.extrasListBox.AddItem " & """" & extrasArray(I) & """" & vbCr
    Next
    extrasListGlob = extrasListGlob & "End Sub"
    
'delete old list
   ' deleteOldExtras
  
    'this is the command that actually inserts the list into the ExtraUpdateForm code behind
    CodeMod.InsertLines Count, extrasListGlob

'Clean-up and close
   ' ExtrasUpdateForm.extrasListBox.ListIndex = 0
    
    'save, close and quit
    'TakeoffApp.SaveAs "C:\Users\mike\Documents\Takeoff.xlsm"
    Application.DisplayAlerts = False
        TakeoffWorkbook.Close savechanges:=False
    Application.DisplayAlerts = True

    TakeoffApp.Quit
    SendKeys "%{F4}", True
End Sub


'this procedure opens a Word document containing the data and puts that data into an array
Private Sub UserForm() '_Initialize()
Dim sourcedoc As Document
Dim I As Long, j As Long, m As Long, n As Long
Dim pColWidths As String
'Define an array to be loaded with the data
Dim arrData() As Variant
Application.ScreenUpdating = False
'Open the file containing the table with items to load
Set sourcedoc = Documents.Open(FileName:="F:\Data Stores\PopulateMultiColumnListBoxDataSource.docx")
'Get the number members = number of rows in the table of details less one for the header row
I = sourcedoc.Tables(1).Rows.Count - 1
'Get the number of columns in the table of details
j = sourcedoc.Tables(1).Columns.Count
'Set the number of columns in the Listbox to match the number of columns in the table of details
ListBox1.ColumnCount = j
'Dimension arrData
ReDim arrData(I - 1, j - 1)
'Load table data into arrData
For n = 0 To j - 1
  For m = 0 To I - 1
    arrData(m, n) = Main.fcnCellText(sourcedoc.Tables(1).Cell(m + 2, n + 1))
  Next m
Next n
'Build ColumnWidths statement
pColWidths = "50"
For n = 2 To j
pColWidths = pColWidths + ";0"
Next n
'Load data into ListBox1
With ListBox1
  .List() = arrData
  .ColumnWidths = pColWidths 'Apply ColumnWidths statement
End With
'Close the file containing the client details
sourcedoc.Close savechanges:=wdDoNotSaveChanges
End Sub

Private Sub CommandButton1_Click()
'Write column data to named bookmarks in document
With ActiveDocument
  Main.FillBMs .Bookmarks("Name"), Me.ListBox1.Column(0)
  Main.FillBMs .Bookmarks("Email"), Me.ListBox1.Column(1)
  Main.FillBMs .Bookmarks("PhoneNumber"), Me.ListBox1.Column(2)
End With
Me.Hide
End Sub
