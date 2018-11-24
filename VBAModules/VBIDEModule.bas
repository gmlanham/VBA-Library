Attribute VB_Name = "VBIDEModule"
Option Explicit
  Public Sub getProcedures()
    ' Declare variables to access the Excel 2007 workbook.
    Dim objXLApp As Excel.Application
    Dim objXLWorkbooks As Excel.Workbooks
    Dim objXLABC As Excel.Workbook
    
    ' Declare variables to access the macros in the workbook.
    Dim VBAEditor As VBIDE.VBE
    Dim objProject As VBIDE.VBProject
    Dim objComponent As VBIDE.VBComponent
    Dim objCode As VBIDE.CodeModule
    
    ' Declare other miscellaneous variables.
    Dim iLine As Integer
    Dim sProcName As String
    Dim pk As vbext_ProcKind
    
    Set VBAEditor = Application.VBE
    
    ' Open Excel and the open the workbook.
    Set objXLApp = New Excel.Application
    
    ' Empty the list box.
    ControlPanel.ProcedureListBox.Clear
    
    ' Get the project details in the workbook.
    Set objProject = VBAEditor.ActiveVBProject

    ' Iterate through each component in the project.
    For Each objComponent In objProject.VBComponents

        ' Find the code module for the project.
        Set objCode = objComponent.CodeModule

        ' Scan through the code module, looking for procedures.
        iLine = 1
        Do While iLine < objCode.CountOfLines
            sProcName = objCode.ProcOfLine(iLine, pk)
            If sProcName <> "" Then
                ' Found a procedure. Display its details, and then skip
                ' to the end of the procedure.
                ControlPanel.ProcedureListBox.AddItem objComponent.Name & ": " & sProcName
                iLine = iLine + objCode.ProcCountLines(sProcName, pk)
            Else
                ' This line has no procedure, so go to the next line.
                iLine = iLine + 1
            End If
        Loop
        Set objCode = Nothing
        Set objComponent = Nothing
    Next
    
    ' Clean up and exit.
    Set objProject = Nothing
    objXLApp.Quit
  End Sub
