Attribute VB_Name = "VBIDEModule"
Option Explicit
  ' Declare variables to access the macros in the workbook.
  Dim VBAEditor As VBIDE.VBE
  Dim VBProj As VBIDE.VBProject
  Dim vbComp As VBIDE.VBComponent
  Dim CodeMod As VBIDE.CodeModule

  Dim ModuleCount As Long
  Dim LineCount As Long

Public Function TotalCodeLinesInVBComponent(vbComp As VBIDE.VBComponent) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This returns the total number of code lines (excluding blank lines and
' comment lines) in the VBComponent referenced by VBComp. Returns -1
' if the VBProject is locked.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim N As Long
  Dim s As String
  Dim LineCount As Long
  
  If vbComp.Collection.Parent.Protection = vbext_pp_locked Then
      TotalCodeLinesInVBComponent = -1
      Exit Function
  End If
  With vbComp.CodeModule
      For N = 1 To .CountOfLines
          s = .Lines(N, 1)
          If Trim(s) = vbNullString Then
              ' blank line, skip it
          ElseIf Left(Trim(s), 1) = "'" Then
              ' comment line, skip it
          Else
              LineCount = LineCount + 1
          End If
      Next N
  End With
  ModuleCount = ModuleCount + 1
  TotalCodeLinesInVBComponent = LineCount
End Function

Public Function TotalCodeLinesInProject(ByVal VBProj As VBIDE.VBProject) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This returns the total number of code lines (excluding blank lines and
' comment lines) in all VBComponents of VBProj. Returns -1 if VBProj
' is locked.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  'Dim VBComp As VBIDE.VBComponent
  Dim LineCount As Long
  ModuleCount = 0
  If VBProj.Protection = vbext_pp_locked Then
      TotalCodeLinesInProject = -1
      Exit Function
  End If
  For Each vbComp In VBProj.VBComponents
    LineCount = LineCount + TotalCodeLinesInVBComponent(vbComp)
  Next vbComp
  
  TotalCodeLinesInProject = LineCount
  Set vbComp = Nothing
End Function

Public Function getProceduresExcel(Optional ByVal DocumentToGet As String, Optional ByRef proceduresCount As Long) As Boolean
'get procedures from VBIDE, to populate listbox
Initialization:
  On Error GoTo ErrorHandler
  Dim proceduresArray()
  Dim sProcName As String
  Dim pk As vbext_ProcKind
  Dim iLine As Integer
  Dim lastLine As Long
  Dim glob As String
  lastLine = 1
  iLine = 1
  Set VBAEditor = Application.VBE
  ControlPanel.ProcedureListBox.Clear
  
  ' Get the project details in the workbook.
  Set VBProj = VBAEditor.VBProjects(2)  '("VBAProject")
  
Main:
  'call lineCount function to get number of lines in VBProj
  LineCount = TotalCodeLinesInProject(VBProj)

  ' Iterate through each component in the project.
  For Each vbComp In VBProj.VBComponents

    ' Find the code module for the project.
    Set CodeMod = vbComp.CodeModule
    
    lastLine = 1
    iLine = 1
    proceduresCount = countProcedures(DocumentToGet)
    ReDim proceduresArray(proceduresCount + 1, 2)
    Dim I As Long
    
    Do While iLine < CodeMod.CountOfLines
      sProcName = CodeMod.ProcOfLine(iLine, pk)
      If sProcName <> "" Then
        ' Found a procedure. Display its details, and then skip
        ' to the end of the procedure.
        iLine = iLine + CodeMod.ProcCountLines(sProcName, pk)
        ControlPanel.ProcedureListBox.AddItem vbComp.Name _
        & ": " & sProcName & ",  Lines: " & iLine - lastLine
        glob = glob & sProcName & vbTab & iLine - lastLine & vbCr
        I = I + 1
        proceduresArray(I, 1) = vbComp.Name & ": " & sProcName
        proceduresArray(I, 2) = iLine - lastLine
        With Workbooks("VBAUtility.xlsm").Sheets("ProceduresReport")
          .Range("A" & I + 2).Value = proceduresArray(I, 1)
          .Range("B" & I + 2).Value = proceduresArray(I, 2)
        End With
        lastLine = iLine
      Else
        ' This line has no procedure, so go to the next line.
        iLine = iLine + 1
      End If
    Loop
    Set CodeMod = Nothing
    Set vbComp = Nothing
  Next
  
  'populate spreadsheet with procedure names
  Dim fileName As String
  fileName = Right(DocumentToGet, Len(DocumentToGet) - InStrRev(DocumentToGet, "\"))
  fileName = Left(fileName, InStr(fileName, ".") - 1)

    With Workbooks("VBAUtility.xlsm").Sheets("ProceduresReport")
      .Range("A1").Value = fileName
      .Range("A2").Value = "Procedure Name"
      .Range("B2").Value = "Lines"
    End With
    
Cleanup:
  On Error Resume Next
  For Each oWB In Workbooks
    If oWB.Name <> "VBAUtility.xlsm" Then oWB.Close savechanges:=False
  Next oWB
  Exit Function
  
ErrorHandler:
  MsgBox "Error while processing 'getProceduresExcel'."
  Resume Cleanup
End Function

Public Sub viewSelectedItem()
'determine the item selected in listbox, parse the item for procedure name
Initialization:
  Dim ProcName As String
  Dim compString As String
  Dim FolderName As String
  Dim I As Long
  Dim s As String
  On Error Resume Next
  If Workbooks.Count = 1 Then
    Set oWB = Workbooks.Open(FileToGet)
    oWB.Windows(1).Visible = False
  End If
  Set VBProj = VBAEditor.VBProjects(2)  '("VBAProject")
  
Main:
  On Error GoTo ErrorHandler
  FolderName = Workbooks("VBAUtility.xlsm").path & "\VBAModules"
  Dim columnNumber As Long
  columnNumber = 0
  With ControlPanel.ProcedureListBox
    For I = 0 To .ListCount - 1
      If .Selected(I) Then
        s = .List(I, columnNumber)
        compString = Left(s, InStr(s, ":") - 1)
        Set vbComp = VBProj.VBComponents(compString)
        'Set CodeMod = vbComp.CodeModule
      
        ProcName = Right(s, Len(s) - (InStr(s, ":") + 1))
        ProcName = Left(ProcName, InStr(ProcName, ",") - 1)
        
        'call procedure for export to Word
        Call ExportVBComponent(vbComp, FolderName, ProcName)
      End If
    Next I
  End With
  
Cleanup:

Exit Sub
ErrorHandler:
  MsgBox "Error in 'VBIDEModule.viewSelectedItem' procedure."
Resume Cleanup
End Sub

Public Function countProcedures(ByVal FileToGet As String)
'count procedures in Project
Initialization:
  Dim dc As Workbook
  Dim vbComp As VBComponent
  Dim I As Long
  Dim strCodeLine As String
  Dim strTemp As String
  Dim lastLine As Long
  Dim countModules As Long
  Dim totalProcedures As Long
  countModules = 0
  'Debug.Print FileToGet
  totalProcedures = 0
Main:
  For Each vbComp In Workbooks(2).VBProject.VBComponents
    With vbComp.CodeModule
     countProcedures = 0
     countModules = countModules + 1
        For I = .CountOfDeclarationLines + 1 To .CountOfLines
            If Trim(.Lines(I, 1)) <> "" Then
                If strTemp <> .Lines _
                        (.ProcBodyLine(.ProcOfLine(I, _
                        vbext_pk_Proc), vbext_pk_Proc), 1) Then
                    
                   countProcedures = countProcedures + 1
                   totalProcedures = totalProcedures + 1
                   strTemp = .Lines(.ProcBodyLine(.ProcOfLine(I, _
                                vbext_pk_Proc), vbext_pk_Proc), 1)
               End If
            End If
        Next I
    End With
  Next vbComp
  'set this functions return value
  countProcedures = totalProcedures
End Function

