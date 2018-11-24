Attribute VB_Name = "VBIDEModule"
Option Explicit
  Dim ModuleCount As Long
  Dim proceduresCount As Long
  Dim LineCount As Long

Public Sub viewSelectedItemWord()
  Dim ProcName As String
  Dim compString As String
  Dim folderName As String
  Dim I As Long
  Dim S As String
  On Error GoTo ErrorHandler
  Set OWord = GetObject(, "Word.Application")
  If Err Then
    Set OWord = New Word.Application
  End If

  Documents.Open fileName:=FileToGet, Visible:=False
  Set VBProj = Application.Documents(FileToGet).VBProject

  folderName = Documents("VBAUtility.docm").path & "\VBAModules\VBAModulesWord"
  Dim columnNumber As Long
  columnNumber = 0
  With ControlPanel.ProcedureListBox
    For I = 0 To .ListCount - 1
      If .Selected(I) Then
        S = .List(I, columnNumber)
        compString = Left(S, InStr(S, ":") - 1)
        Set VBComp = VBProj.VBComponents(compString)
        'Set CodeMod = VBComp.CodeModule
      
        ProcName = Right(S, Len(S) - (InStr(S, ":") + 1))
        ProcName = Left(ProcName, InStr(ProcName, ",") - 1)
        
        ExportVBComponent VBComp, folderName, ProcName
      End If
    Next I
  End With
Cleanup:
  Set VBProj = Nothing
  Set VBComp = Nothing
  'Set CodeMod = Nothing
Exit Sub
ErrorHandler:
  MsgBox "An error occured while processing 'viewSelectedItemWord'."
  Resume Cleanup
End Sub
Public Sub viewSelectedItemExcel()
  Dim ProcName As String
  Dim compString As String
  Dim folderName As String
  Dim fileName As String
  Dim I As Long
  Dim S As String
  
  On Error Resume Next
  Set OExcel = GetObject(, "Excel.Application")
  If Err Then
    Set OExcel = New Excel.Application
  End If
  
  Set OWB = OExcel.Workbooks.Open(FileToGet)
  OWB.Windows(1).Visible = False
  Set VBProj = OWB.VBProject
  
  On Error Resume Next
  folderName = Documents("VBAUtility.docm").path & "\VBAModules\VBAModulesExcel"
  Dim columnNumber As Long
  columnNumber = 0
  With ControlPanel.ProcedureListBox
    For I = 0 To .ListCount - 1
      If .Selected(I) Then
        S = .List(I, columnNumber)
        compString = Left(S, InStr(S, ":") - 1)
        Set VBComp = VBProj.VBComponents(compString)
        If VBComp Is Nothing Then fileName = compString & ".doc"
        'Set CodeMod = VBComp.CodeModule
      
        ProcName = Right(S, Len(S) - (InStr(S, ":") + 1))
        ProcName = Left(ProcName, InStr(ProcName, ",") - 1)
        
        ExportVBComponent VBComp, folderName, ProcName, fileName
        Exit Sub
      End If
    Next I
  End With
Cleanup:
  Set VBProj = Nothing
  Set VBComp = Nothing
  compString = vbNullString
  'Set CodeMod = Nothing
Exit Sub
ErrorHandler:
  MsgBox "An error occured while processing 'viewSelectedItemExcel'."
  Resume Cleanup
End Sub

Public Function TotalCodeLinesInVBComponent(VBComp As VBIDE.VBComponent) As Long
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' This returns the total number of code lines (excluding blank lines and
  ' comment lines) in the VBComponent referenced by VBComp. Returns -1
  ' if the VBProject is locked.
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim N As Long
  Dim S As String
  Dim LineCount As Long
  
  If VBComp.Collection.Parent.Protection = vbext_pp_locked Then
    TotalCodeLinesInVBComponent = -1
    Exit Function
  End If
  With VBComp.CodeModule
    For N = 1 To .CountOfLines
        S = .Lines(N, 1)
        If Trim(S) = vbNullString Then
            ' blank line, skip it
        ElseIf Left(Trim(S), 1) = "'" Then
            ' comment line, skip it
        Else
            LineCount = LineCount + 1
        End If
    Next N
  End With
  ModuleCount = ModuleCount + 1
  TotalCodeLinesInVBComponent = LineCount
End Function

Public Function TotalCodeLinesInProject(VBProj As VBIDE.VBProject) As Long
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
  For Each VBComp In VBProj.VBComponents
    LineCount = LineCount + TotalCodeLinesInVBComponent(VBComp)
  Next VBComp
  
  TotalCodeLinesInProject = LineCount
  Set VBComp = Nothing
End Function
Public Function getProceduresBAS(Optional ByVal ModuleToGet As String) As Boolean
  On Error GoTo ErrorHandler
  Dim CompName As String
  CompName = ModuleToGet
  Dim folderName As String
  folderName = Documents("VBAUtility.docm").path & "\VBAModules"
  CompName = Right(CompName, Len(CompName) - Len(folderName) - 1)
  Debug.Print CompName
  Set VBComp = VBProj.VBComponents(CompName)
  Set CodeMod = VBComp.CodeModule
  ExportVBComponent VBComp, folderName

Cleanup:
  Set VBProj = Nothing
  'oWord.Quit
  On Error Resume Next
  For Each ODoc In Documents
    If ODoc.Name <> "VBAUtility.docm" Then
      ODoc.Close SaveChanges:=False
      'oWord.Quit
    End If
  Next ODoc
  Exit Function
ErrorHandler:
  MsgBox "An error occured processing 'getProceduresWord'."
  Resume Cleanup
End Function
Private Sub test1()
  getProceduresWord FileToGet:=ThisDocument.path & "\WordApps\" & "DAOExcel8.docm"
End Sub
Public Function getProceduresWord(Optional ByVal FileToGet As String) As Boolean
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
  Documents.Open fileName:=FileToGet, Visible:=False
  Set VBProj = Application.Documents(FileToGet).VBProject
  LineCount = TotalCodeLinesInProject(VBProj)
  
  ' Iterate through each component in the project.
  For Each VBComp In VBProj.VBComponents

    ' Find the code module for the project.
    Set CodeMod = VBComp.CodeModule
    
    lastLine = 1
    iLine = 1
    proceduresCount = countProcedures(FileToGet)
    ReDim proceduresArray(proceduresCount + 1, 2)
    Dim I As Long
    On Error Resume Next
    Do While iLine < CodeMod.CountOfLines
      'find procedures names by searching each line
      sProcName = CodeMod.ProcOfLine(iLine, pk)
      If sProcName <> "" Then
          ' Found a procedure. Display its details, and then skip
          ' to the end of the procedure.
          iLine = iLine + CodeMod.ProcCountLines(sProcName, pk)
          
          ControlPanel.ProcedureListBox.AddItem VBComp.Name _
          & ": " & sProcName & ",  Lines: " & iLine - lastLine
          
          glob = glob & sProcName & vbTab & vbTab & iLine - lastLine & vbCr
          
          I = I + 1
          proceduresArray(I, 1) = sProcName
          proceduresArray(I, 2) = iLine - lastLine
          With Documents("VBAUtility.docm")
'               typetext or do something else
          End With
          lastLine = iLine
      Else
          ' This line has no procedure, so go to the next line.
          iLine = iLine + 1
      End If
    Loop
    'ActiveDocument.Close savechanges:=False
    Set CodeMod = Nothing
    Set VBComp = Nothing
  Next
  
  Dim fileName As String
  fileName = Right(FileToGet, Len(FileToGet) - InStrRev(FileToGet, "\"))
  fileName = Left(fileName, InStr(fileName, ".") - 1)
  
  Dim result As Long
  result = MsgBox(prompt:= _
    VBProj.Name & vbCr & _
    "Number of Modules   " & vbTab & ModuleCount & vbCr & _
    "Number of Procedure " & vbTab & proceduresCount & vbCr & _
    "Number of Lines     " & vbTab & vbTab & LineCount & _
    vbCr & vbCr & _
    "Click 'OK' to type procedures on the page.", _
    Buttons:=vbOKCancel, Title:=fileName & " Procedures")
  
  With Selection
    If result = vbOK Then
      .TypeText fileName & " Procedures:" & vbCr & "Procedure Name" & vbTab & vbTab & "Lines" & vbCr
      .TypeText glob & vbCr
      getProceduresWord = False
    Else
      getProceduresWord = True
    End If
  End With
Cleanup:
  Set VBProj = Nothing
  On Error Resume Next
  For Each ODoc In Documents
    If ODoc.Name <> "VBAUtility.docm" Then
      ODoc.Close SaveChanges:=False
    End If
  Next ODoc
  Exit Function
ErrorHandler:
  MsgBox "An error occured processing 'getProceduresWord'."
  Resume Cleanup
End Function

Sub SyncVBAEditor()
'=======================================================================
' SyncVBAEditor
' This syncs the editor with respect to the ActiveVBProject and the
' VBProject containing the ActiveCodePane. This makes the project
' that conrains the ActiveCodePane the ActiveVBProject.
'=======================================================================
  With Application.VBE
    If Not .ActiveCodePane Is Nothing Then
        Set .ActiveVBProject = .ActiveCodePane.CodeModule.Parent.Collection.Parent
    End If
  End With
End Sub

'Public Enum ProcScope
'        ScopePrivate = 1
'        ScopePublic = 2
'        ScopeFriend = 3
'        ScopeDefault = 4
'    End Enum
'
'    Public Enum LineSplits
'        LineSplitRemove = 0
'        LineSplitKeep = 1
'        LineSplitConvert = 2
'    End Enum
'
'    Public Type ProcInfo
'        ProcName As String
'        ProcKind As VBIDE.vbext_ProcKind
'        ProcStartLine As Long
'        ProcBodyLine As Long
'        ProcCountLines As Long
'        ProcScope As ProcScope
'        ProcDeclaration As String
'    End Type
Sub ShowProcedureInfo()
  Dim VBProj As VBIDE.VBProject
  Dim VBComp As VBIDE.VBComponent
  Dim CodeMod As VBIDE.CodeModule
  Dim CompName As String
  Dim ProcName As String
  Dim ProcKind As VBIDE.vbext_ProcKind
  Dim PInfo As ProcInfo
  
  CompName = "BrowseModule"
  ProcName = "browseFiles"
  ProcKind = vbext_pk_Proc
  
  Set VBProj = ActiveDocument.VBProject
  Set VBComp = VBProj.VBComponents(CompName)
  Set CodeMod = VBComp.CodeModule
  
  PInfo = ProcedureInfo(ProcName, ProcKind, CodeMod)
  
'        Debug.Print "ProcName: " & PInfo.ProcName
'        Debug.Print "ProcKind: " & CStr(PInfo.ProcKind)
'        Debug.Print "ProcStartLine: " & CStr(PInfo.ProcStartLine)
'        Debug.Print "ProcBodyLine: " & CStr(PInfo.ProcBodyLine)
'        Debug.Print "ProcCountLines: " & CStr(PInfo.ProcCountLines)
'        Debug.Print "ProcScope: " & CStr(PInfo.ProcScope)
'        Debug.Print "ProcDeclaration: " & PInfo.ProcDeclaration
End Sub

Function ProcedureInfo(ProcName As String, ProcKind As VBIDE.vbext_ProcKind, _
  CodeMod As VBIDE.CodeModule) As ProcInfo

  Dim PInfo As ProcInfo
  Dim BodyLine As Long
  Dim Declaration As String
  Dim FirstLine As String
  
  
  BodyLine = CodeMod.ProcStartLine(ProcName, ProcKind)
  If BodyLine > 0 Then
      With CodeMod
          PInfo.ProcName = ProcName
          PInfo.ProcKind = ProcKind
          PInfo.ProcBodyLine = .ProcBodyLine(ProcName, ProcKind)
          PInfo.ProcCountLines = .ProcCountLines(ProcName, ProcKind)
          PInfo.ProcStartLine = .ProcStartLine(ProcName, ProcKind)
          
          FirstLine = .Lines(PInfo.ProcBodyLine, 1)
          If StrComp(Left(FirstLine, Len("Public")), "Public", vbBinaryCompare) = 0 Then
              PInfo.ProcScope = ScopePublic
          ElseIf StrComp(Left(FirstLine, Len("Private")), "Private", vbBinaryCompare) = 0 Then
              PInfo.ProcScope = ScopePrivate
          ElseIf StrComp(Left(FirstLine, Len("Friend")), "Friend", vbBinaryCompare) = 0 Then
              PInfo.ProcScope = ScopeFriend
          Else
              PInfo.ProcScope = ScopeDefault
          End If
          PInfo.ProcDeclaration = GetProcedureDeclaration(CodeMod, ProcName, ProcKind, LineSplitKeep)
      End With
  End If
  
  ProcedureInfo = PInfo

End Function
    
Public Function GetProcedureDeclaration(CodeMod As VBIDE.CodeModule, _
  ProcName As String, ProcKind As VBIDE.vbext_ProcKind, _
  Optional LineSplitBehavior As LineSplits = LineSplitRemove)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetProcedureDeclaration
' This return the procedure declaration of ProcName in CodeMod. The LineSplitBehavior
' determines what to do with procedure declaration that span more than one line using
' the "_" line continuation character. If LineSplitBehavior is LineSplitRemove, the
' entire procedure declaration is converted to a single line of text. If
' LineSplitBehavior is LineSplitKeep the "_" characters are retained and the
' declaration is split with vbNewLine into multiple lines. If LineSplitBehavior is
' LineSplitConvert, the "_" characters are removed and replaced with vbNewLine.
' The function returns vbNullString if the procedure could not be found.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim LineNum As Long
  Dim S As String
  Dim Declaration As String
  
  On Error Resume Next
  LineNum = CodeMod.ProcBodyLine(ProcName, ProcKind)
  If Err.Number <> 0 Then
      Exit Function
  End If
  S = CodeMod.Lines(LineNum, 1)
  Do While Right(S, 1) = "_"
      Select Case True
          Case LineSplitBehavior = LineSplitConvert
              S = Left(S, Len(S) - 1) & vbNewLine
          Case LineSplitBehavior = LineSplitKeep
              S = S & vbNewLine
          Case LineSplitBehavior = LineSplitRemove
              S = Left(S, Len(S) - 1) & " "
      End Select
      Declaration = Declaration & S
      LineNum = LineNum + 1
      S = CodeMod.Lines(LineNum, 1)
  Loop
  Declaration = SingleSpace(Declaration & S)
  GetProcedureDeclaration = Declaration

End Function
    
Private Function SingleSpace(ByVal Text As String) As String
  Dim Pos As String
  Pos = InStr(1, Text, Space(2), vbBinaryCompare)
  Do Until Pos = 0
    Text = Replace(Text, Space(2), Space(1))
    Pos = InStr(1, Text, Space(2), vbBinaryCompare)
  Loop
  SingleSpace = Text
End Function

Public Function countProcedures(ByVal FileToGet As String)
 Dim dc As Document
 Dim VBComp As VBComponent
 Dim I As Long
 Dim strCodeLine As String
 Dim strTemp As String
 Dim lastLine As Long
 Dim countModules As Long
 Dim totalProcedures As Long
 Documents.Open fileName:=FileToGet, Visible:=False
 Set dc = Documents(FileToGet)
 countModules = 0
 totalProcedures = 0
 For Each VBComp In dc.VBProject.VBComponents
   With VBComp.CodeModule
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
 Next VBComp
 countProcedures = totalProcedures
End Function

