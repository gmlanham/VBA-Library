Attribute VB_Name = "modInsertProcNames"
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modInsertProcNames
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
'
' This module contains the InsertProcedureNameIntoProcedure and supporting procedures.
' InsertProcedureNameIntoProcedure will insert a CONST statement at the top of each
' procedure in the Application.VBE.ActiveCodePane.CodeModule. The user selects the
' name of the constant (e.g., "C_PROC_NAME") and the code insert statements like
'       Const C_PROC_NAME = "InsertProcedureNameIntoProcedure"
' at the beginning of each procedure. It supports procedures whose Declaration spans
' more than one line, e.g.,
'       Public Function Test( X As Integer, _
'                             Y As Integer, _
'                             Z As Integer)
' If comment lines appear DIRECTLY below the procedured declaration (no blank lines
' between the declaration and the start of the comments), the CONST statement is
' placed directly below the comment block. If a constant already exists with name
' user specified name, that constant declaration is deleted and replace with the
' new CONST line.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const C_MSGBOX_TITLE = "Insert Procedure Names"
Private Const C_VBE_CONST_TAG = "__INSERTCONSTLINE__"
Private Const C_VBE_INSERT_MENU As Long = 30005


Sub InsertProcedureNameIntoProcedures()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' InsertProcedureNameIntoProcedures
' This procedure inserts a CONST statement in each procedure in the
' ActiveCodePane.CodeModule. The user is prompted for the name of
' the constant to add. If that constant already exists within each
' procedure, that line of code is deleted and replace with the
' new CONST statement.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const C_PROC_NAME = "InsertProcedureNameIntoProcedure"
Dim ProcName As String
Dim ProcLine As String
Dim ProcType As VBIDE.vbext_ProcKind
Dim StartLine As Long
Dim Msg As String
Dim VBComp As VBIDE.VBComponent
Dim CodeMod As VBIDE.CodeModule
Dim Ndx As Long
Dim Res As Variant
Dim Done As Boolean
Dim ProcBodyLine As Long
Dim SaveProcName As String
Dim ConstName As String
Dim ValidConstName As Boolean
Dim ConstAtLine As Long
Dim EndOfDeclaration As Long


''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure there is an active project.
''''''''''''''''''''''''''''''''''''''''''''''''''
If Application.VBE.ActiveVBProject Is Nothing Then
    MsgBox "There is no active project.", vbOKOnly, C_MSGBOX_TITLE
    Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensuse the ActiveProject is not locked.
''''''''''''''''''''''''''''''''''''''''''''''''''
If Application.VBE.ActiveVBProject.Protection = vbext_pp_locked Then
    MsgBox "The active project is locked.", vbOKOnly, C_MSGBOX_TITLE
    Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure there is an active code pane
''''''''''''''''''''''''''''''''''''''''''''''''''
If Application.VBE.ActiveCodePane Is Nothing Then
    MsgBox "There is no active code pane.", vbOKOnly, C_MSGBOX_TITLE
    Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''
' Prompt the user for the name of the constant
' to insert.
''''''''''''''''''''''''''''''''''''''''''''''''''
'ConstName = InputBox(prompt:="Enter a constant name (e.g. 'C_PROC_NAME') that will be used as " & vbCrLf & _
'    "the constant in which to store the procedure name.", Title:=C_MSGBOX_TITLE)
'If Trim(ConstName) = vbNullString Then
'    Exit Sub
'End If

'hardcode the constant name
ConstName = "ProcedureNameConstant"


''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure that ConstName is a valid name for a
' constant declaration.
'''''''''''''''''''''''''''''''''''''''''''''''''
If IsValidConstantName(ConstName) = False Then
    MsgBox "The constant name: '" & ConstName & "' is invalid.", vbOKOnly, C_MSGBOX_TITLE
    Exit Sub
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the active code module.
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Set CodeMod = Application.VBE.ActiveCodePane.CodeModule


'''''''''''''''''''''''''''''''''''''''''''''''''''
' Skip past any Option statement and any module-level
' variable declations. Start at the first procuedure
' in the module.
'''''''''''''''''''''''''''''''''''''''''''''''''''
StartLine = CodeMod.CountOfDeclarationLines + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the procedure name that is at StartLine.
''''''''''''''''''''''''''''''''''''''''''''''''''''
ProcName = CodeMod.ProcOfLine(StartLine, ProcType)
SaveProcName = ProcName

Do Until Done
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Loop through all procedures in the module.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get the body proc line (the actual declaration line,
    ' ignoring commnets
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ProcBodyLine = CodeMod.ProcBodyLine(ProcName, ProcType)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' See if the constant declaration already exists.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ConstAtLine = ConstNameInProcedure(ConstName, CodeMod, ProcName, ProcType)
    If ConstAtLine > 0 Then
        '''''''''''''''''''''''''''''''''''''''''''
        ' Const line already exist. Delete it and
        ' replace it.
        '''''''''''''''''''''''''''''''''''''''''''
        CodeMod.DeleteLines ConstAtLine, 1
        CodeMod.InsertLines ConstAtLine, "CONST " & ConstName & " = " & Chr(34) & ProcName & Chr(34)
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Skip past the declaration lines and the comment lines that
        ' immediately follow the declarations (no blank lines between
        ' the declarations and the comments).
        ' Insert the CONST declaration.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        EndOfDeclaration = EndOfDeclarationLines(CodeMod, ProcName, ProcType)
        ProcLine = EndOfCommentOfProc(CodeMod, EndOfDeclaration + 1)
        CodeMod.InsertLines ProcLine + 1, "CONST " & ConstName & " = " & Chr(34) & ProcName & Chr(34)
    End If
  
    ''''''''''''''''''''''''''''''''''''''''''
    ' Skip StartLine to the next proc
    ''''''''''''''''''''''''''''''''''''''''''
    StartLine = ProcBodyLine + CodeMod.ProcCountLines(ProcName, ProcType) + 1
    ProcName = CodeMod.ProcOfLine(StartLine, ProcType)
    
    ''''''''''''''''''''''''''''''''
    ' Special handling for the last
    ' procedure in the module in case
    ' it has blank lines following the
    ' end of the procedure body.
    '''''''''''''''''''''''''''''''''
    If ProcName = SaveProcName Then
        Done = True
    Else
        SaveProcName = ProcName
    End If
Loop

End Sub

Function EndOfCommentOfProc(CodeMod As VBIDE.CodeModule, ProcBodyLine As Long) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EndOfCommentOfProc
' This returns the line number of the last comment line in a comment block that IMMEDIATELY
' follow the procedure declaration. For example, with the following code
'       Function MyTest(X As Integer, _
'                       Y As Integer, _
'                       Z As Ineger)
'       '''''''''''''''''''''''''''''''''''' START COMMENT BLOCK - NO BLANK LINES ABOVE
'       ' Some Comments
'       '''''''''''''''''''''''''''''''''''' END COMMENT BLOCK
'
' the function will return the line number of "END COMMENT BLOCK". There must be no
' blank lines between "Z As Integer)" and the first comment line.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim Done As Boolean
    Dim LineNum As String
    Dim LineText As String
    
    LineNum = ProcBodyLine

    Do Until Done
        LineNum = LineNum + 1
        LineText = CodeMod.Lines(LineNum, 1)
        If Left(Trim(LineText), 1) = "'" Then
            Done = False
        Else
            Done = True
        End If
    Loop
    EndOfCommentOfProc = LineNum - 1
End Function


Function IsValidConstantName(ConstName As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsValidConstantName
' This returns True or False indicating whether ConstName
' is a valid constant name. ConstName must not contain any
' spaces, the left-most character must be a letter, and the
' rest of the characters must be alpha or numeric or underscore
' character. Any other character is considered invalid.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const C_PROC_NAME = "IsValidConstantName"

Dim C As String
Dim N As Long
Dim CAsc As Integer
If InStr(1, ConstName, " ") > 0 Then
    IsValidConstantName = False
    Exit Function
End If
If IsNumeric(Left(ConstName, 1)) = True Then
    IsValidConstantName = False
    Exit Function
End If
For N = 2 To Len(ConstName)
    C = Mid(ConstName, N, 1)
    CAsc = Asc(UCase(C))
    Select Case CAsc
        Case Asc("A") To Asc("Z")
        Case Asc("0") To Asc("9")
        Case Asc("_")
        Case Else
            IsValidConstantName = False
            Exit Function
    End Select
Next N
IsValidConstantName = True


End Function


Function ConstNameInProcedure(ConstName As String, CodeMod As VBIDE.CodeModule, _
    ProcName As String, _
    ProcType As VBIDE.vbext_ProcKind) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ConstNameInProcedure
' This returns the line number containing the existing constant declaration,
' or 0 if the procedure does not contain the constant declaration.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const C_PROC_NAME = "ConstNameInProcedure"
Dim LineNum As Long
Dim LineText As String
Dim ProcBodyLine As Long

ProcBodyLine = CodeMod.ProcBodyLine(ProcName, ProcType)
For LineNum = ProcBodyLine To ProcBodyLine + CodeMod.ProcCountLines(ProcName, ProcType)
    LineText = CodeMod.Lines(LineNum, 1)
    If InStr(LineText, " " & ConstName & " ") > 0 Then
        ConstNameInProcedure = LineNum
        Exit Function
    End If
Next LineNum
End Function

Function EndOfDeclarationLines(CodeMod As VBIDE.CodeModule, ProcName As String, _
    ProcType As VBIDE.vbext_ProcKind) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EndOfDeclarationLines
' This return the line number of the last declation lines. This is used to find
' the end of declarations when the declaration span more than one line of text
' in the code. For example, with the code
'       Function MyTest(X As Integer, _
'                       Y As Integer, _
'                       Z As Integer)
' it will return the line number of "Z As Integer)", the line number of the
' end of the declarations block.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const C_PROC_NAME = "EndOfDeclarationLines"
Dim LineNum As Long
Dim LineText As String

LineNum = CodeMod.ProcBodyLine(ProcName, ProcType)
Do Until Right(CodeMod.Lines(LineNum, 1), 1) <> "_"
    LineNum = LineNum + 1
Loop

EndOfDeclarationLines = LineNum


End Function

