Attribute VB_Name = "FileSystemModule"
Option Explicit
Public Sub LogErrorToFile()
    'http://www.everythingaccess.com/simplyvba/globalerrorhandler/callstack.htm
    Dim FileNum As Long
    Dim LogLine As String
    
    On Error Resume Next ' If this procedure fails, something fairly major has gone wrong.
    
    FileNum = FreeFile
    Open "C:\MyTest\ErrorLog.txt" For Append Access Write Lock Write As FileNum

        Print #FileNum, Now() & " - " & CStr(Err.Number) & " - " & CStr(Err.Description)
            
        'We will seperate the call stack onto seperate lines in the log
  With Err.CallStack
            Do
                Print #FileNum, "       --> " & .ProjectName & "." & _
                                .ModuleName & "." & _
                                .procedureName & ", " & _
                                "#" & .lineNumber & ", " & _
                                .LineCode & vbCrLf
                Debug.Print "       --> " & .ProjectName & "." & .ModuleName & "." _
                                & .procedureName & ", " _
                                & "#" & .lineNumber _
                                & ", " & .LineCode & vbCrLf _
                                
            Loop While .NextLevel
  End With

    Close FileNum

End Sub

Public Sub saveProcedureList()
  On Error Resume Next
  Dim objFSO As Variant
  Dim fldr As Variant
  Dim txtFile As Variant
  Dim I As Integer
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set fldr = objFSO.CreateFolder("C:\MyTest")
  Set txtFile = objFSO.CreateTextFile("C:\MyTest\testfile.txt", True)
  
  For I = 0 To ControlPanel.ProcedureListBox.ListCount - 1
      txtFile.Write (ControlPanel.ProcedureListBox.List(I)) & vbCrLf
  Next
Cleanup:
  txtFile.Close
End Sub

Public Sub ReadSystemIni()
  On Error GoTo ErrorHandler

  'Declare variables.
  Dim fso As New FileSystemObject
  Dim ts As TextStream
  Dim fileString As String
  'Open file.
  Set ts = fso.OpenTextFile(Environ("windir") & "\system.ini")
  'Loop while not at the end of the file.
  Do While Not ts.AtEndOfStream
    fileString = fileString & vbCr & ts.ReadLine
  Loop
  MsgBox prompt:="File Content= " & fileString, Buttons:=vbInformation + vbOKOnly, Title:="System.ini"
  'Close the file.
Cleanup:
   ts.Close
  Exit Sub
ErrorHandler:
  LogErrorToFile
  Resume Cleanup:
End Sub

Public Sub SizeSystemIni()
  On Error GoTo ErrorHandler

   Dim fso As New FileSystemObject
   Dim f As File
   'Get a reference to the File object.
   Set f = fso.GetFile(Environ("windir") & "\system.ini")
   MsgBox prompt:="File Size= " & f.Size, Buttons:=vbInformation + vbOKOnly, Title:="System.ini File Size"
Cleanup:
  Set fso = Nothing
  Set f = Nothing
  Exit Sub
ErrorHandler:
  LogErrorToFile
  Resume Cleanup:
End Sub

Public Function ListWinDirFiles() As Boolean
  On Error GoTo ErrorHandler

Const procedureName = "ListWinDirFiles"
On Error GoTo ErrorHandler:
   Dim fso As New FileSystemObject
   Dim f As Folder
   Dim sf As Folder
   Dim path As String
   Dim fileString As String
   Dim fileInfoArray() As String
   Dim I As Long
   I = 0
   'Initialize path.
   path = Environ("windir")
   'Get a reference to the Folder object.
   Set f = fso.GetFolder(path)
   'Iterate through subfolders.
   ReDim fileInfoArray(f.SubFolders.Count, 2)
   For Each sf In f.SubFolders
    I = I + 1
    fileInfoArray(I, 1) = sf.Name
    fileInfoArray(I, 2) = sf.DateLastAccessed
    fileString = fileString & vbCr & sf.Name & vbTab & sf.DateLastAccessed
   Next
   
  Dim result As Long
  result = MsgBox(prompt:=fileString, Buttons:=vbOKCancel, Title:="WinDir Files")
  With Selection
    If result = vbOK Then
      .TypeText "List of Files:" & vbCr & "File Name" & vbTab & "Date Last Accessed"
      .TypeText fileString
      ListWinDirFiles = False
    Else
      ListWinDirFiles = True
    End If
  End With
Cleanup:
  Set f = Nothing
  Exit Function
ErrorHandler:
  MsgBox prompt:="Error Number= " & Err.Number & vbTab & Err.Description & vbCr & "Procedure= " & procedureName, Buttons:=vbOKOnly + vbExclamation, Title:="Error"
  LogErrorToFile
  Resume Cleanup:
End Function

Public Sub ShowDrivePath()
  On Error GoTo ErrorHandler

   Dim fso As New FileSystemObject
   Dim mydrive As Drive
   Dim path As String
   'Initialize path.
   path = "C:\"
   'Get object.
   Set mydrive = fso.GetDrive(path)
   'Check for success.
   MsgBox prompt:="Drive= " & mydrive.DriveLetter, Buttons:=vbInformation + vbOKOnly, Title:="Drive"
Cleanup:
  Set fso = Nothing
  Set mydrive = Nothing
  Exit Sub
ErrorHandler:
  LogErrorToFile
  Resume Cleanup:
End Sub


