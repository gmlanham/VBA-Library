Attribute VB_Name = "QuickQuoteModule"
Option Explicit

'code source- http://www.freevbcode.com/ShowCode.asp?ID=43
Private Declare Function GetComputerName Lib "kernel32" _
Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As _
Long) As Long

Private Sub test()
    Debug.Print ComputerName
End Sub

Public Function ComputerName() As String
  Dim sBuffer As String
  
  Dim lAns As Long
 
  sBuffer = Space$(255)
  lAns = GetComputerName(sBuffer, 255)
  If lAns <> 0 Then
        'read from beginning of string to null-terminator
        ComputerName = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
   Else
        Err.Raise Err.LastDllError, , _
          "A system call returned an error code of " _
           & Err.LastDllError
   End If

End Function

Sub openQuoteUtility()
  Dim newPath As String
  If ComputerName = "MIKE2" Then
    newPath = "C:\Users\mike\Documents\My Projects\Project.QT"
    Else
    newPath = "M:\Estimating and Invoicing\Estimating and Invoicing 2011\Estimating 2011\QT"
  End If
  Documents.Open (newPath & "\QuoteUtility4.docm")
  Word.Application.Visible = True
End Sub
Sub openTakeoffUtility()
  Dim newPath As String
  If ComputerName = "MIKE2" Then
    newPath = "C:\Users\mike\Documents\My Projects\Project.QT"
    Else
    newPath = "M:\Estimating and Invoicing\Estimating and Invoicing 2011\Estimating 2011\QT"
  End If
  Workbooks.Open (newPath & "\TakeoffUtility4.xlsm")
  Excel.Application.Visible = True
  Run "BrowseButtonClick"
End Sub


