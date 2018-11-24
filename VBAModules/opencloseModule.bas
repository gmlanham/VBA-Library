Attribute VB_Name = "opencloseModule"
Option Explicit
Sub showControlPanel()
  ControlPanel.Show
End Sub

Private Sub ListBm()
  EnumerateDocBkMrks (ThisDocument.Path & "/" & "QuoteUtility3.docm")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : EnumerateDocBkMrks
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Generate a listing of all the Bookmarks containing within the
'             specified word document and print them to the immediate window.
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2010-Sep-10                 Initial Release
'---------------------------------------------------------------------------------------
Function EnumerateDocBkMrks(sFileName As String)
On Error GoTo Error_Handler
'Requires a reference to the Word object library
Dim oApp                As Word.Application
Dim oDoc                As Word.Document
Dim dBkMrk              As Bookmark
 
On Error Resume Next
    Set oApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then 'Word isn't running so start it
        Set oApp = CreateObject("Word.Application")
    End If
On Error GoTo 0
 
    Set oDoc = oApp.Documents.Open(sFileName)
    oApp.Visible = True 'Control whether or not Word becomes
                         'visible to the user
    
    'Loop through each form field
    For Each dBkMrk In oDoc.Range.Bookmarks
        Debug.Print dBkMrk.Name
    Next
 
Error_Handler_Exit:
'    On Error Resume Next
'    oDoc.Close False
'    oApp.Quit
'    Set oDoc = Nothing
'    Set oApp = Nothing
'    Exit Function
 
Error_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
            "     Error Number: " & Err.Number & vbCrLf & _
            "     Error Source: EnumerateDocBkMrks" & vbCrLf & _
            "     Error Description: " & Err.Description, _
            vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function


