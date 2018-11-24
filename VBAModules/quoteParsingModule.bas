Attribute VB_Name = "quoteParsingModule"
Option Explicit
Public clientInfoArray()
Public parsedinfoArray()
Public quoteDate As String
Public measurementsArray()
Public opiArray()
Public ourPriceArray(1, 5)
Public extrasArray()
Public rowCountArray()
Public Price As String

'use FileDialogOpen popup to open a Quote to save
Sub browseQuote()
On Error GoTo errorhander
Dim oWord As Word.Application
Dim newPath As String
Dim QuoteToParse As String
  With Selection
    .WholeStory
    .Delete
  End With
  Const msoFileDialogOpen = 1
  Set oWord = CreateObject("Word.Application")
  If computerNameModule.ComputerName = "MIKE2" Then
      newPath = ThisDocument.Path & "\Estimating 2012"
      Else
      newPath = "M:\Estimating and Invoicing\Estimating and Invoicing 2012\Estimating 2012"
  End If

  oWord.ChangeFileOpenDirectory newPath
  'oWord.Width = 80
  With oWord.fileDialog(msoFileDialogOpen)
      .Title = "Select Quote Template"
      .AllowMultiSelect = False
  End With

  If oWord.fileDialog(msoFileDialogOpen).Show = -1 Then
      QuoteToParse = oWord.fileDialog(msoFileDialogOpen).SelectedItems.Item(1)
      Documents.Open (QuoteToParse)
      Selection.WholeStory
      Selection.Copy
      ActiveDocument.Close SaveChanges:=False
      
  'The activedocument was the Quote selected in the Dialog,
  'the activedocument is copied, then closed, then the quoteParsingUtility is activated
  'and the clip board containing the Quote selected is pasted into the Utility doc. Sweet,
      ThisDocument.Activate
      Selection.Paste
  End If
  
  ActiveDocument.Bookmarks("\StartOfDoc").Select
cleanup:
  oWord.Quit
  hiddenPopup.Hide
Exit Sub
ErrorHandler:
  MsgBox "An error occured processing 'browseQuote'."
  Resume cleanup
End Sub
 
 'this coverts the Quote to Text if it contains tables,
 Sub convertToText()
    'there are 4 tables in the Quote, as each one is delete then next becomes Tables(1)
    Dim I As Long
    'first save the date that the Quote was created, so that it is not lost during table conversion
    On Error Resume Next
    quoteDate = ActiveDocument.Tables(1).Cell(I, 1).Range.Text
    quoteDate = Trim(quoteDate)
    quoteDate = CleanString(quoteDate)

    'For I = 1 To 5
    'On Error Resume Next
        ThisDocument.Tables(1).Select
        Selection.Rows.convertToText Separator:=wdSeparateByTabs, NestedTables:=True
    'Next I
 End Sub

 'this converts the Quote to a table with one column,
 Sub convertToTable()
    Selection.WholeStory
    WordBasic.TextToTable ConvertFrom:=0, NumColumns:=1, _
        InitialColWidth:=wdAutoPosition, Format:=0, Apply:=1184, AutoFit:=0, _
        SetDefault:=0, Word8:=0, Style:="Table Grid"
    
    'delete empty rows from top of table
    'an empty row has 4 characters, so Len(row)=4 if empty
    Dim Row1 As Range
    Dim Row2 As Range
    Dim Row3 As Range
    Set Row1 = ThisDocument.Tables(1).Rows(1).Range
    Set Row2 = ThisDocument.Tables(1).Rows(2).Range
    Set Row3 = ThisDocument.Tables(1).Rows(3).Range
    Row1.Select
    If Len(Row1) = 4 Then
        Selection.Rows.Delete
        Row2.Select
        If Len(Row2) = 4 Then
            Selection.Rows.Delete
            Row3.Select
            If Len(Row3) = 4 Then Selection.Rows.Delete
        End If
    End If

    With Selection
      .Tables(1).ApplyStyleHeadingRows = Not Selection.Tables(1). _
        ApplyStyleHeadingRows
      .Tables(1).ApplyStyleRowBands = Not Selection.Tables(1). _
        ApplyStyleRowBands
      .Tables(1).ApplyStyleFirstColumn = Not Selection.Tables(1). _
        ApplyStyleFirstColumn
      .Borders(wdBorderTop).LineStyle = wdLineStyleNone
      .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
      .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
      .Borders(wdBorderRight).LineStyle = wdLineStyleNone
      .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
      .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
      .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    End With
    ActiveDocument.Save
    Selection.MoveUp Unit:=wdLine, Count:=1
 End Sub
 
 'parse clientInfo using combination of 'vbTab' and ':' to find two strings in one cell
 'need the labels, use Find to set flags and associate the flag with the correct string.
 Sub parseClientInfo()
    Dim I As Long
    Dim clientInfo As String
    Dim clientInfoLength As Long
    Dim colonPosition As Long
    Dim firstPart As String
    Dim secondPart As String
    Dim firstPartLength As Long
    Dim secondPartLength As Long
    Dim clientInfoTrim As Long
    Dim tabPosition As Long
    Dim foundLabel As String
    Dim countClientInfo As Long
    
    Dim labelArray(8)
    labelArray(1) = "Phone:"
    labelArray(2) = "Cell:"
    labelArray(3) = "Fax:"
    labelArray(4) = "Email:"
    labelArray(5) = "Re:"
    labelArray(6) = "Track #:"
    labelArray(7) = "Attn:"
    labelArray(8) = "Contact:"
    
    countClientInfo = countRowsModule.rowCountArray(1)
    ReDim clientInfoArray(countClientInfo)
    ReDim parsedinfoArray(countClientInfo, 2)

    For I = 1 To UBound(clientInfoArray)
        clientInfo = ActiveDocument.Tables(1).Cell(I, 1).Range.Text
        clientInfo = Trim(clientInfo)
        clientInfo = CleanString(clientInfo)
        clientInfoLength = Len(clientInfo)
        colonPosition = InStr(clientInfo, ":")
        tabPosition = InStr(clientInfo, vbTab)
        clientInfoArray(I) = Left(clientInfo, clientInfoLength)
        parsedinfoArray(I, 1) = clientInfoArray(I)
        'MsgBox colonPosition
       
       If colonPosition > 10 Then
            firstPart = clientInfo
            firstPartLength = tabPosition
            firstPart = Left(firstPart, firstPartLength - 1)
            firstPart = Trim(firstPart)
            parsedinfoArray(I, 1) = firstPart
            
            secondPart = clientInfo
            secondPartLength = clientInfoLength - colonPosition
            secondPart = Mid(secondPart, colonPosition + 1, secondPartLength)
            secondPart = Trim(secondPart)
            secondPart = CleanString(secondPart)
            
            'save the label associated with the secondPart
            'the label is immediate before the ':', 3-7 characters
            'use Case statement
            Dim J As Long
            For J = 1 To UBound(labelArray)
                foundLabel = InStr(clientInfo, labelArray(J))
                If foundLabel <> 0 Then
                    Select Case labelArray(J)
                        Case "Phone:": secondPart = "Phone: " & secondPart
                        Case "Cell:": secondPart = "Cell: " & secondPart
                        Case "Fax:": secondPart = "FAX: " & secondPart
                        Case "Email:": secondPart = "Email: " & secondPart
                        Case "Re:": secondPart = "Re: " & secondPart
                        Case "Track #:": secondPart = "Track #: " & secondPart
                        Case "Attn:": secondPart = "Phone: " & secondPart
                        Case "Contact:": secondPart = "Contact: " & secondPart
                    MsgBox prompt:=labelArray(J) & " " & secondPart, _
                    Buttons:=vbOKOnly + vbInformation, Title:="Parse Client Info"
                    End Select
                End If
            Next J
        'set array with concatenate label
        parsedinfoArray(I, 2) = Trim(secondPart)
        End If
    Next I
 End Sub
 
 'this code sets data arrays for each Quote section
' Sub getQuoteSections()
'    'Application.ScreenUpdating = False
'    countRowsModule.countSectionRows
'    Dim I As Long
'    Dim clientInfoArray(4, 2)
'    Dim ourPrice As String
'
'    'count items, redim for size
'    Dim countClientInfo As Long
'    Dim countMeasurements As Long
'    Dim countOPI As Long
'    Dim countOurPrice As Long
'    Dim countExtras As Long
'
'    countClientInfo = findNumberOfRowsModule.rowCountArray(1)
'    countMeasurements = findNumberOfRowsModule.rowCountArray(2)
'    countOPI = findNumberOfRowsModule.rowCountArray(3)
'    countOurPrice = findNumberOfRowsModule.rowCountArray(4)
'    countExtras = findNumberOfRowsModule.rowCountArray(5)
'
'    Dim measurementsStart As Long
'    measurementsStart = countClientInfo + 2
'    ReDim measurementsArray(countMeasurements)
'    For I = 1 To UBound(measurementsArray) - 2
'       measurementsArray(I) = Trim(CleanString(ThisDocument.Tables(1).Cell(measurementsStart + I, 1).Range.Text))
'    Next
'
'    Dim opiStart As Long
'    opiStart = countClientInfo + countMeasurements + 3
'    ReDim opiArray(countOPI)
'    For I = 1 To UBound(opiArray) - 2
'       opiArray(I) = Trim(CleanString(ThisDocument.Tables(1).Cell(opiStart + I, 1).Range.Text))
'    Next
'
'    'parsePrice
'    parsePrice
'
'    Dim extrasStart As Long
'    extrasStart = countClientInfo + countMeasurements + countOPI + 8
'    ReDim extrasArray(countExtras)
'    For I = 1 To UBound(extrasArray) - 3
'       extrasArray(I) = Trim(CleanString(ThisDocument.Tables(1).Cell(extrasStart + I, 1).Range.Text))
'    Next
'
''Clean-up and close
'    'TakeoffApp.SaveAs "C:\Users\mike\Documents\Takeoff.xlsm"
'    Application.DisplayAlerts = False
'        'QuoteSaverWorkbook.Close savechanges:=True
'        'oDoc.Close savechanges:=True
'    Application.DisplayAlerts = True
'
''Application.ScreenUpdating = True
'ActiveDocument.Bookmarks("\StartOfDoc").Select
'End Sub

Sub parsePrice()

    Dim ourPrice As String
    Dim Tax As String
    Dim priceWithTax As String
    Dim ourPriceRange As Range
    Dim ourPriceLength As Long
    Dim plusPosition As Long
    Dim colonPosition As Long
    Dim ourPriceLabel As String
    
    Selection.GoTo What:=wdGoToBookmark, Name:="OurPriceBM"
    Set ourPriceRange = Selection.Cells(1).Range
    ourPrice = Trim(CleanString(ourPriceRange.Text))
    ourPriceLength = Len(ourPrice)
    
    'if the Our Price: label and the price are both on the same line then parse them
    If ourPriceLength > 15 Then
        colonPosition = InStr(ourPrice, ":")
        Selection.MoveRight Unit:=wdCharacter, Count:=colonPosition, Extend:=wdExtend
        Selection.Delete Unit:=wdCharacter, Count:=1
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.Cells(1).Range.Text = "Our price:"
    End If
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    'if cell range does not contain the price then delete the row
    Do While Len(Selection.Cells(1).Range) < 6
        Selection.Rows.Delete
        'Selection.MoveDown Unit:=wdLine, Count:=1
    Loop
    
   ' Selection.Cells(1).Range.Text = Trim(CleanString(ourPrice))
    ourPrice = Trim(CleanString(Selection.Cells(1).Range.Text))
    plusPosition = InStr(ourPrice, "+")
    Price = Left(ourPrice, plusPosition - 1)
    Price = Trim(Price)

End Sub

