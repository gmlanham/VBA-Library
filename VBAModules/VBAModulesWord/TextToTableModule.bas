Attribute VB_Name = "TextToTableModule"
Option Explicit

 'this converts the Quote to a table with one column,
 Sub convertToTable()
    Selection.WholeStory
    WordBasic.TextToTable ConvertFrom:=0, NumColumns:=1, _
        InitialColWidth:=wdAutoPosition, Format:=0, Apply:=1184, AutoFit:=0, _
        SetDefault:=0, Word8:=0, Style:="Table Grid"
DeleteEmptyRows:
'    'delete empty rows from top of table
'    'an empty row has 4 characters, so Len(row)=4 if empty
'    Dim Row1 As Range
'    Dim Row2 As Range
'    Dim Row3 As Range
'    Set Row1 = ThisDocument.Tables(1).Rows(1).Range
'    Set Row2 = ThisDocument.Tables(1).Rows(2).Range
'    Set Row3 = ThisDocument.Tables(1).Rows(3).Range
'    Row1.Select
'    If Len(Row1) = 4 Then
'        Selection.Rows.Delete
'        Row2.Select
'        If Len(Row2) = 4 Then
'            Selection.Rows.Delete
'            Row3.Select
'            If Len(Row3) = 4 Then Selection.Rows.Delete
'        End If
'    End If
FormatTable:
'    Selection.Tables(1).ApplyStyleHeadingRows = Not Selection.Tables(1). _
'        ApplyStyleHeadingRows
'    Selection.Tables(1).ApplyStyleRowBands = Not Selection.Tables(1). _
'        ApplyStyleRowBands
'    Selection.Tables(1).ApplyStyleFirstColumn = Not Selection.Tables(1). _
'        ApplyStyleFirstColumn
'    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
'    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
'    Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
'    Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
'    Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
'    ActiveDocument.Save
'    Selection.MoveUp Unit:=wdLine, Count:=1
 End Sub


