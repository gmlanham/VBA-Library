Attribute VB_Name = "countRowsModule"
Option Explicit
Public rowCountArray(8)

'this code used the bookmarks to extend selections to blocks of text between two bookmarks
Sub countSectionRows()
'Application.ScreenUpdating = False
    ThisDocument.Tables(1).Cell(1, 1).Select
    On Error GoTo Err_Handler
    rowCountClientInfo
    rowCountMeasurements
    rowCountOPI
    rowCountOurPrice
    rowCountExtras
    rowCountFoundation
    rowCountExcavation
    rowCountSeasonal
    ThisDocument.Tables(1).Cell(1, 1).Select
    'showRowCounts
'Application.ScreenUpdating = True

Exit Sub
Err_Handler:
    'Application.ScreenUpdating = True
        MsgBox prompt:="Sorry, the standard sections are not present.", _
        Buttons:=vbOKOnly + vbExclamation, Title:="Count Section Rows"

ActiveDocument.Bookmarks("\StartOfDoc").Select
End Sub

'extend selection to Excavation Section using bookmarks to select the rows
Sub selectExcavation()
    'select top row of section and exted range to next bookmark,
    With Selection
        .GoTo What:=wdGoToBookmark, Name:="ExcavationBM"
        .MoveDown Unit:=wdLine, Count:=1
        .Extend
        .GoTo What:=wdGoToBookmark, Name:="FoundationBM"
        .MoveLeft Unit:=wdCharacter, Count:=1
        .ExtendMode = False
    End With
End Sub

'extend selection to entire section using bookmarks to count the rows
Sub rowCountClientInfo()
    'select top row of section and exted range to bookmark, count rows in section
    Selection.GoTo What:=wdGoToBookmark, Name:="quoteDateBM"
    Selection.Extend
    Selection.GoTo What:=wdGoToBookmark, Name:="MeasurementsBM"
    Selection.MoveUp Unit:=wdLine, Count:=1
    rowCountArray(1) = Selection.Rows.Count
    Selection.ExtendMode = False
End Sub

'extend selection to entire section using bookmarks to count the rows
Sub rowCountMeasurements()
    'select top row of section and exted range to bookmark, count rows in section
    Selection.GoTo What:=wdGoToBookmark, Name:="MeasurementsBM"
    Selection.Extend
    Selection.GoTo What:=wdGoToBookmark, Name:="OPIBM"
    Selection.MoveUp Unit:=wdLine, Count:=1
    rowCountArray(2) = Selection.Rows.Count
    Selection.ExtendMode = False
End Sub

'extend selection to entire section using bookmarks to count the rows
Sub rowCountOPI()
    'select top row of section and exted range to bookmark, count rows in section
    Selection.GoTo What:=wdGoToBookmark, Name:="OPIBM"
    Selection.Extend
    Selection.GoTo What:=wdGoToBookmark, Name:="OurPriceBM"
    Selection.MoveUp Unit:=wdLine, Count:=1
    rowCountArray(3) = Selection.Rows.Count
    Selection.ExtendMode = False
End Sub

'extend selection to entire section using bookmarks to count the rows
Sub rowCountOurPrice()
    'select top row of section and exted range to bookmark, count rows in section
    Selection.GoTo What:=wdGoToBookmark, Name:="OurPriceBM"
    Selection.Extend
    Selection.GoTo What:=wdGoToBookmark, Name:="ExtrasBM"
    Selection.MoveUp Unit:=wdLine, Count:=1
    rowCountArray(4) = Selection.Rows.Count
    Selection.ExtendMode = False
End Sub

'extend selection to entire section using bookmarks to count the rows
Sub rowCountExtras()
    'select top row of section and exted range to bookmark, count rows in section
    Selection.GoTo What:=wdGoToBookmark, Name:="ExtrasBM"
    Selection.Extend
    Selection.GoTo What:=wdGoToBookmark, Name:="endBM"
    Selection.MoveUp Unit:=wdLine, Count:=1
    rowCountArray(5) = Selection.Rows.Count
    Selection.ExtendMode = False
End Sub

'extend selection to entire section using bookmarks to count the rows
Sub rowCountFoundation()
    'select top row of section and exted range to bookmark, count rows in section
    Dim FoundationFound As Boolean
    Dim excavationFound As Boolean
    Dim SeasonalFound As Boolean
    
    ThisDocument.Tables(1).Select
    Selection.Find.Text = "Foundation upgrades"
    FoundationFound = Selection.Find.Execute
    
    ThisDocument.Tables(1).Select
    Selection.Find.Text = "Excavation and Backfill upgrades"
    excavationFound = Selection.Find.Execute
    
    ThisDocument.Tables(1).Select
    Selection.Find.Text = "Seasonal upgrades"
    SeasonalFound = Selection.Find.Execute
    
    If FoundationFound = True And excavationFound = True Then
        Selection.GoTo What:=wdGoToBookmark, Name:="FoundationBM"
        Selection.Extend
        Selection.GoTo What:=wdGoToBookmark, Name:="ExcavationBM"
        Selection.MoveUp Unit:=wdLine, Count:=1
        rowCountArray(6) = Selection.Rows.Count
        Selection.ExtendMode = False
    End If
    
    If FoundationFound = True And SeasonalFound = True Then
        Selection.GoTo What:=wdGoToBookmark, Name:="FoundationBM"
        Selection.Extend
        Selection.GoTo What:=wdGoToBookmark, Name:="SeasonalBM"
        Selection.MoveUp Unit:=wdLine, Count:=1
        rowCountArray(6) = Selection.Rows.Count
        Selection.ExtendMode = False
    End If
    
End Sub

'extend selection to entire section using bookmarks to count the rows
Sub rowCountExcavation()
    'select top row of section and exted range to bookmark, count rows in section
    Dim excavationFound As Boolean
    Dim SeasonalFound As Boolean
    
    ThisDocument.Tables(1).Select
    Selection.Find.Text = "Excavation and Backfill upgrades"
    excavationFound = Selection.Find.Execute
    
    ThisDocument.Tables(1).Select
    Selection.Find.Text = "Seasonal upgrades"
    SeasonalFound = Selection.Find.Execute
    
    If excavationFound = True And SeasonalFound = True Then
        Selection.GoTo What:=wdGoToBookmark, Name:="ExcavationBM"
        Selection.Extend
        Selection.GoTo What:=wdGoToBookmark, Name:="SeasonalBM"
        Selection.MoveUp Unit:=wdLine, Count:=1
        rowCountArray(7) = Selection.Rows.Count
        Selection.ExtendMode = False
    End If
    If excavationFound = True And SeasonalFound = False Then
        Selection.GoTo What:=wdGoToBookmark, Name:="ExcavationBM"
        Selection.Extend
        Selection.GoTo What:=wdGoToBookmark, Name:="endBM"
        Selection.MoveUp Unit:=wdLine, Count:=1
        rowCountArray(7) = Selection.Rows.Count
        Selection.ExtendMode = False
    End If
    
    
End Sub

'extend selection to entire section using bookmarks to count the rows
Sub rowCountSeasonal()
    'select top row of section and exted range to bookmark, count rows in section
    Dim SeasonalFound As Boolean
    
    ThisDocument.Tables(1).Select
    Selection.Find.Text = "Seasonal upgrades"
    SeasonalFound = Selection.Find.Execute
    If SeasonalFound = True Then
        Selection.GoTo What:=wdGoToBookmark, Name:="SeasonalBM"
        Selection.Extend
        Selection.GoTo What:=wdGoToBookmark, Name:="endBM"
        Selection.MoveUp Unit:=wdLine, Count:=1
        rowCountArray(8) = Selection.Rows.Count
    Selection.ExtendMode = False
    End If
End Sub

Sub showRowCounts()
MsgBox prompt:="Row Counts: " & _
    vbCr & "clientInfo: " & rowCountArray(1) & _
    vbCr & "Measurements: " & rowCountArray(2) & _
    vbCr & "OPI: " & rowCountArray(3) & _
    vbCr & "OurPrice: " & rowCountArray(4) & _
    vbCr & "Extras: " & rowCountArray(5) & _
    vbCr & "Foundation: " & rowCountArray(6) & _
    vbCr & "Excavation: " & rowCountArray(7) & _
    vbCr & "Seasonal: " & rowCountArray(8), Buttons:=vbOKOnly + vbInformation, Title:="Section Array Row Counts"
End Sub





