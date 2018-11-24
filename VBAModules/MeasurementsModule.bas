Attribute VB_Name = "MeasurementsModule"
Option Explicit
Option Base 1

Public Sub getTakeoff()
'shortcut: Ctrl+m
'run all the calculations necessary to determine cost for labor and materials based on specs.
    Application.Run "takeoffItems"          'ctrl+k
    Application.Run "getWallDescriptions"   'ctrl+a
    Application.Run "sortZeros"             'ctrl+o
    Application.Run "selectNonZeros"        'ctrl+p
    Application.Run "clearZeros"            'ctrl+e
    Application.Run "sortWallType"          'ctrl+s

    'Run "BuildQuoteModule.setMeasurements"
End Sub

Sub takeoffItems()
Attribute takeoffItems.VB_ProcData.VB_Invoke_Func = "k\n14"
'shortcut: Ctrl+k
'copy and paste the items specified for the job
    Sheets("TakeoffInput").Select
    Range("A4:B83").Select
    Selection.Copy

    Sheets("Measurements").Select
    Range("A2").Select
    ActiveSheet.Paste
End Sub

Sub getWallDescriptions()
Attribute getWallDescriptions.VB_ProcData.VB_Invoke_Func = "a\n14"
'shortcut: Ctrl+a
'copy and takeoff inputs to measurements calculation sheet
    Sheets("TakeoffInput").Select
    Range("Z4:Z26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Measurements").Select
    Range("B2").Select
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub sortZeros()
'separated the spec items from the items not being used by sorting on background color
'the background color is white for the unused items.
    ActiveWorkbook.Worksheets("Measurements").Sort.SortFields.clear
    ActiveWorkbook.Worksheets("Measurements").Sort.SortFields.Add Key:=Range( _
        "A2:A81"), SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Measurements").Sort
        .SetRange Range("A2:C81")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub selectNonZeros()
Attribute selectNonZeros.VB_ProcData.VB_Invoke_Func = "p\n14"
'shortcut: Ctrl+p
' select non-zero items using count
    Dim RC As Range, listCount
    listCount = Cells(1, 5).Value
    Set RC = Range(Cells(2, 1), Cells(listCount + 1, 4))
    RC.Select
    Selection.Copy
End Sub

Sub clearZeros()
Attribute clearZeros.VB_ProcData.VB_Invoke_Func = "e\n14"
'shortcut: Ctrl+e
'clean up the items list by getting rid of all teh unused (dimensions=0)items.
    Dim RC As Range, listCount
    listCount = Cells(1, 5).Value
    Set RC = Range(Cells(listCount + 2, 1), Cells(listCount + 71, 4))
    RC.Select
    Selection.clear
End Sub

Sub sortWallType()
Attribute sortWallType.VB_ProcData.VB_Invoke_Func = "s\n14"
'shortcut: Ctrl+s
'count the walls, then sorts them by wall type in a custom order.
    Dim RC As Range, listCount
    listCount = Cells(2, 5).Value
    Set RC = Range(Cells(2, 1), Cells(listCount + 1, 4))
    RC.Select
    ActiveWorkbook.Worksheets("Measurements").Sort.SortFields.clear
    ActiveWorkbook.Worksheets("Measurements").Sort.SortFields.Add Key:=Range( _
        "C2:C7"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "walkout,house,garage", DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Measurements").Sort
        .SetRange RC
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    listCount = Cells(1, 5).Value
    Set RC = Range(Cells(1, 1), Cells(listCount + 1, 2))
    RC.Select
    Selection.Copy
End Sub


