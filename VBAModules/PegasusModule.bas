Attribute VB_Name = "PegasusModule"
Option Explicit
Public Sub populatePegasusDivisions()
Dim itemCount As Long
itemCount = 0
  insertPegasusDivisions
  Call setPegasusDivision("ExcavationPegasus")
  Call setPegasusDivision("WaterPegasus", itemCount)
  formatWaterPegasus (itemCount)
  Call setPegasusDivision("SpreadEagle")
  formatSpreadEagle
  Call setPegasusDivision("NewDivision")
  formatNewDivision
  Call setPegasusDivision("Materials", itemCount)
  Call formatMaterialsPegasus(itemCount)
  AbbeyModule.PopulateAbbeyDivisions
End Sub
Private Sub testWaterPegasus()
  formatWaterPegasus (1)
End Sub
Public Sub formatSpreadEagle()
  selectSectionsModule.selectSection ("spreadeagle")
    If Len(Selection) < 5 Then
    With Selection
      .GoTo What:=wdGoToBookmark, Name:="spreadEagleBM"
      .MoveDown Unit:=wdLine, Count:=3, Extend:=wdExtend
      .Delete
    End With
    Exit Sub
  End If

  With Selection.Find
    .Text = " @ /yd"
    .Replacement.Text = ""
    .Execute Replace:=wdReplaceAll
    .Text = "40"
    .Replacement.Text = "supply and install 40"
    .Execute Replace:=wdReplaceAll
  End With
End Sub
Public Sub formatNewDivision()
  'selectSectionsModule.selectNewDivision
'    With Selection
'    .GoTo what:=wdGoToBookmark, Name:="newDivisionBM"
'    .MoveDown unit:=wdLine, Count:=1
'    .Extend
'    .GoTo what:=wdGoToBookmark, Name:="labourBM"
'    .MoveLeft unit:=wdCharacter, Count:=1
'    Debug.Print Len(Selection)
'    .ExtendMode = False
'  End With
  selectSectionsModule.selectSection ("newdivision")
  If Len(Selection) < 5 Then
    With Selection
      .GoTo What:=wdGoToBookmark, Name:="newDivisionBM"
      .MoveDown Unit:=wdLine, Count:=3, Extend:=wdExtend
      .Delete
    End With
    Exit Sub
  End If

  With Selection.Find
    .Text = "New Division"
    .Replacement.Text = divisionTitle
    .Execute Replace:=wdReplaceAll
  End With

  With Selection.Find
    .Text = " @ /hr"
    .Replacement.Text = ""
    .Execute Replace:=wdReplaceAll
  End With
End Sub
Public Sub formatWaterPegasus(Optional ByVal itemCount As Long)
  Dim atFound As Boolean
  atFound = False
  selectSectionsModule.selectSection ("waterpegasus")
  If Len(Selection) < 5 Then
    With Selection
      .GoTo What:=wdGoToBookmark, Name:="waterPegasusBM"
      .MoveDown Unit:=wdLine, Count:=3, Extend:=wdExtend
      .Delete
      WaterPegasusFound = False
    End With
    Exit Sub
  End If
  With Selection.Find
    .Text = " @ "
    .Replacement.Text = ""
    .Execute Replace:=wdReplaceAll
    .Text = "1 "
    .Replacement.Text = ""
    .Execute Replace:=wdReplaceAll
  End With
  With Selection
    .GoTo What:=wdGoToBookmark, Name:="waterPegasusBM"
    .MoveDown Unit:=wdLine, Count:=2
    .TypeText "("
    .MoveLeft Unit:=wdCharacter, Count:=1
    .MoveDown Unit:=wdLine, Count:=3, Extend:=wdExtend
    .ParagraphFormat.TabStops.ClearAll
    .Font.Size = 9
    With .Find
      .Text = vbTab
      .Replacement.Text = " "
      .Execute Replace:=wdReplaceAll
    End With
    .MoveRight Unit:=wdCharacter, Count:=1
    .MoveUp Unit:=wdLine, Count:=1
    .TypeBackspace
    .TypeText ". "
    .MoveDown Unit:=wdLine, Count:=1
    .TypeText ".)" & vbCr
    .GoTo What:=wdGoToBookmark, Name:="waterPegasusBM"
    .MoveLeft Unit:=wdCharacter, Count:=1
    .TypeText vbCr
  End With
  
End Sub
Private Sub testMaterials()
  formatMaterialsPegasus (6)
End Sub
Public Sub formatMaterialsPegasus(Optional ByVal itemCount As Long)
  selectSectionsModule.selectMaterials
  With Selection
    .Font.Underline = wdUnderlineNone
    
    If itemCount <= 2 Then
      .MoveRight Unit:=wdCharacter, Count:=1
      .MoveUp Unit:=wdLine, Count:=1
      .TypeBackspace
      .GoTo What:=wdGoToBookmark, Name:="materialsBM"
      .MoveDown Unit:=wdLine, Count:=1
      .TypeBackspace
      .GoTo What:=wdGoToBookmark, Name:="materialsBM"
      .MoveDown Unit:=wdLine, Count:=1
      .TypeBackspace
      .ParagraphFormat.TabStops(InchesToPoints(4)).Position = _
        InchesToPoints(3.4)
    Else
      .ParagraphFormat.TabStops.Add Position:=InchesToPoints(0.25), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
      .GoTo What:=wdGoToBookmark, Name:="materialsBM"
      .EndKey Unit:=wdLine, Extend:=wdExtend
      .TypeText "Materials Required:"
    End If
  End With
End Sub
Public Sub insertPegasusDivisions()
  With Selection
    .GoTo What:=wdGoToBookmark, Name:="labourBM"
    .MoveUp Unit:=wdLine, Count:=1
    .TypeText vbCr & "Excavation:" & vbCr
    .MoveUp Unit:=wdLine, Count:=1
    ActiveDocument.Bookmarks.Add Range:=.Range, Name:="ExcavationPegasusBM"
    ExcavationPegasusFound = True
  End With
  With Selection
    .GoTo What:=wdGoToBookmark, Name:="labourBM"
    .MoveUp Unit:=wdLine, Count:=1
    .TypeText vbCr & "Water/Sewer/Storm Service:" & vbCr
    .MoveUp Unit:=wdLine, Count:=1
    ActiveDocument.Bookmarks.Add Range:=.Range, Name:="waterPegasusBM"
  End With
  With Selection
    .GoTo What:=wdGoToBookmark, Name:="labourBM"
    .MoveUp Unit:=wdLine, Count:=1
    .TypeText vbCr & "Spread:" & vbCr
    .MoveUp Unit:=wdLine, Count:=1
    ActiveDocument.Bookmarks.Add Range:=.Range, Name:="spreadEagleBM"
  End With
  With Selection
    .GoTo What:=wdGoToBookmark, Name:="labourBM"
    .MoveUp Unit:=wdLine, Count:=1
    .TypeText vbCr & "New Division:" & vbCr
    .MoveUp Unit:=wdLine, Count:=1
    ActiveDocument.Bookmarks.Add Range:=.Range, Name:="newDivisionBM"
  End With
  selectSectionsModule.selectMaterials
  With Selection
    .Delete
    .GoTo What:=wdGoToBookmark, Name:="materialsBM"
    .EndKey Unit:=wdLine, Extend:=wdExtend
    .TypeText "Concrete Required:"
  End With
End Sub

Sub setPegasusDivision(ByVal divisionName As String, Optional ByRef itemCount As Long, Optional ByRef divisionTitle As String)  'ByRef returns changed value for this variable
  On Error GoTo ErrorHandler:
  Const procedureName = "setPegasusDivision"
  Dim startTimer As Single
  Dim endTimer As Single
  Dim statusbarString As String
  endTimer = Timer
  Dim I As Long
  Dim J As Long
  'Dim itemCount As Long
  Dim divisionArray()
  startTimer = Timer - StartTime
  J = 0
  'the setSectionArray function returns a data array for the section specified by the arguement, "Waterproofing"
  If InStr(divisionName, "Materials") <> 0 Then
    divisionArray = TakeoffDataSetModule.setSectionArrayMaterials
  Else
    divisionArray = TakeoffDataSetModule.setSectionArray(divisionName)
  End If
  itemCount = UBound(divisionArray)
  With Selection
    If itemCount = 0 Then
      .GoTo What:=wdGoToBookmark, Name:=divisionName & "BM"
      .EndKey Unit:=wdLine, Extend:=wdExtend
      .Delete
      ActiveDocument.Bookmark.Delete Range:=.Range, Name:=divisionName & "BM"
    Else
      .GoTo What:=wdGoToBookmark, Name:=divisionName & "BM"
      .MoveDown Unit:=wdLine, Count:=1
      '.MoveRight unit:=wdcharater, Count:=1
      For I = 1 To itemCount
        If Trim(divisionArray(I)) <> vbNullString Then
          .TypeText Trim(divisionArray(I)) & vbCr
        'If I < itemCount Then Selection.TypeText vbCr
        End If
      Next I
    End If
  End With
  
Cleanup:
  endTimer = Timer - StartTime
  statusbarString = "Finished " & procedureName & ": " & divisionName & "... " & vbTab & _
    "Section Timer= " & Format(endTimer - startTimer, "#.0") & " seconds, " & vbTab & _
    "Total Time= " & Format(endTimer / 60, "#.0") & " minutes"
  Application.StatusBar = statusbarString
  Debug.Print statusbarString
  startTimer = endTimer
  Exit Sub
ErrorHandler:
  Debug.Print "An error was thrown by " & procedureName & _
  vbCr & Err.Number & ": " & Err.Description
  Resume Cleanup:
End Sub




