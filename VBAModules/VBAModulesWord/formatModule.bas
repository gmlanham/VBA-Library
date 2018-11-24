Attribute VB_Name = "formatModule"
Option Explicit
Sub setWordOptions()
  Word.Application.Options.CheckSpellingAsYouType = False
End Sub
Sub formatReport(ByVal documentName As String)
  Documents(documentName).Select
  Selection.Font.Name = "Arial"
  Dim lineNumber As Long
  Dim totalLines As Long
  totalLines = Word.Application.Selection.Paragraphs.Count
  Selection.GoTo 0  'What:=wdGoToBookmark, Name:=("\startofdoc")
  For lineNumber = 1 To totalLines
    With Selection
      .MoveDown unit:=wdLine, Count:=1
      .HomeKey unit:=wdLine
      .MoveRight unit:=wdWord, Count:=1, Extend:=wdExtend
      .MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
      If InStr(Trim(Selection), "'") Then
        .HomeKey unit:=wdLine
        .EndKey unit:=wdLine, Extend:=wdExtend
        .Font.Color = 39168
        .Font.Bold = True
      End If
      .HomeKey unit:=wdLine
      .EndKey unit:=wdLine, Extend:=wdExtend
      If InStr(Trim(Selection), "Sub") Or InStr(Trim(Selection), "Function") Then
        .HomeKey unit:=wdLine
        .EndKey unit:=wdLine, Extend:=wdExtend
        .Font.Bold = True
        .Font.Size = 12
      End If
    End With
  Next lineNumber
  Selection.GoTo What:=wdGoToBookmark, Name:=("\startofdoc")

End Sub
Sub formatSection()
  With Selection
    .ExtendMode = True
    .GoTo What:=wdGoToBookmark, Name:=("\startofdoc")
    With .ParagraphFormat
      .SpaceBefore = 0
      .SpaceBeforeAuto = False
      .SpaceAfter = 0
      .SpaceAfterAuto = False
      .LineSpacingRule = wdLineSpaceSingle
      .LineUnitBefore = 0
      .LineUnitAfter = 0
      .LeftIndent = InchesToPoints(0.25)
      With .TabStops
        .ClearAll
        .Add Position:=InchesToPoints(3)
      End With
    End With
    .GoTo What:=wdGoToBookmark, Name:=("\startofdoc")
    .TypeText vbCr
    ActiveDocument.Bookmarks.Add Name:="BM", Range:=.Range
    .Find.Text = "Procedures"
    .Find.Execute
    .HomeKey unit:=wdLine
    .EndKey unit:=wdLine, Extend:=wdExtend
    '.Style = ActiveDocument.Styles("Heading 2")
    .Font.Bold = True
    .ParagraphFormat.LeftIndent = InchesToPoints(0)
    .MoveRight unit:=wdCharacter, Count:=1
    '.MoveDown unit:=wdLine, Count:=1
    .EndKey unit:=wdLine, Extend:=wdExtend
    .Font.Underline = wdUnderlineThick
    .GoTo What:=wdGoToBookmark, Name:=("\startofdoc")
  End With
End Sub
