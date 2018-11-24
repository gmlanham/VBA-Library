Attribute VB_Name = "DeclarationModule"
Option Explicit
  ' Declare variables to access the macros
  Public VBAEditor As VBIDE.VBE
  Public VBProj As VBIDE.VBProject
  Public VBComp As VBIDE.VBComponent
  Public CodeMod As VBIDE.CodeModule
  
  Public OWord As Word.Application
  Public ODoc As Document
  Public FileToGet As String
  Public newPath As String
  Public OExcel As Excel.Application
  Public OWorkbooks As Excel.Workbooks
  Public OWB As Workbook
  
  Public Enum ProcScope
    ScopePrivate = 1
    ScopePublic = 2
    ScopeFriend = 3
    ScopeDefault = 4
  End Enum
  
  Public Enum LineSplits
    LineSplitRemove = 0
    LineSplitKeep = 1
    LineSplitConvert = 2
  End Enum
  
  Public Type ProcInfo
    ProcName As String
    ProcKind As VBIDE.vbext_ProcKind
    ProcStartLine As Long
    ProcBodyLine As Long
    ProcCountLines As Long
    ProcScope As ProcScope
    ProcDeclaration As String
  End Type
  
  Public excelFlag As Boolean

Public Sub setPublicVariables()
  Set OWord = Word.Application
  Set ODoc = OWord.Documents("VBAUtility.docm")
  newPath = ODoc.path
End Sub
