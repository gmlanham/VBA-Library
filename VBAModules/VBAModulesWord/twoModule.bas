Attribute VB_Name = "twoModule"
Option Explicit

Private Sub test()
Dim countItems As Long
countItems = 1
  Debug.Print twoVar(countItems), countItems
End Sub

Public Function twoVar(Optional ByRef countItems As Long) As Boolean
  countItems = 3
  twoVar = True
End Function
