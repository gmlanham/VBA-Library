Attribute VB_Name = "twoModule"
Option Explicit
'this demonstrates getting two values returned from a function,
'one value is returned with the function name
'a second value is returned an Optional ByRef arguement
'note that the arguement is changed in the function, so the the new value is acceble to the calling sub
Private Sub test()
Dim countItems As Long
countItems = 1
  Debug.Print twoVar(countItems), countItems
End Sub

Public Function twoVar(Optional ByRef countItems As Long) As Boolean
  countItems = 3
  twoVar = True
End Function
