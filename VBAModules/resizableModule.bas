Attribute VB_Name = "resizableModule"
Option Explicit
'Reference Windows Forms 2.0 Object Library
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const C_USERFORM_CLASSNAME = "ThunderDFrame"

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

Sub testResizable()
  resizableForm.Show
End Sub
Sub MakeFormResizable(UF As MSForms.UserForm)
'This makes the userform UF resizable
Dim UFHWnd As Long
Dim WinInfo As Long
Dim R As Long

UFHWnd = resizableModule.HWndOfUserForm(UF)

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
WinInfo = WinInfo Or WS_SIZEBOX

R = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)


End Sub

Function HWndOfUserForm(UF As MSForms.UserForm) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HWndOfUserForm
' This returns the window handle (HWnd) of the userform referenced
' by UF. It first looks for a top-level window, then a child
' of the Application window, then a child of the ActiveWindow.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim AppHWnd As Long
Dim DeskHWnd As Long
Dim WinHWnd As Long
Dim UFHWnd As Long
Dim Cap As String
Dim WindowCap As String

Cap = UF.Caption

' First, look in top level windows
UFHWnd = FindWindow(C_USERFORM_CLASSNAME, Cap)
If UFHWnd <> 0 Then
    HWndOfUserForm = UFHWnd
    Exit Function
End If
' Not a top level window. Search for child of application.
AppHWnd = Application.hwnd
UFHWnd = FindWindowEx(AppHWnd, 0&, C_USERFORM_CLASSNAME, Cap)
If UFHWnd <> 0 Then
    HWndOfUserForm = UFHWnd
    Exit Function
End If
' Not a child of the application.
' Search for child of ActiveWindow (Excel's ActiveWindow, not
' Window's ActiveWindow).
If Application.ActiveWindow Is Nothing Then
    HWndOfUserForm = 0
    Exit Function
End If
WinHWnd = WindowHWnd(Application.ActiveWindow)
UFHWnd = FindWindowEx(WinHWnd, 0&, C_USERFORM_CLASSNAME, Cap)
HWndOfUserForm = UFHWnd

End Function


Function WindowHWnd(W As Excel.Window) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WindowHWnd
' This returns the HWnd of the Window referenced by W.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim AppHWnd As Long
Dim DeskHWnd As Long
Dim WHWnd As Long
Dim Cap As String

AppHWnd = Application.hwnd
DeskHWnd = FindWindowEx(AppHWnd, 0&, C_EXCEL_DESK_CLASSNAME, vbNullString)
If DeskHWnd > 0 Then
    Cap = WindowCaption(W)
    WHWnd = FindWindowEx(DeskHWnd, 0&, C_EXCEL_WINDOW_CLASSNAME, Cap)
End If
WindowHWnd = WHWnd

End Function
