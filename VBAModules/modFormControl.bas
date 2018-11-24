Attribute VB_Name = "modFormControl"
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modFormControl
' By Chip Pearson, chip@cpearson.com, www.cpearson.com
' 21-March-2008
' URL: http://www.cpearson.com/Excel/FormControl.aspx
' Requires: modWindowCaption at http://www.cpearson.com/Excel/FileExtensions.aspx
'
' ----------------------------------
' Functions In This Module:
' ----------------------------------
'   SetFormParent
'       Sets a userform's parent to the Application or the ActiveWindow.
'   IsCloseButtonVisible
'       Returns True or False indicating whether the userform's Close button
'       is visible.
'   ShowCloseButton
'       Displays or hides the userform's Close button.
'   IsCloseButtonEnabled
'       Returns True or False indicating whether the userform's Close button
'       is enabled.
'   EnableCloseButton
'       Enables or disables a userform's Close button.
'   ShowTitleBar
'       Displays or hides a userform's Title Bar. The title bar cannot be
'       hidden if the form is resizable.
'   IsTitleBarVisible
'       Returns True or False indicating if the userform's Title Bar is visible.
'   MakeFormResizable
'       Makes the form resizable or not resizable. If the form is made resizable,
'       the title bar cannot be hidden.
'   IsFormResizable
'       Returns True or False indicating whether the userform is resizable.
'   SetFormOpacity
'       Sets the opacity of a form from fully opaque to fully invisible.
'   HasMaximizeButton
'       Returns True or False indicating whether the userform has a
'       maximize button.
'   HasMinimizeButton
'       Returns True or False indicating whether the userform has a
'       minimize button.
'   ShowMaximizeButton
'       Displays or hides a Maximize Window button on the userform.
'   ShowMinimizeButton
'       Displays or hides a Minimize Window button on the userform.
'   HWndOfUserForm
'       Returns the window handle (HWnd) of a userform.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const C_USERFORM_CLASSNAME = "ThunderDFrame"
Private Const C_EXCEL_APP_CLASSNAME = "XLMain"
Private Const C_EXCEL_DESK_CLASSNAME = "XLDesk"
Private Const C_EXCEL_WINDOW_CLASSNAME = "Excel7"
Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000
Private Const MF_ENABLED = &H0&
Private Const MF_DISABLED = &H2&
Private Const MF_GRAYED = &H1&
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const GWL_HWNDPARENT = (-8)
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Private Const C_ALPHA_FULL_TRANSPARENT As Byte = 0
Private Const C_ALPHA_FULL_OPAQUE As Byte = 255
Private Const WS_DLGFRAME = &H400000
Private Const WS_THICKFRAME = &H40000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000

Public Enum FORM_PARENT_WINDOW_TYPE
    FORM_PARENT_NONE = 0
    FORM_PARENT_APPLICATION = 1
    FORM_PARENT_WINDOW = 2
End Enum

Private Declare Function SetParent Lib "user32" ( _
    ByVal hWndChild As Long, _
    ByVal hWndNewParent As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal crey As Byte, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function DrawMenuBar Lib "user32" ( _
    ByVal hwnd As Long) As Long

Private Declare Function GetMenuItemCount Lib "user32" ( _
    ByVal hMenu As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal bRevert As Long) As Long
    
Private Declare Function RemoveMenu Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long
    
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long
    
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" ( _
    ByVal hwnd As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Private Declare Function EnableMenuItem Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal wIDEnableItem As Long, _
    ByVal wEnable As Long) As Long


Function ShowMaximizeButton(UF As MSForms.UserForm, _
    HideButton As Boolean) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShowMaximizeButton
' Displays (if HideButton is False) or hides (if HideButton is True)
' a maximize window button.
' NOTE: If EITHER a Minimize or Maximize button is displayed,
' BOTH buttons are visible but may be disabled.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim R As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    ShowMaximizeButton = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
If HideButton = False Then
    WinInfo = WinInfo Or WS_MAXIMIZEBOX
Else
    WinInfo = WinInfo And (Not WS_MAXIMIZEBOX)
End If
R = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)

ShowMaximizeButton = (R <> 0)

End Function

Function ShowMinimizeButton(UF As MSForms.UserForm, _
    HideButton As Boolean) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShowMinimizeButton
' Displays (if HideButton is False) or hides (if HideButton is True)
' a minimize window button.
' NOTE: If EITHER a Minimize or Maximize button is displayed,
' BOTH buttons are visible but may be disabled.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim R As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    ShowMinimizeButton = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
If HideButton = False Then
    WinInfo = WinInfo Or WS_MINIMIZEBOX
Else
    WinInfo = WinInfo And (Not WS_MINIMIZEBOX)
End If
R = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)

ShowMinimizeButton = (R <> 0)

End Function

Function HasMinimizeButton(UF As MSForms.UserForm) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HasMinimizeButton
' Returns True if the userform has a minimize button, False
' otherwise.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim R As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    HasMinimizeButton = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

If WinInfo And WS_MINIMIZEBOX Then
    HasMinimizeButton = True
Else
    HasMinimizeButton = False
End If

End Function

Function HasMaximizeButton(UF As MSForms.UserForm) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HasMaximizeButton
' Returns True if the userform has a maximize button, False
' otherwise.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim R As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    HasMaximizeButton = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

If WinInfo And WS_MAXIMIZEBOX Then
    HasMaximizeButton = True
Else
    HasMaximizeButton = False
End If

End Function


Function SetFormParent(UF As MSForms.UserForm, _
    Parent As FORM_PARENT_WINDOW_TYPE) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetFormParent
' Set the UserForm UF as a child of (1) the Application, (2) the
' Excel ActiveWindow, or (3) no parent. Returns TRUE if successful
' or FALSE if unsuccessful.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WindHWnd As Long
Dim R As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    SetFormParent = False
    Exit Function
End If

Select Case Parent
    Case FORM_PARENT_APPLICATION
        R = SetParent(UFHWnd, Application.hwnd)
    Case FORM_PARENT_NONE
        R = SetParent(UFHWnd, 0&)
    Case FORM_PARENT_WINDOW
        If Application.ActiveWindow Is Nothing Then
            SetFormParent = False
            Exit Function
        End If
        WindHWnd = WindowHWnd(Application.ActiveWindow)
        If WindHWnd = 0 Then
            SetFormParent = False
            Exit Function
        End If
        R = SetParent(UFHWnd, WindHWnd)
    Case Else
        SetFormParent = False
        Exit Function
End Select
SetFormParent = (R <> 0)

End Function


Function IsCloseButtonVisible(UF As MSForms.UserForm) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsCloseButtonVisible
' Returns TRUE if UserForm UF has a close button, FALSE if there
' is no close button.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim R As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    IsCloseButtonVisible = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
IsCloseButtonVisible = (WinInfo And WS_SYSMENU)

End Function

Function ShowCloseButton(UF As MSForms.UserForm, HideButton As Boolean) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShowCloseButton
' This displays (if HideButton is FALSE) or hides (if HideButton is
' TRUE) the Close button on the userform
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim R As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
If HideButton = False Then
    ' set the SysMenu bit
    WinInfo = WinInfo Or WS_SYSMENU
Else
    ' clear the SysMenu bit
    WinInfo = WinInfo And (Not WS_SYSMENU)
End If

R = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
ShowCloseButton = (R <> 0)

End Function


Function IsCloseButtonEnabled(UF As MSForms.UserForm) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsCloseButtonEnabled
' This returns TRUE if the close button is enabled or FALSE if
' the close button is disabled.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim hMenu As Long
Dim ItemCount As Long
Dim PrevState As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    IsCloseButtonEnabled = False
    Exit Function
End If
' Get the menu handle
hMenu = GetSystemMenu(UFHWnd, 0&)
If hMenu = 0 Then
    IsCloseButtonEnabled = False
    Exit Function
End If

ItemCount = GetMenuItemCount(hMenu)
' Disable the button. This returns MF_DISABLED or MF_ENABLED indicating
' the previous state of the item.
PrevState = EnableMenuItem(hMenu, ItemCount - 1, MF_DISABLED Or MF_BYPOSITION)

If PrevState = MF_DISABLED Then
    IsCloseButtonEnabled = False
Else
    IsCloseButtonEnabled = True
End If
' restore the previous state
EnableCloseButton UF, (PrevState = MF_DISABLED)

DrawMenuBar UFHWnd

End Function


Function EnableCloseButton(UF As MSForms.UserForm, Disable As Boolean) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EnableCloseButton
' This function enables (if Disable is False) or disables (if
' Disable is True) the "X" button on a UserForm UF.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim hMenu As Long
Dim ItemCount As Long
Dim Res As Long

' Get the HWnd of the UserForm.
UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    EnableCloseButton = False
    Exit Function
End If
' Get the menu handle
hMenu = GetSystemMenu(UFHWnd, 0&)
If hMenu = 0 Then
    EnableCloseButton = False
    Exit Function
End If

ItemCount = GetMenuItemCount(hMenu)
If Disable = True Then
    Res = EnableMenuItem(hMenu, ItemCount - 1, MF_DISABLED Or MF_BYPOSITION)
Else
    Res = EnableMenuItem(hMenu, ItemCount - 1, MF_ENABLED Or MF_BYPOSITION)
End If
If Res = -1 Then
    EnableCloseButton = False
    Exit Function
End If
DrawMenuBar UFHWnd

EnableCloseButton = True


End Function

Function ShowTitleBar(UF As MSForms.UserForm, HideTitle As Boolean) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShowTitleBar
' Displays (if HideTitle is FALSE) or hides (if HideTitle is TRUE) the
' title bar of the userform UF.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim R As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    ShowTitleBar = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

If HideTitle = False Then
    ' turn on the Caption bit
    WinInfo = WinInfo Or WS_CAPTION
Else
    ' turn off the Caption bit
    WinInfo = WinInfo And (Not WS_CAPTION)
End If
R = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
ShowTitleBar = (R <> 0)
End Function

Function IsTitleBarVisible(UF As MSForms.UserForm) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsTitleBarVisible
' Returns TRUE if the title bar of UF is visible or FALSE if the
' title bar is not visible.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim UFHWnd As Long
Dim WinInfo As Long
Dim R As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    IsTitleBarVisible = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

IsTitleBarVisible = (WinInfo And WS_CAPTION)

End Function

Function MakeFormResizable(UF As MSForms.UserForm, Sizable As Boolean) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MakeFormResizable
' This makes the userform UF resizable (if Sizable is TRUE) or not
' resizable (if Sizalbe is FALSE). Returns TRUE if successful or FALSE
' if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim UFHWnd As Long
Dim WinInfo As Long
Dim R As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    MakeFormResizable = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
If Sizable = True Then
    WinInfo = WinInfo Or WS_SIZEBOX
Else
    WinInfo = WinInfo And (Not WS_SIZEBOX)
End If

R = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
MakeFormResizable = (R <> 0)


End Function

Function IsFormResizable(UF As MSForms.UserForm) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsFormResizable
' Returns TRUE if UF is resizable, FALSE if UF is not resizable.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim R As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    IsFormResizable = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

IsFormResizable = (WinInfo And WS_SIZEBOX)

End Function


Function SetFormOpacity(UF As MSForms.UserForm, Opacity As Byte) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetFormOpacity
' This function sets the opacity of the UserForm referenced by the
' UF parameter. Opacity specifies the opacity of the form, from
' 0 = fully transparent (invisible) to 255 = fully opaque. The function
' returns True if successful or False if an error occurred. This
' requires Windows 2000 or later -- it will not work in Windows
' 95, 98, or ME.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim UFHWnd As Long
Dim WinL As Long
Dim Res As Long

SetFormOpacity = False

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    Exit Function
End If

WinL = GetWindowLong(UFHWnd, GWL_EXSTYLE)
If WinL = 0 Then
    Exit Function
End If

Res = SetWindowLong(UFHWnd, GWL_EXSTYLE, WinL Or WS_EX_LAYERED)
If Res = 0 Then
    Exit Function
End If

Res = SetLayeredWindowAttributes(UFHWnd, 0, Opacity, LWA_ALPHA)
If Res = 0 Then
    Exit Function
End If

SetFormOpacity = True

End Function


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


Function ClearBit(Value As Long, ByVal BitNumber As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ClearBit
' Clears the specified bit in Value and returns the result. Bits are
' numbered, right (most significant) 31 to left (least significant) 0.
' BitNumber is made positive and then MOD 32 to get a valid bit number.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim SetMask As Long
Dim ClearMask As Long

BitNumber = Abs(BitNumber) Mod 32

SetMask = Value
If BitNumber < 30 Then
    ClearMask = Not (2 ^ (BitNumber - 1))
    ClearBit = SetMask And ClearMask
Else
    ClearBit = Value And &H7FFFFFFF
End If

End Function

