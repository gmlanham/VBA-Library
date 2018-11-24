Attribute VB_Name = "RibbonControl"
Option Explicit
Public myRibbon As IRibbonUI
Sub Onload(ribbon As IRibbonUI)
'Create a ribbon instance for use in this project
Set myRibbon = ribbon
End Sub
'Callback for DropDown onAction
Sub myDDMacro(ByVal control As IRibbonControl, selectedID As String, _
               selectedIndex As Integer)
Select Case selectedIndex
  Case 1
    Macros.Macro1
  Case 2
    Macros.Macro2
  Case 3
    Macros.Macro3
End Select
'Force the ribbon to restore the control to original state
myRibbon.InvalidateControl control.id
End Sub
'Callback for DropDown GetItemCount
Sub GetItemCount(ByVal control As IRibbonControl, ByRef count)
count = 4
End Sub
'Callback for DropDown GetItemLabel
Sub GetItemLabel(ByVal control As IRibbonControl, _
    Index As Integer, ByRef label)
label = Choose(Index + 1, "Select from list", "Macro 1", "Macro 2", "Macro 3")
End Sub
'Callback DropDown GetSelectedIndex
Sub GetSelectedItemIndex(ByVal control As IRibbonControl, ByRef Index)
'This procedure is used to ensure the first item in the dropdown is displayed.
Select Case control.id
  Case Is = "DD1"
    Index = 0
  Case Else
End Select
End Sub
'Callback for Button onAction
Sub MyBtnMacro(ByVal control As IRibbonControl)
Select Case control.id
  Case Is = "Btn1"
    Macros.ShowEditor
  Case Is = "Btn1"
    'Do something else if a control id "Btn2" existed.
End Select
End Sub
'Callback for Toogle onAction
Sub ToggleonAction(control As IRibbonControl, pressed As Boolean)
Select Case control.id
  Case Is = "TB1"
    ActiveWindow.View.ShowBookmarks = Not ActiveWindow.View.ShowBookmarks
  Case Is = "TB2"
  'Note:  "pressed" returns the toggle state.  So we could use this instead.
    If pressed Then
      ActiveWindow.View.ShowHiddenText = False
    Else
      ActiveWindow.View.ShowHiddenText = True
    End If
    If Not ActiveWindow.View.ShowHiddenText Then
      ActiveWindow.View.ShowAll = False
    End If
End Select
'Force the ribbon to redefine the control wiht correct image and label
myRibbon.InvalidateControl (control.id)
End Sub
'Callback for togglebutton getLabel
Sub getLabel(control As IRibbonControl, ByRef returnedVal)
Select Case control.id
  Case Is = "TB1"
    If Not ActiveWindow.View.ShowBookmarks Then
      returnedVal = "Show Bookmarks"
    Else
      returnedVal = "Hide Bookmarks"
    End If
  Case Is = "TB2"
    If Not ActiveWindow.View.ShowHiddenText Then
      returnedVal = "Show Text"
    Else
      returnedVal = "Hide Text"
    End If
End Select
End Sub
'Callback for togglebutton getImage
Sub GetImage(control As IRibbonControl, ByRef returnedVal)
Select Case control.id
  Case Is = "TB1"
   If ActiveWindow.View.ShowBookmarks Then
      returnedVal = "_3DTiltRightClassic"
    Else
      returnedVal = "_3DTiltLeftClassic"
   End If
  Case Is = "TB2"
    If ActiveWindow.View.ShowHiddenText Then
      returnedVal = "WebControlHidden"
    Else
      returnedVal = "SlideShowInAWindow"
    End If
End Select
End Sub
'Callback for togglebutton getPressed
Sub buttonPressed(control As IRibbonControl, ByRef toggleState)
'toggleState is used tp set the toggle state (i.e., true or false) and determine how the
'toggle appears on the ribbon (i.e., flusn or sunken).
Select Case control.id
  Case Is = "TB1"
    If Not ActiveWindow.View.ShowBookmarks Then
      toggleState = True
    Else
      toggleState = False
    End If
  Case Is = "TB2"
    If Not ActiveWindow.View.ShowHiddenText Then
      toggleState = True
    Else
      toggleState = False
    End If
End Select
End Sub

