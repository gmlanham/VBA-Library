Attribute VB_Name = "modTest"
Option Explicit
Option Compare Text

Sub AAA_ShowTheForm()
UserForm1.Show vbModal

End Sub


Sub ExportModules()

Dim N As Long
Dim FolderName As String
N = Application.InputBox("Enter 1 to export 'modFormControl'." & vbCrLf & _
    "Enter 2 to export 'modWindowCaption'." & vbCrLf & _
    "Enter 3 to export both modules." & vbCrLf & _
    "Enter any other number to quit without export.", "Export Modules", 3, , , , , Type:=1)

    FolderName = BrowseFolder("Select A Folder To Which The Files Will Be Exported.")
    
    

End Sub



