Attribute VB_Name = "categoriesModule"
Option Explicit
    Dim objNamespace As NameSpace
    Dim objCategories As Categories
    Dim objCategory As Category
    Dim strOutput As String
    Public standardColor(26)
 Sub assignStandardColors()
        standardColor(0) = Outlook.OlCategoryColor.olCategoryColorNone
        standardColor(1) = Outlook.OlCategoryColor.olCategoryColorRed
        standardColor(2) = Outlook.OlCategoryColor.olCategoryColorOrange
        standardColor(3) = Outlook.OlCategoryColor.olCategoryColorPeach
        standardColor(4) = Outlook.OlCategoryColor.olCategoryColorYellow
        standardColor(5) = Outlook.OlCategoryColor.olCategoryColorGreen
        standardColor(6) = Outlook.OlCategoryColor.olCategoryColorTeal
        standardColor(7) = Outlook.OlCategoryColor.olCategoryColorOlive
        standardColor(8) = Outlook.OlCategoryColor.olCategoryColorBlue
        standardColor(9) = Outlook.OlCategoryColor.olCategoryColorPurple
        standardColor(22) = Outlook.OlCategoryColor.olCategoryColorMaroon
        standardColor(23) = Outlook.OlCategoryColor.olCategoryColorSteel
        standardColor(24) = Outlook.OlCategoryColor.olCategoryColorGray
        standardColor(25) = Outlook.OlCategoryColor.olCategoryColorDarkGray
        standardColor(14) = Outlook.OlCategoryColor.olCategoryColorBlack
        standardColor(15) = Outlook.OlCategoryColor.olCategoryColorDarkRed
        standardColor(16) = Outlook.OlCategoryColor.olCategoryColorDarkOrange
        standardColor(17) = Outlook.OlCategoryColor.olCategoryColorDarkPeach
        standardColor(18) = Outlook.OlCategoryColor.olCategoryColorDarkYellow
        standardColor(19) = Outlook.OlCategoryColor.olCategoryColorDarkGreen
        standardColor(20) = Outlook.OlCategoryColor.olCategoryColorDarkTeal
        standardColor(21) = Outlook.OlCategoryColor.olCategoryColorDarkOlive
        standardColor(10) = Outlook.OlCategoryColor.olCategoryColorDarkBlue
        standardColor(11) = Outlook.OlCategoryColor.olCategoryColorDarkPurple
        standardColor(12) = Outlook.OlCategoryColor.olCategoryColorDarkMaroon
        standardColor(13) = Outlook.OlCategoryColor.olCategoryColorDarkSteel
 End Sub
'this works, it returns all Categories in the Master List
Sub categoriesList()
    Dim I As Integer
    I = 0
    ' Obtain a NameSpace object reference.
    Set objNamespace = Application.GetNamespace("MAPI")
    
    ' Check if the Categories collection for the Namespace
    ' contains one or more Category objects.
    If objNamespace.Categories.Count > 0 Then
        ' Enumerate the Categories collection.
        For Each objCategory In objNamespace.Categories
            taskAppointmentForm.categoriesComboBox.AddItem objCategory.Name
            I = I + 1
            'Debug.Print I, objNamespace.Categories.Item(I).Name
        Next
    End If
    
    ' Clean up.
    Set objCategory = Nothing
    Set objNamespace = Nothing
End Sub

'this code gets Categories
Sub categoriesList2()
    
    Dim I As Integer
    I = 0
    ' Obtain a NameSpace object reference.
    Set objNamespace = Application.GetNamespace("MAPI")
    
    ' Check if the Categories collection for the Namespace
    ' contains one or more Category objects.
    If objNamespace.Categories.Count > 0 Then
        ' Enumerate the Categories collection.
        For Each objCategory In objNamespace.Categories
            categoriesForm.categoriesListBox.AddItem objCategory.Name
            I = I + 1
            'Debug.Print I, objNamespace.Categories.Item(I).Name
        Next
    End If
    
    ' Clean up.
    Set objCategory = Nothing
    Set objNamespace = Nothing
End Sub

Sub EnumerateFoldersInStores()
 Dim colStores As Outlook.Stores
 Dim oStore As Outlook.Store
 Dim oRoot As Outlook.Folder
 
 On Error Resume Next
 Set colStores = Application.Session.Stores
 For Each oStore In colStores
    Set oRoot = oStore.GetRootFolder
    Debug.Print (oRoot.folderPath)
    EnumerateFolders oRoot
 Next
End Sub
 
Private Sub EnumerateFolders(ByVal oFolder As Outlook.Folder)
 Dim folders As Outlook.folders
 Dim Folder As Outlook.Folder
 Dim foldercount As Integer

 On Error Resume Next
 Set folders = oFolder.folders
 foldercount = folders.Count
 
 'Check if there are any folders below oFolder
 If foldercount Then
    For Each Folder In folders
    Debug.Print (Folder.folderPath)
    EnumerateFolders Folder
    Next
 End If
End Sub

Sub EnumerateCategoriesForStores()
 Dim oStores As Outlook.Stores
 Dim oStore As Outlook.Store
 Dim oCategories As Outlook.Categories
 Dim oCategory As Outlook.Category
 
 On Error Resume Next
 Set oStores = Application.Session.Stores
 For Each oStore In oStores
    Debug.Print oStore.DisplayName
    Debug.Print "--------------Categories-----------------"
    Set oCategories = oStore.Categories
    For Each oCategory In oCategories
        Debug.Print Chr(9) & oCategory.Name
    Next
    Debug.Print ""
 Next
 
End Sub
 
Sub standardCategories()
    Dim I As Integer
    Dim CategoriesArray()
    Dim categoriesCount As Integer
    I = 0
    'populate listboxes
    With categoriesForm.standardCategoriesListBox
        .AddItem "Meet Client"
        .AddItem "Survey"
        .AddItem "Planning"
        .AddItem "First Call"
        .AddItem "Excavation"
        .AddItem "Footing Stand "
        .AddItem "Footing Pour "
        .AddItem "Footing Strip "
        .AddItem "Wall Stand "
        .AddItem "Wall Pour "
        .AddItem "Wall Strip "
        .AddItem "Waterproofing"
        .AddItem "Cleanup"
        .AddItem "Backfill"
        .AddItem "Travel"
        .AddItem "Inspection"
        .AddItem "Rework"
        .AddItem "Deleted"
        .AddItem "Archived"
        .AddItem "Testing"
        .AddItem "Personal"
    End With
    categoriesCount = 22 'objNamespace.Categories.Count
    ReDim CategoriesArray(categoriesCount, 2)
    ' Obtain a NameSpace object reference.
    Set objNamespace = Application.GetNamespace("MAPI")
    
    ' Check if the Categories collection for the Namespace
    ' contains one or more Category objects.
    If objNamespace.Categories.Count > 0 Then
    
        ' Enumerate the Categories collection.
        For Each objCategory In objNamespace.Categories
            'taskAppointmentForm.categoriesComboBox.AddItem objCategory.Name
            I = I + 1
            On Error Resume Next
            CategoriesArray(I, 1) = objNamespace.Categories.Item(I).Name
            CategoriesArray(I, 2) = objNamespace.Categories.Item(I).Color
            'categoriesForm.standardCategoriesListBox.AddItem objNamespace.Categories.Item(I).Name
            'Debug.Print I, objNamespace.Categories.Item(I).Name, objNamespace.Categories.Item(I).Color
        Next
    End If
    
    ' Clean up.
    Set objCategory = Nothing
    Set objNamespace = Nothing
    
End Sub

'this works to show the Color Categories DialogBox
Sub categoriesChooser()
'http://msdn.microsoft.com/en-us/library/bb175161(v=office.12).aspx
    'Creates an appointment to access ShowCategoriesDialog
    Dim olApptItem As Outlook.AppointmentItem
    
    'Creates appointment item
    Set olApptItem = Application.CreateItem(olAppointmentItem)
    'olApptItem.Body = "Please meet with me regarding these sales figures."
    'olApptItem.Recipients.Add ("Mike Lanham")
    'olApptItem.Subject = "Sales Reports"
    'Display the appointment
    olApptItem.Display
    'Display the Show Categories dialog box
    olApptItem.ShowCategoriesDialog
    'Debug.Print olApptItem.Categories
End Sub

Sub getCategories()
    Set Session = CreateObject("Redemption.RDOSession")
    Session.Logon
    Set objNamespace.Categories = Session.Categories
    For Each objCategory In objNamespace.Categories
       Debug.Print objCategory.Name
    Next
End Sub

Sub setCategories()
    Set Session = CreateObject("Redemption.RDOSession")
    Session.Logon
    Set Store = Session.GetSharedMailbox("dmitry")
    Set Categories = Store.Categories
    Set Category = Categories.Add("Redemption Category", olCategoryColorPeach)

End Sub

Sub sessionObject()
Set Session = CreateObject("Redemption.RDOSession")
Session.MAPIOBJECT = Application.Session.MAPIOBJECT
Set Inbox = Session.GetDefaultFolder(olFolderInbox)
For Each Msg In Inbox.Items
  Debug.Print (Msg.Subject)
Next

End Sub

'this code sets categories on the selected Task
'http://mark.aufflick.com/blog/2006/10/27/setting-outlook-categories-with-vba
'the code on the website has an error, it uses Inspector in the Dim statement, rather than Explorer
'it works after that one change.
Public Sub addCategory()
    Dim objOutlook As Outlook.Application
    Dim objExplorer As Outlook.Explorer

    Dim strDateTime As String

    ' Instantiate an Outlook Application object.
    Set objOutlook = CreateObject("Outlook.Application")

    ' The ActiveExplorer is the currently open item.
    Set objExplorer = objOutlook.ActiveExplorer

    ' Check and see if anything is open.
    If Not objExplorer Is Nothing Then
        ' Get the current item.
        Dim arySelection As Object
        Set arySelection = objExplorer.Selection
        Dim X As Integer
        Dim strCats As String
        For X = 1 To arySelection.Count
            strCats = arySelection.Item(X).Categories
            If Not strCats = "" Then
                strCats = strCats & ","
            End If
            strCats = strCats & "Archived,Meet Client,Survey,Planning,First Call,Excavation,Footing Stand,Footing Pour,Footing Strip,Wall Stand,Wall Pour,Wall Strip,Waterproofing,Cleanup,Backfill,Inspection,Rework,Travel,Personal,Testing,Deleted"
            arySelection.Item(X).Categories = strCats
            arySelection.Item(X).Save
        Next X
        
    Else
        ' Show error message with only the OK button.
        MsgBox "No explorer is open", vbOKOnly
    End If

    ' Set all objects equal to Nothing to destroy them and
    ' release the memory and resources they take.
    Set objOutlook = Nothing
    Set objExplorer = Nothing
End Sub

Sub enumerateCategories()
    Dim objCategories As Outlook.Categories
    Dim objCategory As Outlook.Category
    
    Set objCategories = Application.Session.Categories
    For Each objCategory In objCategories
        Debug.Print objCategory.Name
    Next
    ' Set all objects equal to Nothing to destroy them and
    ' release the memory and resources they take.
    Set objCategories = Nothing
    Set objCategory = Nothing
End Sub

Sub addStandardCategories()
'this macro sets Standard Categories
'the code steps through the Categories and adds a Category if not found
'source: http://msdn.microsoft.com/en-us/library/ff424467.aspx
    assignStandardColors
    
    Dim objCategories As Outlook.Categories
    Dim objCategory As Outlook.Category
    Dim standardCategoriesArray(26, 2)
    'populate standard categories array
        standardCategoriesArray(1, 1) = "Meet Client"
        standardCategoriesArray(2, 1) = "Survey"
        standardCategoriesArray(3, 1) = "Planning"
        standardCategoriesArray(4, 1) = "First Call"
        standardCategoriesArray(5, 1) = "Excavation"
        standardCategoriesArray(6, 1) = "Footing Stand"
        standardCategoriesArray(7, 1) = "Footing Pour"
        standardCategoriesArray(8, 1) = "Footing Strip"
        standardCategoriesArray(9, 1) = "Wall Stand"
        standardCategoriesArray(10, 1) = "Wall Pour"
        standardCategoriesArray(11, 1) = "Wall Strip"
        standardCategoriesArray(12, 1) = "Waterproofing"
        standardCategoriesArray(13, 1) = "Cleanup"
        standardCategoriesArray(14, 1) = "Backfill"
        standardCategoriesArray(15, 1) = "Travel"
        standardCategoriesArray(16, 1) = "Inspection"
        standardCategoriesArray(17, 1) = "Rework"
        standardCategoriesArray(18, 1) = "Deleted"
        standardCategoriesArray(19, 1) = "Archived"
        standardCategoriesArray(20, 1) = "Testing"
        standardCategoriesArray(21, 1) = "Personal"
        standardCategoriesArray(22, 1) = "Category W"
        standardCategoriesArray(23, 1) = "Category X"
        standardCategoriesArray(24, 1) = "Category Y"
        standardCategoriesArray(25, 1) = "Category Z"
    Dim I As Integer
    'set the category colors using the standard color array
    For I = 1 To 25
        standardCategoriesArray(I, 2) = standardColor(I)
    Next I
    Set objCategories = Application.Session.Categories
    Dim foundMatch As Boolean
    For I = 1 To 25
        For Each objCategory In objCategories
            'skip add category if already exist
            If standardCategoriesArray(I, 1) = objCategory.Name Then
                foundMatch = True
                GoTo NextI
            End If
        Next
        On Error Resume Next
        objCategory = objCategories.Add(standardCategoriesArray(I, 1), standardCategoriesArray(I, 2), Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyNone)
        Debug.Print standardCategoriesArray(I, 1), standardCategoriesArray(I, 2)
NextI:
    Next I
ExitLoop:
    Set objCategories = Nothing
    Set objCategory = Nothing
    MsgBox "Standard Categories are set."
End Sub



















