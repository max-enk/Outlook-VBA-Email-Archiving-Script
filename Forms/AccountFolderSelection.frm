VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AccountFolderSelection 
   Caption         =   "Mail Folder Selection"
   ClientHeight    =   6840
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11208
   OleObjectBlob   =   "AccountFolderSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AccountFolderSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Private variables
'' Retainment period
Private RetainDays As Collection                ' Value of the retain period in days
Private RetainLabels As Collection              ' Value of the retain period in user-readable form
Private RetainIndex As Integer                  ' Index of the retain period
'' Account-realted Variables
Private AccountName As String                   ' Account name
Private AccountFolders As Collection            ' Account folders
Private AccountFolderSettings As Collection     ' Account folder settings


' Initialize the UserForm
Private Sub UserForm_Initialize()
    ' Variable Declaration
    '' Loop Variables
    Dim i As Integer                            ' Loop counter for iterating through collections
    Dim j As Integer                            ' Loop counter for iterating through collections
    '' UserForm Objects
    Dim CheckFolder As MSForms.CheckBox         ' CheckBox for folder selection
    Dim LabelFolder As MSForms.Label            ' Label for folder name
    Dim LabelLastArchived As MSForms.Label      ' TextBox for last archived date
    Dim ComboBoxRetain As MSForms.ComboBox      ' ComboBox for retain period
    '' UserForm Variables
    Dim AllChecked As Boolean                   ' Check if all found accounts have been priorly archived
    '' Object Placements
    Dim TopOffset As Integer                    ' Offset for control placement
    Dim FrameBottom As Integer                  ' Bottom position of the last control item

    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Define Account-related variables
    AccountName = Account.GetAccountName()
    Set AccountFolders = Account.GetAccountFolders()
    Set AccountFolderSettings = Account.GetAccountFolderSettings()
    
    ' Controls
    Me.FolderFrame.Controls.Clear               ' Clear prior frame settings
    TopOffset = 5                               ' Set the initial TopOffset for frame items
    AllChecked = True                           ' Control variable for initial state of CheckAll
    
    ' Initialize RetainDays and RetainLabels collections
    Set RetainDays = New Collection
    Set RetainLabels = New Collection
    ' Fill RetainDays with the date value
    RetainDays.Add 0
    RetainDays.Add 7
    RetainDays.Add 14
    RetainDays.Add 21
    RetainDays.Add 30
    RetainDays.Add 60
    RetainDays.Add 90
    RetainDays.Add 180
    RetainDays.Add 270
    RetainDays.Add 365
    RetainDays.Add 730
    RetainDays.Add 1095
    RetainDays.Add -1
    ' Fill RetainLabels with the date
    RetainLabels.Add "nothing"
    RetainLabels.Add "1 Week"
    RetainLabels.Add "2 Weeks"
    RetainLabels.Add "3 Weeks"
    RetainLabels.Add "1 Month"
    RetainLabels.Add "2 Months"
    RetainLabels.Add "3 Months"
    RetainLabels.Add "6 Months"
    RetainLabels.Add "9 Months"
    RetainLabels.Add "1 Year"
    RetainLabels.Add "2 Years"
    RetainLabels.Add "3 Years"
    RetainLabels.Add "everything"
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Populate the FolderFrame with dynamic controls for folders
    For i = 1 To AccountFolders.Count
        ' Create a checkbox for folder selection
        Set CheckFolder = Me.FolderFrame.Controls.Add("Forms.CheckBox.1", "CheckFolder" & i)
        CheckFolder.Caption = ""   ' No text in checkbox itself
        CheckFolder.Top = TopOffset
        CheckFolder.Left = 10
        If AccountFolderSettings(i)(1) = "never" Then
            CheckFolder.value = False           ' Folder has never been archived
        Else
            CheckFolder.value = True            ' Folder has been archived: preselect it
        End If
        
        ' Create a label for the folder path
        Set LabelFolder = Me.FolderFrame.Controls.Add("Forms.Label.1", "LabelFolder" & i)
        LabelFolder.Caption = Replace(AccountFolders(i), "\\" & AccountName, "", vbTextCompare)
        LabelFolder.Top = TopOffset
        LabelFolder.Left = 30
        LabelFolder.Width = 270
        
        ' Create a textbox for the last archived date
        Set LabelLastArchived = Me.FolderFrame.Controls.Add("Forms.Label.1", "LabelLastArchived" & i)
        LabelLastArchived.Caption = AccountFolderSettings(i)(1)
        LabelLastArchived.Top = TopOffset
        LabelLastArchived.Left = 300
        LabelLastArchived.Width = 100
        
        ' Create a combobox for retain period
        Set ComboBoxRetain = Me.FolderFrame.Controls.Add("Forms.ComboBox.1", "ComboBoxRetain" & i)
        ComboBoxRetain.Top = TopOffset
        ComboBoxRetain.Left = 400
        ComboBoxRetain.Width = 100
        
        ' Populate the combobox with retain period options
        For j = 1 To RetainLabels.Count
            ComboBoxRetain.AddItem RetainLabels(j)
        Next j
        
        ' Set the initial choice of retainment period
        RetainIndex = GetIndex(RetainDays, AccountFolderSettings(i)(2))
        If RetainIndex >= 0 Then
            ComboBoxRetain.ListIndex = RetainIndex
        Else
            ' Set to 6 months
            ComboBoxRetain.ListIndex = 7
            LogDebug "No valid retain interval for folder " & AccountFolders(i) & " found. Setting value to 6 months."
        End If
        
        
        ' Update control variable for CheckAll
        If Not CheckFolder.value Then
            AllChecked = False
        End If
        
        ' Set the TopOffset for the next set of controls
        TopOffset = TopOffset + 25
    Next i
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Static Headers
    ' Form Label
    LabelFolderSelection.Caption = "Select folders of account " & AccountName & " to be archived:"
    ' Checkmark
    CheckAll.Left = 32
    CheckAll.Width = 100
    If AllChecked Then
        CheckAll.value = True
    Else
        CheckAll.value = False
    End If
    ' Folder Header
    HeaderFolder.Caption = " Folder"
    HeaderFolder.Left = 50
    HeaderFolder.Width = 270
    HeaderFolder.Height = 20
    HeaderFolder.Font.Bold = True
    HeaderFolder.BackColor = RGB(200, 200, 200)
    ' Date Header
    HeaderDate.Caption = " Last Archived"
    HeaderDate.Left = 320
    HeaderDate.Width = 100
    HeaderDate.Height = 20
    HeaderDate.Font.Bold = True
    HeaderDate.BackColor = RGB(200, 200, 200)
    ' Retain Header
    HeaderRetain.Caption = " Retain"
    HeaderRetain.Left = 420
    HeaderRetain.Width = 100
    HeaderRetain.Height = 20
    HeaderRetain.Font.Bold = True
    HeaderRetain.BackColor = RGB(200, 200, 200)
    

    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Final control adjustments
    FrameBottom = TopOffset                 ' Bottom position of the last control item

    ' Adjust the ScrollHeight of the FolderFrame to accommodate all items if needed
    If FrameBottom > Me.FolderFrame.Height Then
        ' Include vertical Scrollbar
        With Me.FolderFrame
            .Scrollbars = fmScrollBarsVertical
        End With
        
        ' Adjust scrollheight
        Me.FolderFrame.ScrollHeight = FrameBottom
    End If
End Sub



' Event handler for the Continue button
Private Sub ButtonContinue_Click()
    ' Variable Declaration
    '' Loop Variables
    Dim i As Integer                            ' Loop counter for iterating through collections
    '' UserForm Objects
    Dim CheckFolder As MSForms.CheckBox         ' Check box for account
    Dim ComboBoxRetain As MSForms.ComboBox      ' ComboBox for retain period
    '' Account-realted Variables
    Dim ArchiveFolders As Collection            ' Archive folders
    Dim ArchiveFolderSettings As Collection     ' Archive folder settings
    Dim FolderSettings As Collection            ' Settings for a single folder
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Build SelectedAccounts collection
    '' Clear previous selections
    Set ArchiveFolders = New Collection
    Set ArchiveFolderSettings = New Collection
    
    '' Loop through dynamically added checkboxes in the FolderFrame
    For i = 1 To AccountFolders.Count
        Set CheckFolder = Me.FolderFrame.Controls("CheckFolder" & i)
        Set ComboBoxRetain = Me.FolderFrame.Controls("ComboBoxRetain" & i)
        
        If CheckFolder.value = True Then
            ' Get the index of the selected retain period
            RetainIndex = ComboBoxRetain.ListIndex
            
            ' Add the corresponding retain days value to FolderSettings only if valid
            If RetainIndex >= 0 Then
                ' Create a new sub-collection for each folder setting
                Set FolderSettings = New Collection
                FolderSettings.Add CurrentDate                  ' Add the current date
                FolderSettings.Add RetainDays(RetainIndex + 1)  ' Add the corresponding retain days value
                
                ' Add settings
                ArchiveFolderSettings.Add FolderSettings
                
                ' Add the selected folder to the collection
                ArchiveFolders.Add AccountFolders(i)
            Else
                ' Skip the folder
                MsgBox "No valid retain interval for folder " & AccountFolders(i) & " found. Skipping archiving this folder.", vbExclamation
            End If
        End If
    Next i
    
    '' Ensure at least one folder is selected
    If ArchiveFolders.Count = 0 Then
        MsgBox "Please select at least one folder.", vbExclamation
        Exit Sub
    End If
    
    ' Add Archive folders and settings to the account object
    Account.SetArchiveFolders = ArchiveFolders
    Account.SetArchiveFolderSettings = ArchiveFolderSettings
    
    '' Hide the UserForm
    Me.Hide
End Sub



' Event handler for the Cancel button
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Check if the form is being closed by the red "X" button
    If CloseMode = vbFormControlMenu Then
        ' Set Continue to False
        Continue = False
    End If
End Sub



' Event handler for the Cancel button
Private Sub ButtonCancel_Click()
    ' Set continue to false
    Continue = False
    
    ' Hide the UserForm
    Me.Hide
End Sub



' Event handler for the CheckAll checkbox
Private Sub CheckAll_Click()
    Dim i As Integer
    Dim CheckFolder As MSForms.CheckBox
    Dim ShouldCheck As Boolean
    
    ' Determine the desired state based on the CheckAll checkbox
    ShouldCheck = CheckAll.value
    
    ' Loop through dynamically added checkboxes in the AccountFrame
    For i = 1 To AccountFolders.Count
        ' Reference each checkbox by its new name
        Set CheckFolder = Me.FolderFrame.Controls("CheckFolder" & i)
        If Not CheckFolder Is Nothing Then
            CheckFolder.value = ShouldCheck
        End If
    Next i
    
    ' Update UserForm
    Me.Repaint
End Sub





