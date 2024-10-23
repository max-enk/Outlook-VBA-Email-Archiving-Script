VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FolderSelection 
   Caption         =   "Archive Folder Selection"
   ClientHeight    =   4740
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8784.001
   OleObjectBlob   =   "FolderSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FolderSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Initialization of the FolderSelection UserForm
Private BrowseHandlers As Collection



Private Sub UserForm_Initialize()
    ' Variable Declaration
    '' Loop Variables
    Dim i As Integer                            ' Loop counter for iterating through collections
    '' Userform Objects
    Dim LabelAccount As MSForms.Label           ' Label for account name
    Dim TextBoxPath As MSForms.TextBox          ' Label for account path
    Dim ButtonBrowse As MSForms.CommandButton   ' Browse button for account
    Dim handler As BrowseButtonHandler          ' Browse button handler
    '' Object Placements
    Dim TopOffset As Integer                    ' AccountFrame item offset
    Dim FrameBottom As Integer                  ' Bottom of AccountFrame
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Controls
    Me.FolderFrame.Controls.Clear               ' Clear prior frame settings
    TopOffset = 5                               ' Set the initial TopOffset for frame items
    Set BrowseHandlers = New Collection         ' Initialize the collection of handlers



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Populate the FolderFrame with dynamic controls for accounts
    For i = 1 To SelectedAccounts.Count
        ' Create a label for the account name
        Set LabelAccount = Me.FolderFrame.Controls.Add("Forms.Label.1", "LabelAccountName" & i)
        LabelAccount.Caption = SelectedAccounts(i).GetAccountName()
        LabelAccount.Top = TopOffset
        LabelAccount.Left = 10
        LabelAccount.Width = 150

        ' Create a textbox for the archive path
        Set TextBoxPath = Me.FolderFrame.Controls.Add("Forms.TextBox.1", "TextBoxPath" & i)
        TextBoxPath.Text = SelectedAccounts(i).GetArchivePath()
        TextBoxPath.Top = TopOffset
        TextBoxPath.Left = 150
        TextBoxPath.Width = 170

        ' Create a button for browsing folders
        Set ButtonBrowse = Me.FolderFrame.Controls.Add("Forms.CommandButton.1", "ButtonBrowse" & i)
        ButtonBrowse.Caption = "Browse"
        ButtonBrowse.Top = TopOffset
        ButtonBrowse.Left = 330
        ButtonBrowse.Width = 50
        ButtonBrowse.Height = 18
        
        ' Create a new handler for this button
        Set handler = New BrowseButtonHandler
        Set handler.BrowseButton = ButtonBrowse
        Set handler.TextBoxPath = TextBoxPath

        ' Add the handler to the collection
        BrowseHandlers.Add handler
        
        ' Set the TopOffset for the next set of controls
        TopOffset = TopOffset + 25
    Next i
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Static Headers
    '' Account Header
    HeaderAccount.Left = 20
    HeaderAccount.Width = 150
    HeaderAccount.Height = 100
    HeaderAccount.Font.Bold = True
    HeaderAccount.BackColor = RGB(200, 200, 200)
    '' Path Header
    HeaderPath.Left = 170
    HeaderPath.Width = 250
    HeaderPath.Height = 100
    HeaderPath.Font.Bold = True
    HeaderPath.BackColor = RGB(200, 200, 200)

    
    
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
    '' UserForm Objects
    Dim TextBoxPath As MSForms.TextBox      ' Label for account path
    '' Paths
    Dim TextBoxValue As String              ' Account Path in text box
    Dim AccountPath As String               ' Account Path in class object
    '' Flags
    Dim AllValid As Boolean                 ' Checks if all paths in checkbox are valid
    '' Other
    Dim AccountName As String               ' Account Name in class object
    Dim AccountFolderName As String         ' Accoount file name in class object
    
    
    ' Setup
    AllValid = True
    

    ' Check for every account
    For i = 1 To SelectedAccounts.Count
        ' Get values from account config
        AccountName = SelectedAccounts(i).GetAccountName()
        AccountFolderName = SelectedAccounts(i).GetFileName()
        AccountPath = SelectedAccounts(i).GetArchivePath()
        
        
        ' Get value from textbox
        Set TextBoxPath = BrowseHandlers(i).TextBoxPath
        TextBoxValue = TextBoxPath.Text
        
        
        ' Check if the path is valid
        If Not IsValidPath(TextBoxValue) Then
            ' If conversion fails, show a message and reset the value
            MsgBox "Please enter a valid path for account " & AccountName & ".", vbExclamation, "Invalid Input."
            TextBoxPath.Text = AccountPath  ' Reset to previous value
            AllValid = False
            
            Exit For
        End If
        
        
        ' Check if the path ends in "\"
        If Right(TextBoxValue, Len("\")) <> "\" Then
            TextBoxValue = TextBoxValue & "\"
        End If
        
        
        ' Check if the path ends in the account file name
        If Right(TextBoxValue, Len(AccountFolderName & "\")) <> AccountFolderName & "\" Then
            TextBoxValue = TextBoxValue & AccountFolderName & "\"
        End If
        
        
        ' Update text box value
        TextBoxPath.Text = TextBoxValue
        

        ' Check if archive path has changed
        If TextBoxPath <> AccountPath Then
            ' Update account archive path
            SelectedAccounts(i).ChangeArchiveLocation TextBoxValue
        End If
    Next i
    
    If AllValid Then
        ' Hide the UserForm
        Me.Hide
    End If
End Sub



' Event handler for the Cancel button
Private Sub ButtonCancel_Click()
    ' Set continue to false
    Continue = False
    
    ' Hide the UserForm
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



' Event handler to set focus on the continue button
Private Sub UserForm_Activate()
    ' set focus
    ButtonContinue.SetFocus
End Sub



