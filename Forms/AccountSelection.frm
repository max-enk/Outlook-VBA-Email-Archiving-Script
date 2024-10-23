VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AccountSelection 
   Caption         =   "Account Selection"
   ClientHeight    =   4740
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8784.001
   OleObjectBlob   =   "AccountSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AccountSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Initialize the UserForm
Private Sub UserForm_Initialize()
    ' Variable Declaration
    '' Loop Variables
    Dim i As Integer                        ' Loop counter for iterating through collections
    Dim j As Integer                        ' Loop counter for iterating through collections
    '' UserForm Objects
    Dim CheckAccount As MSForms.CheckBox    ' Check box for account
    Dim LabelAccount As MSForms.Label       ' Label for account name
    Dim LabelDate As MSForms.Label          ' Label for archive date
    '' UserForm Variables
    Dim AllChecked As Boolean               ' Check if all found accounts have been priorly archived
    '' Object placements
    Dim TopOffset As Integer                ' AccountFrame item offset
    Dim FrameBottom As Integer              ' Bottom of AccountFrame
    


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Controls
    Me.AccountFrame.Controls.Clear          ' Clear prior frame settings
    TopOffset = 5                           ' Set the initial TopOffset for frame items
    AllChecked = True                       ' Control variable for initial state of CheckAll



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Populate the AccountFrame with dynamic controls for accounts
    For i = 1 To Accounts.Count
        ' Create a checkbox for the account
        Set CheckAccount = Me.AccountFrame.Controls.Add("Forms.CheckBox.1", "CheckAccountName" & i)
        CheckAccount.Caption = ""   ' No text in checkbox itself
        CheckAccount.Top = TopOffset
        CheckAccount.Left = 10
        If Accounts(i).GetExecutionDate() = "never" Then
            CheckAccount.value = False          ' Account has never been archived
        Else
            CheckAccount.value = True           ' Account has been archived: preselect it
        End If

        ' Create a label for the account name
        Set LabelAccount = Me.AccountFrame.Controls.Add("Forms.Label.1", "LabelAccountName" & i)
        LabelAccount.Caption = Accounts(i).GetAccountName()
        LabelAccount.Top = TopOffset
        LabelAccount.Left = 30
        LabelAccount.Width = 250

        ' Create a label for the date
        Set LabelDate = Me.AccountFrame.Controls.Add("Forms.Label.1", "LabelDateArchived" & i)
        LabelDate.Caption = Accounts(i).GetExecutionDate()
        LabelDate.Top = TopOffset
        LabelDate.Left = 280
        
        ' Update control variable for CheckAll
        If Not CheckAccount.value Then
            AllChecked = False
        End If

        ' Increment TopOffset for the next account entry
        TopOffset = TopOffset + 20
    Next i
        
        
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Static Headers
    '' Checkmark
    CheckAll.Left = 32
    CheckAll.Width = 100
    If AllChecked Then
        CheckAll.value = True
    Else
        CheckAll.value = False
    End If
    '' Account Header
    HeaderAccount.Left = 52
    HeaderAccount.Width = 250
    HeaderAccount.Height = 20
    HeaderAccount.Font.Bold = True
    HeaderAccount.BackColor = RGB(200, 200, 200)
    '' Date Header
    HeaderDate.Left = 302
    HeaderDate.Width = 120
    HeaderDate.Height = 20
    HeaderDate.Font.Bold = True
    HeaderDate.BackColor = RGB(200, 200, 200)



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Final Frame control adjustments
    FrameBottom = TopOffset                 ' Bottom position of the last control item

    '' Adjust the ScrollHeight of the AccountFrame to accommodate all items if needed
    If FrameBottom > Me.AccountFrame.Height Then
        ' Include vertical Scrollbar
        With Me.AccountFrame
            .Scrollbars = fmScrollBarsVertical
        End With
        
        ' Adjust scrollheight
        Me.AccountFrame.ScrollHeight = FrameBottom
    End If
End Sub



' Event handler for the Continue button
Private Sub ButtonContinue_Click()
    ' Variable Declaration
    '' Loop Variables
    Dim i As Integer                                ' Loop counter for iterating through collections
    '' UserForm Objects
    Dim CheckAccount As MSForms.CheckBox            ' Check box for account
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Build SelectedAccounts collection
    '' Clear previous selections
    Set SelectedAccounts = New Collection
    
    '' Loop through dynamically added checkboxes in the AccountFrame
    For i = 1 To Accounts.Count
        Set CheckAccount = Me.AccountFrame.Controls("CheckAccountName" & i)
        
        If CheckAccount.value = True Then
            ' Add the selected account to the collection
            SelectedAccounts.Add Accounts(i)
        End If
    Next i
    
    '' Ensure at least one account is selected
    If SelectedAccounts.Count = 0 Then
        MsgBox "Please select at least one account.", vbExclamation
        Exit Sub
    End If
    
    '' Hide the UserForm
    Me.Hide
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



' Event handler for the CheckAll checkbox
Private Sub CheckAll_Click()
    Dim i As Integer
    Dim CheckAccount As MSForms.CheckBox
    Dim ShouldCheck As Boolean
    
    ' Determine the desired state based on the CheckAll checkbox
    ShouldCheck = CheckAll.value
    
    ' Loop through dynamically added checkboxes in the AccountFrame
    For i = 1 To Accounts.Count
        ' Reference each checkbox by its new name
        Set CheckAccount = Me.AccountFrame.Controls("CheckAccountName" & i)
        If Not CheckAccount Is Nothing Then
            CheckAccount.value = ShouldCheck
        End If
    Next i
    
    ' Update UserForm
    Me.Repaint
End Sub



' Event handler to set focus on the continue button
Private Sub UserForm_Activate()
    ' set focus
    ButtonContinue.SetFocus
End Sub

