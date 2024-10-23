Attribute VB_Name = "AutoArchiveRoutine"
Sub AutoArchive()
    LogDebug "Main AutoArchive routine has started." & vbCrLf
    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set variables
    DefaultPath = GetSetting(ConfigContents, "ArchivePath")             ' Get default path from config file
    DefaultPeriod = GetSetting(ConfigContents, "RetainmentPeriod")      ' Get default retainment period from config file
    Continue = True
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Account Collection
    '' Initialize Outlook namespace and accounts
    Set Namespace = Application.GetNamespace("MAPI")
    Set MailAccounts = Namespace.Accounts
    
    
    '' Check if accounts exist
    If MailAccounts.Count = 0 Then
        MsgBox "No accounts found.", vbInformation
        
        ' DEBUG
        LogDebug "Exited Main Script prematurely at Account Collection." & vbCrLf
        
        ' Clear global variables
        ClearAll

        End
    End If
    
    
    ' Create class objects
    Set Accounts = New Collection
    For i = 1 To MailAccounts.Count
        ' Create a new instance of AccountProfileHandler
        Set Account = New AccountProfileHandler
        
        ' Initialize the instance with the selected account
        Account.Initialize MailAccounts(i)
        
        ' Add the new instance to the collection
        Accounts.Add Account
    Next i
    
    
    '' DEBUG: List accounts
    LogDebug "Accounts Found:"
    For i = 1 To Accounts.Count
        LogDebug " - " & Accounts(i).GetAccountName()
    Next i
    LogDebug ""
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Account Selection
    '' Show the UserForm and let the user select accounts
    AccountSelection.Show
        
    '' Check if user aborted process
    If Not Continue Then
        MsgBox "Operation cancelled.", vbInformation
        
        ' DEBUG
        LogDebug "Exited Main Script prematurely at Account Selection." & vbCrLf
        
        ' Clear global variables
        ClearAll
    
        End
    End If
    
    '' DEBUG: List selected accounts
    LogDebug "Accounts Selected:"
    For i = 1 To SelectedAccounts.Count
        LogDebug " - " & SelectedAccounts(i).GetAccountName()
    Next i
    LogDebug ""



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Folder Selection
    '' Show the UserForm and let the user select the archive locations
    FolderSelection.Show

    '' Check if user aborted process
    If Not Continue Then
        MsgBox "Operation cancelled.", vbInformation
        
        ' DEBUG
        LogDebug "Exited Main Script prematurely at Folder Selection." & vbCrLf

        ' Clear global variables
        ClearAll

        End
    End If
    
    '' DEBUG: List account archive folders
    LogDebug "Account archive folders:"
    For i = 1 To SelectedAccounts.Count
        LogDebug " - " & SelectedAccounts(i).GetArchivePath()
    Next i
    LogDebug ""
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Mail Folder Selection
    For i = 1 To SelectedAccounts.Count
        ' Initiate collections
        Set Account = SelectedAccounts(i)
        
        ' Create account folder list
        Account.AccountFolderSetup
        
        ' Check if it failed
        If Not Continue Then
            Exit For
        End If
        
        ' Create and call the MailFolderForm
        Set AccountFolderForm = New AccountFolderSelection
        AccountFolderForm.Show
        
        ' Check if it got cancelled
        If Not Continue Then
            Exit For
        End If
    Next i
    
    '' Check if Mail Folder Selection failed
    If Not Continue Then
        MsgBox "Operation cancelled.", vbInformation
                
        ' DEBUG
        LogDebug "Exited Main Script prematurely at Mail Folder Selection." & vbCrLf
        
        ' Clear global variables
        ClearAll

        End
    End If
    
    '' DEBUG: List account folders to be archived
    For i = 1 To SelectedAccounts.Count
        Dim Folders As Collection
        Dim FolderSettings As Collection
        
        Set Folders = SelectedAccounts(i).GetArchiveFolders()
        Set FolderSettings = SelectedAccounts(i).GetArchiveFolderSettings()
        
        LogDebug "Selected " & Folders.Count & " folders for archiving of account " & SelectedAccounts(i).GetAccountName() & ":"

        For j = 1 To Folders.Count
            LogDebug " - " & "Keeping last " & FolderSettings(j)(2) & " days of folder " & Folders(j)
        Next j
        
        LogDebug ""
    Next i
    
    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Final userprompt to start archiving
    StartPrompt.Show
    
    If Not Continue Then
        MsgBox "Operation cancelled.", vbInformation
        
        ' DEBUG
        LogDebug "Exited Main Script prematurely at final Userprompt." & vbCrLf
        
        ' Clear global variables
        ClearAll

        End
    End If



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Archive procedure
    For i = 1 To SelectedAccounts.Count
        ' Initiate collections
        Set Account = SelectedAccounts(i)
        
        ' start archive procedure
        Account.ArchiveSelectedFolders
        
        ' Check if it got cancelled
        If Not Continue Then
            Exit For
        End If
        
        ' Update account config file
        SelectedAccounts(i).WriteAccountConfigContents
    Next i
    
    If Not Continue Then
        MsgBox "Operation failed.", vbInformation
        
        ' DEBUG
        LogDebug "Exited Main Script prematurely at archive procedure." & vbCrLf
        
        ' Clear global variables
        ClearAll

        End
    End If
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Update main config file
    Call ReplaceSetting(ConfigContents, "ExecutionDate", CurrentDate)
    Call WriteFile(ConfigFile, ConfigContents)



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Finalization
    LogDebug "Main AutoArchive routine has completed." & vbCrLf
    MsgBox "Archive process completed successfully!", vbInformation, "Archive Completed"
    
    ' Clear global variables
    ClearAll

End Sub
