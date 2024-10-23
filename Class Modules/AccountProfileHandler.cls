VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccountProfileHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class module: AccountProfileHandler



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Attributes
'' Name and account
Private Account As Outlook.Account              ' Outlook account
Private AccountName As String                   ' Account name
Private AccountRoot As Outlook.Folder            ' Account root folder
Private AccountFolders As Collection            ' Account folders
Private AccountFolderSettings As Collection     ' Account folder settings

'' Config file
Private FileName As String                      ' Account name suitable for file explorer
Private AccountConfigFile As String             ' Account config file
Private AccountConfigContents As Collection     ' Account config file contents

'' Archive file
Private ArchivePath As String                   ' Archive folder
Private ArchiveFile As String                   ' Archive file
Private ArchiveFolders As Collection            ' Archive folders
Private ArchiveFolderSettings As Collection     ' Archive folder settings
Private ArchiveMail As Collection               ' Mail items to be archived
Private ArchiveRoot As Outlook.Folder           ' Archive root folder
Private ArchiveBackup As String                 ' Archive backup file

'' Settings
Private ExecutionDate As String                 ' Last archived

'' Counters
Private i As Integer                            ' Loop counter for iterating through collections
Private j As Integer                            ' Loop counter for iterating through collections



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialize method
Public Sub Initialize(ByVal OutlookAccount As Outlook.Account)
    ' Account and names
    Set Account = OutlookAccount
    AccountName = Account.DisplayName
    
    ' Config file
    FileName = Replace(Replace(AccountName, ".", "_"), "@", "_at_")
    AccountConfigFile = SourcePath & FileName & ".conf"
    
    
    ' Check if the ConfigFile exists and read its contents
    If Dir(AccountConfigFile) <> "" Then
        ' Read the file contents
        Set AccountConfigContents = ReadFile(AccountConfigFile)
        
        ' Extract ExecutionDate and ArchivePath from the ConfigFileContents
        ExecutionDate = GetSetting(AccountConfigContents, "ExecutionDate")
        ArchivePath = GetSetting(AccountConfigContents, "ArchivePath")
        
        ' Set archive and backup files
        ArchiveFile = ArchivePath & "archive.pst"
        ArchiveBackup = ArchiveFile & ".bak"
    
    ' if no file exists, set default values
    Else
        ' Create the default archive path if it does not exist
        If Dir(DefaultPath, vbDirectory) = "" Then
            MkDir DefaultPath
            LogDebug "Created folder for new archives: " & DefaultPath & vbCrLf
        End If
        
        ' Account has never been archieved
        ExecutionDate = "never"
        
        ' Set default archive path
        ArchivePath = DefaultPath & FileName & "\"
        
        ' Set archive and backup files
        ArchiveFile = ArchivePath & "archive.pst"
        ArchiveBackup = ArchiveFile & ".bak"
        
        ' Create config contents
        Set AccountConfigContents = New Collection
        AccountConfigContents.Add "ExecutionDate=never"
        AccountConfigContents.Add "ArchivePath=" & ArchivePath
    End If
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions
'' Set Archive Path and File
Public Sub ChangeArchiveLocation(ByVal Folder As String)
    ' Check if the archive file exists
    If Dir(ArchiveFile) <> "" Then
        ' Query to move existing archive
        MoveArchive = MsgBox("The archive location has changed from:" & vbCrLf & _
                     ArchivePath & vbCrLf & _
                     "to:" & vbCrLf & _
                     Folder & vbCrLf & _
                     "Do you want to move the archive.pst file to the new location?", _
                     vbYesNo + vbQuestion, "Move Archive")
        
        If MoveArchive = vbYes Then
            ' Folder does not exist, create it
            If Dir(Folder, vbDirectory) = "" Then MkDir Folder

            ' Unload the file if its loaded
            UnloadArchive ArchiveFile
            
            ' Check if file is still in use
            If Not IsFileInUse(ArchiveFile) Then
                ' Move the archive file to the new location
                FileCopy ArchiveFile, Folder & "archive.pst"
                
                ' Delete the old archive.pst file
                Kill ArchiveFile
                
                ' Check if backup file exists and move it
                If Dir(ArchiveBackup) <> "" Then
                    FileCopy ArchiveBackup, Folder & "archive.pst.bak"
                    
                    ' Delete the old backup file after moving
                    Kill ArchiveBackup
                End If
                
                ' Check if the old ArchivePath directory is empty and delete it
                If CountFilesInDirectory(ArchivePath) = 0 Then
                    RmDir ArchivePath  ' Remove the directory if it's empty
                    LogDebug "Deleted empty directory: " & ArchivePath & vbCrLf
                End If

                
                ' Set new archive path and file
                ArchivePath = Folder
                ArchiveFile = ArchivePath & "archive.pst"
                ArchiveBackup = ArchiveFile & ".bak"
                
                ' Update config
                Call ReplaceSetting(AccountConfigContents, "ArchivePath", ArchivePath)
                Call WriteFile(AccountConfigFile, AccountConfigContents)
                
                ' DEBUG
                LogDebug "Updated archive location of account " & AccountName & " to: " & Folder & vbCrLf
            Else
                ' DEBUG
                LogDebug "Failed to update archive location of account " & AccountName & vbCrLf
                
                HandleArchiveFileInUse
                
                If Not Continue Then LogDebug "Move operation canceled by user due to file in use." & vbCrLf
            End If
        End If
    Else
        ' Set new archive path and file
        ArchivePath = Folder
        ArchiveFile = ArchivePath & "archive.pst"
        ArchiveBackup = ArchiveFile & ".bak"
        
        ' Update config
        Call ReplaceSetting(AccountConfigContents, "ArchivePath", ArchivePath)
        Call WriteFile(AccountConfigFile, AccountConfigContents)

        LogDebug "Updated ArchivePath and ArchiveFile for account " & AccountName & " to: " & Folder & vbCrLf
    End If
End Sub



'' Build account and archive folders
Public Sub AccountFolderSetup()
    ' Run rules to ensure all mailitems are in the correct folder
    RunAllAccountRules Account

    ' Get account root folder
    Set AccountRoot = Account.DeliveryStore.GetRootFolder
    
    ' Get list of all eligible mail folders
    Set AccountFolders = New Collection
    Call FetchMailFolder(AccountRoot, AccountFolders)
    
    ' Check if folder list is empty
    If AccountFolders.Count = 0 Then
        MsgBox "Account: " & AccountName & vbCrLf & _
            "No folders to be archived." & vbCrLf & _
            "Please check the account setup.", _
            vbExclamation + vbOKOnly, _
            "No Folders Found"
        
        Continue = False
        Exit Sub
    End If

    ' DEBUG
    LogDebug "Eligible folders for archiving of account " & AccountName & ":"
    For i = 1 To AccountFolders.Count
        LogDebug " - " & AccountFolders(i)
    Next i
    LogDebug ""
    
    
    ' Settings for AccountFolders found
    Set AccountFolderSettings = New Collection
    
    ' Initialize temporary variables
    Dim FolderName As String
    Dim FolderSettings As Collection
    
    For i = 1 To AccountFolders.Count
        ' Get folder name and clean it
        FolderName = FormatFolderPath(AccountFolders(i), AccountName)
        
        ' Get the settings for the folder
        Set FolderSettings = GetFolderSettings(AccountConfigContents, FolderName)

        ' Check if settings were found
        If FolderSettings.Count > 0 Then
            ' Add found settings to AccountFolderSettings as a single entry
            AccountFolderSettings.Add FolderSettings
        Else
            ' Add a collection with "never" and DefaultPeriod to AccountFolderSettings
            Dim DefaultSettings As Collection
            Set DefaultSettings = New Collection
            
            DefaultSettings.Add "never"                 ' Archive Date
            DefaultSettings.Add DefaultPeriod           ' Retainment Period
            
            AccountFolderSettings.Add DefaultSettings
        End If
    Next i
End Sub



'' Archive procedure
Public Sub ArchiveSelectedFolders()
    Dim FolderName As String
    Dim MailItem As Object
    Dim ItemDate As String
    Dim ItemSubject As String
    Dim DaysPassed As Long
    
    LogDebug "Archiving mail for account " & AccountName & vbCrLf
    
    ' Check if the archive path exists, create if it doesn't
    If Dir(ArchivePath, vbDirectory) = "" Then
        MkDir ArchivePath
        LogDebug "Created archive path: " & ArchivePath & vbCrLf
    End If


    ' Check if the archive file already exists
    If Dir(ArchiveFile) = "" Then
        ' Add the PST file to Outlook and get the root folder
        Namespace.AddStoreEx ArchiveFile, olStoreDefault

        ' Retrieve the archive store by file path
        Set ArchiveRoot = GetStoreRootByPath(ArchiveFile)

        ' Rename the archive root folder (if necessary)
        ArchiveRoot.name = "Archive of " & AccountName
        
        ' DEBUG: Print the root folder name and path
        LogDebug "Created archive PST: " & ArchiveFile & vbCrLf
    Else
        ' Backup existing archive if it exists
        BackupExistingArchive
        
        ' Check if it got cancelled
        If Not Continue Then Exit Sub

        ' Ensure PST is loaded in Outlook
        Namespace.AddStoreEx ArchiveFile, olStoreDefault
        
        ' Retrieve the archive store by file path
        Set ArchiveRoot = GetStoreRootByPath(ArchiveFile)
        
        ' DEBUG: Print the root folder name and path
        LogDebug "Loaded Archive PST: " & ArchiveFile & vbCrLf
    End If
    
    ' Create the duplicate folder
    Call CreateFolderByPath(ArchiveRoot, "\Duplicates")
    Set DuplicateFolder = GetFolderByPath(ArchiveRoot, "\Duplicates")
    
    ' Mirror the folders found in ArchiveFolders
    For i = 1 To ArchiveFolders.Count
        ' Format the folder path to exclude the leading account name part
        FolderName = FormatFolderPath(ArchiveFolders(i), AccountName)
        
        LogDebug " - Archiving mail in folder " & FolderName
        
        ' Create non-existent folders in the archive root
        Call CreateFolderByPath(ArchiveRoot, FolderName)
        
        ' Retrieve the corresponding folders in the mail account and archive
        Set AccountFolder = GetFolderByPath(AccountRoot, FolderName)
        Set ArchiveFolder = GetFolderByPath(ArchiveRoot, FolderName)
        
        ' Ensure both AccountFolder and ArchiveFolder are valid
        If AccountFolder Is Nothing Or ArchiveFolder Is Nothing Then
            Continue = False
            
            LogDebug "Failed finding resprective account or archive folders for " & FolderName & vbCrLf
            
            Exit For
        End If
        
        Set ValidItems = New Collection
        
        LogDebug "   Retainment period: " & ArchiveFolderSettings(i)(2)
        
        ' Iterate through each item in the account folder
        For Each Item In AccountFolder.Items
            Select Case Item.Class
                ' Mail item
                Case olMail
                    ' Subject
                    ItemSubject = Replace(Replace(Replace(CStr(Item.Subject), vbCrLf, ""), vbCr, ""), vbLf, "")
                    
                    ' For mail items, use the ReceivedTime
                    ItemDate = Format(Item.ReceivedTime, "dd/mm/yyyy")
                    
                    ' Calculate the number of days passed since the mail was received
                    DaysPassed = DiffDates(ItemDate, CurrentDate)
                    
                    ' Compare with the archive threshold specified in AccountFolderSettings
                    If DaysPassed > ArchiveFolderSettings(i)(2) Then
                        ' Add to the collection of valid items
                        ValidItems.Add Item
                        
                        ' Debugging statement for added item
                        LogDebug "   - Valid mail item: '" & ItemSubject & "' since days passed: " & DaysPassed & ", Item Date: " & ItemDate
                    Else
                        ' Debugging statement for items not moved
                        LogDebug "   - Invalid mail item: '" & ItemSubject & "' since days passed: " & DaysPassed & ", Item Date: " & ItemDate
                    End If
                ' Report item
                Case olReport
                    ' Subject
                    ItemSubject = Replace(Replace(Replace(CStr(Item.Subject), vbCrLf, ""), vbCr, ""), vbLf, "")
                    
                    ' For mail items, use the ReceivedTime
                    ItemDate = Format(Item.CreationTime, "dd/mm/yyyy")
                    
                    ' Calculate the number of days passed since the mail was received
                    DaysPassed = DiffDates(ItemDate, CurrentDate)
                    
                    ' Compare with the archive threshold specified in AccountFolderSettings
                    If DaysPassed > ArchiveFolderSettings(i)(2) Then
                        ' Add to the collection of valid items
                        ValidItems.Add Item
                        
                        ' Debugging statement for added item
                        LogDebug "   - Valid report item: '" & ItemSubject & "' since days passed: " & DaysPassed & ", Item Date: " & ItemDate
                    Else
                        ' Debugging statement for items not moved
                        LogDebug "   - Invalid report item: '" & ItemSubject & "' since days passed: " & DaysPassed & ", Item Date: " & ItemDate
                    End If
                ' Meeting item
                Case olMeetingRequest, olMeetingCancellation, olMeetingForwardNotification, _
                     olMeetingResponseNegative, olMeetingResponsePositive, olMeetingResponseTentative
                    ' Subject
                    ItemSubject = Replace(Replace(Replace(CStr(Item.Subject), vbCrLf, ""), vbCr, ""), vbLf, "")
                    
                    ' For meeting-related items, use the SentOn date
                    ItemDate = Format(Item.SentOn, "dd/mm/yyyy")
                    
                    ' Calculate the number of days passed since the meeting was sent
                    DaysPassed = DiffDates(ItemDate, CurrentDate)
                    
                    ' Compare with the archive threshold
                    If DaysPassed > ArchiveFolderSettings(i)(2) Then
                        ' Add to the collection of valid items
                        ValidItems.Add Item
                        
                        ' Debugging statement
                        LogDebug "   - Valid meeting item: '" & ItemSubject & "' since days passed: " & DaysPassed & ", Item Date: " & ItemDate
                    Else
                        ' Debugging statement for items not moved
                        LogDebug "   - Invalid meeting item: '" & ItemSubject & "' since days passed: " & DaysPassed & ", Item Date: " & ItemDate
                    End If
            End Select
        Next Item

        LogDebug ""

        If ValidItems.Count > 0 Then
            ' Create and call the ArchiveProgressForm
            Set ArchiveProgressForm = New ArchiveProgress
            ArchiveProgressForm.ProcessItems
        Else
            LogDebug " - No mail for archiving found in folder " & FolderName
        End If
        
        LogDebug ""
        
        If Not Continue Then Exit For
    Next i
End Sub



' Update Account Config File
Public Sub WriteAccountConfigContents()
    Dim Folder As String
    Dim Replaced As Boolean
    Dim Contents As Collection
    Set Contents = New Collection
    
    ' Add general settings
    Contents.Add "ExecutionDate=" & CurrentDate
    Contents.Add "ArchivePath=" & ArchivePath
    
    ' Check for changes in folders that have been archived before
    For i = 1 To AccountConfigContents.Count
        ' Checks if folder thats existing in AccountConfigContents is sceduled to be archived again
        Replaced = False
        
        For j = 1 To ArchiveFolders.Count
            Folder = FormatFolderPath(ArchiveFolders(j), AccountName)
            
            If Left(AccountConfigContents(i), Len(Folder & "|")) = Folder & "|" Then
                ' Add new settings to the config
                Contents.Add Folder & "|" & ArchiveFolderSettings(j)(1) & "|" & ArchiveFolderSettings(j)(2)
                Replaced = True
            End If
            
            If Replaced Then
                Exit For
            End If
        Next j
        
        ' Add the old value if folder is not sceduled to be archived again
        If Not Replaced And Left(AccountConfigContents(i), 1) = "\" Then
            Contents.Add AccountConfigContents(i)
        End If
    Next i
    
    ' Add new entries
    For i = 1 To ArchiveFolders.Count
        Folder = FormatFolderPath(ArchiveFolders(i), AccountName)
        
        ' If its not in the config file, add it
        If GetEntry(AccountConfigContents, Folder & "|") = "" Then
            Contents.Add Folder & "|" & ArchiveFolderSettings(i)(1) & "|" & ArchiveFolderSettings(i)(2)
        End If
    Next i
    
    ' Set new config file contents
    Set AccountConfigContents = Contents
    
    ' Update config file
    Call WriteFile(AccountConfigFile, AccountConfigContents)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions
'' Function to backup existing archive if it exists
Private Sub BackupExistingArchive()
    ' Onload the archive file
    UnloadArchive ArchiveFile
    
    ' Check if file is still in use
    If Not IsFileInUse(ArchiveFile) Then
        ' Check if the backup file already exists and delete it
        If Dir(ArchiveBackup) <> "" Then Kill ArchiveBackup

        ' Create a backup of the existing archive file
        FileCopy ArchiveFile, ArchiveBackup
    
        ' DEBUG
        LogDebug "Backup created at: " & ArchiveBackup & vbCrLf
    Else
        ' DEBUG
        LogDebug "Failed to create backup at: " & ArchiveBackup & vbCrLf
        
        HandleArchiveFileInUse

        If Not Continue Then LogDebug "Backup operation canceled by user due to file in use." & vbCrLf
    End If
End Sub



' User prompt if archive file could not be unloaded
Private Sub HandleArchiveFileInUse()
    Dim UnloadAndClose As VbMsgBoxResult
    
    ' Prompt the user for further action if the file is still in use
    UnloadAndClose = MsgBox( _
        "The archive file '" & ArchiveFile & "' is currently in use, and the script cannot access, move, or modify it." & vbCrLf & vbCrLf & _
        "To proceed, it's recommended to unload all active archive files from Outlook and restart the application." & vbCrLf & _
        "Would you like to unload all active archives and close Outlook now?" & vbCrLf & _
        "Note: Outlook will close but a manual restart is required.", _
        vbYesNo + vbQuestion, _
        "Archive File In Use")

    If UnloadAndClose = vbYes Then
        ' Unload all active stores and restart Outlook
        CloseOutlook
    Else
        Continue = False  ' Indicates that the user canceled the operation
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Getter and Setter functions
'' AccountName
Public Property Get GetAccountName() As String
    GetAccountName = AccountName
End Property



' FileName
Public Property Get GetFileName() As String
    GetFileName = FileName
End Property



' AccountFolders
Public Property Get GetAccountFolders() As Collection
    Set GetAccountFolders = AccountFolders
End Property



' AccountFolderSettings
Public Property Let SetAccountFolderSettings(ByVal FolderSettings As Collection)
    Set AccountFolderSettings = FolderSettings
End Property

Public Property Get GetAccountFolderSettings() As Collection
    Set GetAccountFolderSettings = AccountFolderSettings
End Property


' ArchiveFolders
Public Property Let SetArchiveFolders(ByVal Folders As Collection)
    Set ArchiveFolders = Folders
End Property

Public Property Get GetArchiveFolders() As Collection
    Set GetArchiveFolders = ArchiveFolders
End Property



' ArchiveFolderSettings
Public Property Let SetArchiveFolderSettings(ByVal FolderSettings As Collection)
    Set ArchiveFolderSettings = FolderSettings
End Property

Public Property Get GetArchiveFolderSettings() As Collection
    Set GetArchiveFolderSettings = ArchiveFolderSettings
End Property



' ArchivePath
Public Property Get GetArchivePath() As String
    GetArchivePath = ArchivePath
End Property



' ArchiveFile
Public Property Get GetArchiveFile() As String
    GetArchiveFile = ArchiveFile
End Property



' ExecutionDate
Public Property Get GetExecutionDate() As String
    GetExecutionDate = ExecutionDate
End Property

