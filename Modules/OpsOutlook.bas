Attribute VB_Name = "OpsOutlook"
' Start the Send/Receive All Folders process
Sub SendReceiveAllMail()
    Dim App As Outlook.Application
    
    Set App = Application
    
    ' Execute the Send/Receive All Folders command
    App.Session.SendAndReceive (True)
    
    ' DEBUG
    Debug.Print "Fetching new mail" & vbCrLf
End Sub



' Unload the archive if it is currently loaded in Outlook
Sub UnloadArchive(ArchiveFile As String)
    ' Variables
    Dim ArchiveRoot As Outlook.Folder
    Dim RetryAttempts As Integer
    Dim MaxAttempts As Integer
    Dim StoreUnloaded As Boolean

    ' Define the maximum number of retry attempts to release the file lock
    MaxAttempts = 5
    RetryAttempts = 0
    StoreUnloaded = False

    ' Attempt to get the root folder based on the archive file path
    Set ArchiveRoot = GetStoreRootByPath(ArchiveFile)

    ' Check if the ArchiveRoot is found
    If Not ArchiveRoot Is Nothing Then
        ' Outer loop that checks if the file is still in use
        Do While IsFileInUse(ArchiveFile) And RetryAttempts < MaxAttempts
            ' Attempt to remove the archive store from Outlook UI
            If Not ArchiveRoot Is Nothing Then
                Namespace.RemoveStore ArchiveRoot
                LogDebug "Unloaded archive file " & ArchiveFile & " from Outlook UI." & vbCrLf
            Else
                LogDebug "Archive file " & ArchiveFile & " not loaded in Outlook UI." & vbCrLf
            End If


            ' Attempt to get the root folder based on the archive file path
            Set ArchiveRoot = GetStoreRootByPath(ArchiveFile)

            ' Check if the file is still in use after the unload attempt
            If Not IsFileInUse(ArchiveFile) Then
                StoreUnloaded = True
                LogDebug "Attempt " & RetryAttempts + 1 & ": File is no longer in use." & vbCrLf
                Exit Do ' Exit the loop if the file is no longer in use
            End If

            ' Run DoEvents to allow the system to process any pending events or messages
            DoEvents

            ' Increment the retry attempts counter
            RetryAttempts = RetryAttempts + 1
        Loop
        
        ' Final status message based on whether unloading was successful or not
        If StoreUnloaded Then
            LogDebug "Successfully unloaded the archive file " & ArchiveFile & " from Outlook UI after " & RetryAttempts & " attempts." & vbCrLf
        Else
            LogDebug "Failed to completely unload the archive file " & ArchiveFile & " after " & MaxAttempts & " attempts." & vbCrLf
        End If
    Else
        ' If the archive is not found, print a message
        LogDebug "Archive file " & ArchiveFile & " is not loaded in Outlook UI." & vbCrLf
    End If

    ' Clear object references
    Set ArchiveRoot = Nothing
End Sub



' Unload all known archives
Sub UnloadAllArchives()
    Dim i As Integer
    Dim confFile As String
    Dim confFileCollection As Collection
    Dim confFileContents As Collection
    Dim FileName As String
    
    ' Check if SelectedAccounts contains any items
    If Not SelectedAccounts Is Nothing Then
        ' Unload every archive file of selected accounts
        For i = 1 To SelectedAccounts.Count
            UnloadArchive SelectedAccounts(i).GetArchiveFile
        Next i
    Else
        ' Initialize the collection to store .conf files
        Set confFileCollection = New Collection

        ' Get all .conf files in the source directory (excluding "autoarchive.conf")
        confFile = Dir(SourcePath & "*.conf")
        
        Do While confFile <> ""
            ' Add .conf file to collection if it's not "autoarchive.conf"
            If confFile <> "autoarchive.conf" Then
                confFileCollection.Add confFile
            End If
            
            ' Get the next .conf file in the directory
            confFile = Dir
        Loop

        ' Define Namespace if it's not defined
        If Namespace Is Nothing Then
            Set Namespace = Application.GetNamespace("MAPI")
        End If

        ' Loop through the collection of .conf files
        For i = 1 To confFileCollection.Count
            ' Read the contents of the .conf file
            Set confFileContents = ReadFile(SourcePath & confFileCollection(i))

            ' Construct the archive file path from the .conf file
            FileName = GetSetting(confFileContents, "ArchivePath") & "archive.pst"

            ' If a valid ArchivePath is found, try to unload its archive
            If FileName <> "" Then
                ' Unload the archive file if it exists
                UnloadArchive FileName
            End If
        Next i
    End If
End Sub



' Get folders for archiving from current account
Function FetchMailFolder(Root As Outlook.Folder, FolderList As Collection)
    ' Variable declaration
    Dim Subfolder As Outlook.Folder
    Dim Item As Object
    Dim HasItem As Boolean
    Dim IsLocal As Boolean

    ' Check if the folder is a local folder
    IsLocal = (InStr(Root.name, "(This computer only)") > 0)

    ' Initialize the flag for mail items
    HasItem = False
    
    ' Check if there are any mail items in the folder
    For Each Item In Root.Items
        If Item.Class = olMail Or _
           Item.Class = olReport Or _
           Item.Class = olMeetingRequest Or _
           Item.Class = olMeetingCancellation Or _
           Item.Class = olMeetingForwardNotification Or _
           Item.Class = olMeetingResponseNegative Or _
           Item.Class = olMeetingResponsePositive Or _
           Item.Class = olMeetingResponseTentative Then
            HasItem = True
            Exit For
        End If
    Next Item

    ' Exclude local folders and add folder if it contains mail items
    If Not IsLocal And HasItem Then
        FolderList.Add Root.FolderPath
    End If

    ' Recursively get all subfolders
    For Each Subfolder In Root.Folders
        FetchMailFolder Subfolder, FolderList
    Next Subfolder
End Function



' Run all active rules for mail account inbox
Sub RunAllAccountRules(Account As Outlook.Account)
    Dim Rule As Outlook.Rule
    Dim Rules As Outlook.Rules
    Dim i As Integer
    Dim Inbox As Outlook.Folder

    ' Retrieve the rules collection for the account's store
    Set Rules = Account.DeliveryStore.GetRules()

    ' Check if the account has any rules
    If Rules.Count = 0 Then
        LogDebug "No rules found for account: " & Account.DisplayName & vbCrLf
        Exit Sub
    End If

    ' Get the inbox folder for the account
    Set Inbox = Account.DeliveryStore.GetDefaultFolder(olFolderInbox)
    
    ' Loop through and execute enabled rules
    LogDebug "Running " & Rules.Count & " rules for account: " & Account.DisplayName
    For i = 1 To Rules.Count
        Set Rule = Rules(i)
        If Rule.Enabled Then
            Rule.Execute ShowProgress:=True, Folder:=Inbox
            LogDebug " - Rule: " & Rule.name
        End If
    Next i
    LogDebug ""
End Sub



' Check duplicate items
Function IsDuplicate(Item As Object) As Boolean
    Dim ArchiveItem As Object
    Dim ItemFound As Boolean
    ItemFound = False

    ' Check if ArchiveFolder contains items
    If ArchiveFolder.Items.Count > 0 Then
        ' Loop through each mail item in the ArchiveFolder
        For Each ArchiveItem In ArchiveFolder.Items
            ' Compare the Class of both items
            If Item.Class = ArchiveItem.Class Then
                Select Case Item.Class
                    ' Check for mail items
                    Case olMail
                        If Item.ReceivedTime = ArchiveItem.ReceivedTime And _
                           Item.SentOn = ArchiveItem.SentOn And _
                           Item.Subject = ArchiveItem.Subject And _
                           Item.SenderName = ArchiveItem.SenderName And _
                           Item.SenderEmailAddress = ArchiveItem.SenderEmailAddress And _
                           Item.Body = ArchiveItem.Body And _
                           Item.BodyFormat = ArchiveItem.BodyFormat Then
                            ItemFound = True
                            Exit For
                        End If
                    ' Check for report items
                    Case olReport
                        If Item.CreationTime = ArchiveItem.CreationTime And _
                           Item.Subject = ArchiveItem.Subject And _
                           Item.Size = ArchiveItem.Size And _
                           Item.Body = ArchiveItem.Body Then
                            ItemFound = True
                            Exit For
                        End If
                    ' Check for meeting items
                    Case olMeetingRequest, olMeetingCancellation, olMeetingForwardNotification, _
                         olMeetingResponseNegative, olMeetingResponsePositive, olMeetingResponseTentative
                        If Item.Subject = ArchiveItem.Subject And _
                           Item.SenderName = ArchiveItem.SenderName And _
                           Item.SentOn = ArchiveItem.SentOn And _
                           Item.Body = ArchiveItem.Body And _
                           Item.BodyFormat = ArchiveItem.BodyFormat Then
                            ItemFound = True
                            Exit For
                        End If
                End Select
            End If
        Next ArchiveItem
    End If

    ' Return True if duplicate is found, otherwise False
    IsDuplicate = ItemFound
End Function



' Retrieve the root folder of a store by its file path
Function GetStoreRootByPath(ByVal FilePath As String) As Outlook.Folder
    Dim Store As Outlook.Store
    Dim RootFolder As Outlook.Folder

    ' Iterate through each store to find the one that matches the file path
    For Each Store In Namespace.Stores
        If Store.FilePath = FilePath Then
            ' Get the root folder for the store
            Set GetStoreRootByPath = Store.GetRootFolder
            
            ' DEBUG
            'LogDebug "File " & FilePath & " has root folder: " & Store.GetRootFolder & vbCrLf
            Exit Function
        End If
    Next Store

    ' If no matching store is found, return Nothing
    Set GetStoreRootByPath = Nothing
    
    ' DEBUG
    'LogDebug "Archive file " & FilePath & " has no active root folder." & vbCrLf
End Function



' Function to create a folder (including subfolders) in the archive root
Sub CreateFolderByPath(ByVal Root As Outlook.Folder, ByVal FolderPath As String)
    Dim FolderParts As Variant
    Dim FolderCollection As Collection
    Dim CurrentFolderName As String
    Dim RemainingFolderPath As String
    Dim Subfolder As Outlook.Folder
    Dim FolderFound As Boolean
    Dim i As Integer

    ' Split the input FolderPath into parts based on the delimiter "\"
    FolderParts = Split(FolderPath, "\")

    ' Initialize a new collection to hold only the valid folder parts
    Set FolderCollection = New Collection

    ' Loop through FolderParts and add only non-empty elements to the collection
    For i = LBound(FolderParts) To UBound(FolderParts)
        If Trim(FolderParts(i)) <> "" Then
            FolderCollection.Add FolderParts(i)
        End If
    Next i

    ' Stop if no valid folder parts are found after cleanup
    If FolderCollection.Count = 0 Then Exit Sub

    ' The first folder name to check/create
    CurrentFolderName = FolderCollection(1)

    ' Initialize a flag to indicate whether the folder is found or not
    FolderFound = False

    ' Check if the folder already exists within the root folder
    For Each Subfolder In Root.Folders
        If Subfolder.name = CurrentFolderName Then
            Set Root = Subfolder
            FolderFound = True
            Exit For
        End If
    Next Subfolder

    ' If not found, create the folder
    If Not FolderFound Then
        Set Root = Root.Folders.Add(CurrentFolderName)
    End If

    ' Rebuild the remaining path by iterating over the collection, starting from the second item
    RemainingFolderPath = "\"
    For i = 2 To FolderCollection.Count
        RemainingFolderPath = RemainingFolderPath & FolderCollection(i) & "\"
    Next i

    ' Call recursively if there's more path to process
    If RemainingFolderPath <> "\" Then
        CreateFolderByPath Root, RemainingFolderPath
    End If
End Sub



' Function to get a folder by its path from a specified root folder
Public Function GetFolderByPath(ByVal Root As Outlook.Folder, ByVal FolderPath As String) As Outlook.Folder
    Dim FolderParts As Variant
    Dim FolderCollection As Collection
    Dim CurrentFolderName As String
    Dim RemainingFolderPath As String
    Dim Subfolder As Outlook.Folder
    Dim i As Integer
    Dim FolderFound As Boolean
    
    ' Split the input FolderPath into parts based on the delimiter "\"
    FolderParts = Split(FolderPath, "\")
    
    ' Initialize a new collection to hold only the valid folder parts
    Set FolderCollection = New Collection
    
    ' Loop through FolderParts and add only non-empty elements to the collection
    For i = LBound(FolderParts) To UBound(FolderParts)
        If Trim(FolderParts(i)) <> "" Then
            FolderCollection.Add FolderParts(i)
        End If
    Next i
    
    ' Stop if no valid folder parts are found after cleanup
    If FolderCollection.Count = 0 Then
        Set GetFolderByPath = Root
        Exit Function
    End If

    ' The first folder name to search within the root folder
    CurrentFolderName = FolderCollection(1)

    ' Initialize a flag to indicate whether the folder is found or not
    FolderFound = False
    
   ' Check if the folder already exists within the root folder
    For Each Subfolder In Root.Folders
        If Subfolder.name = CurrentFolderName Then
            Set Root = Subfolder
            FolderFound = True
            
            'LogDebug "Found folder: '" & CurrentFolderName & "' in root folder: '" & Root.name & "'"
            
            Exit For
        End If
    Next Subfolder

    ' If not found, return nothing and exit
    If Not FolderFound Then
        LogDebug "Could not find folder '" & CurrentFolderName & "' in root folder '" & Root.name & "'"
        
        Set GetFolderByPath = Nothing
        Exit Function
    End If

    ' If there's only one folder left, we've reached the end, return it
    If FolderCollection.Count = 1 Then
        Set GetFolderByPath = Root
        Exit Function
    End If

    ' Rebuild the remaining path by iterating over the collection, starting from the second item
    RemainingFolderPath = "\"  ' Ensure the leading backslash is present
    For i = 2 To FolderCollection.Count
        RemainingFolderPath = RemainingFolderPath & FolderCollection(i) & "\"
    Next i

    ' Call GetFolderByPath recursively
    Set GetFolderByPath = GetFolderByPath(Root, RemainingFolderPath)
End Function



' Clean up and restart Outlook
Sub CloseOutlook()
    ' Unload all active archive stores
    UnloadAllArchives
    
    ' Clear global variables
    ClearAll

    ' Close Outlook after all archives are unloaded
    ' Quit Outlook application gracefully
    Application.Quit

    ' Allow a short delay to give Outlook time to close
    'Delay 3000

    ' Forcefully kill the Outlook process using Taskkill
    'Shell "taskkill /F /IM outlook.exe", vbHide
End Sub

