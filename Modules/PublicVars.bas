Attribute VB_Name = "PublicVars"
' Global Variable Declaration
' General Paths
Public SourcePath As String                     ' Path to the directory storing configuration files
Public DefaultPath As String                    ' Default path for new archives

' Files and Contents
Public ConfigFile As String                     ' Path to the main autoarchive configuration file
Public LogFile As String                        ' Path to the current log file
Public ConfigContents As Collection             ' Contents of autoarchive.conf

' Dates and Periods
Public CurrentDate As String                    ' Current date in "DD/MM/YYYY" format
Public DaysSinceExecution As Long               ' Number of days since last execution of AutoArchive
Public DefaultPeriod As String                  ' Retainment period of mail items

' Flags
Public RunNow As Boolean                        ' Flag determined by user to run AutoArchive
Public Continue As Boolean                      ' Flag tracking script abortion

' Outlook Objects
Public Namespace As Outlook.Namespace           ' Outlook namespace object for accessing accounts
Public MailAccounts As Outlook.Accounts         ' All Outlook accounts
Public Accounts As Collection                   ' All Outlook accounts (class object)
Public SelectedAccounts As Collection           ' User-selected Outlook accounts for archiving
Public Account As AccountProfileHandler         ' Specific instance of an account class object

' Objects for archive process
Public AccountFolder As Outlook.Folder          ' Source folder for mails of current account
Public ArchiveFolder As Outlook.Folder          ' Destination folder for mails of current account
Public DuplicateFolder As Outlook.Folder        ' Archive folder for duplicate mailitems
Public ValidItems As Collection                 ' Selection of valid items to be moved

' Forms
Public MailFolderForm As AccountFolderSelection ' Temporary instance of MailFolderSelection UserForm
Public ArchiveProgressForm As ArchiveProgress   ' Temproary instance of ArchiveProcess UserForm
    
' Counters
Public i As Integer                             ' Loop counter for iterating through collections
Public j As Integer                             ' Loop counter for iterating through collections



' Resets all public variables
Sub ClearAll()
    ' General Paths
    SourcePath = ""                             ' Reset SourcePath to empty string
    DefaultPath = ""                            ' Reset DefaultPath to empty string

    ' Files and Contents
    ConfigFile = ""                             ' Reset ConfigFile to empty string
    LogFile = ""                                ' Reset LogFile to empty string
    
    If Not ConfigContents Is Nothing Then
        Set ConfigContents = Nothing            ' Clear the ConfigContents collection
    End If

    ' Dates and Periods
    CurrentDate = ""                            ' Reset CurrentDate to empty string
    DaysSinceExecution = 0                      ' Reset DaysSinceExecution to 0
    DefaultPeriod = ""                          ' Reset DefaultPeriod to empty string

    ' Flags
    RunNow = False                              ' Reset RunNow flag to False
    Continue = False                            ' Reset Continue flag to False

    ' Outlook Objects
    Set Namespace = Nothing                     ' Release Outlook Namespace object
    Set MailAccounts = Nothing                  ' Release Outlook Accounts object
    
    If Not Accounts Is Nothing Then
        Set Accounts = Nothing                  ' Clear the Accounts collection
    End If
    
    If Not SelectedAccounts Is Nothing Then
        Set SelectedAccounts = Nothing          ' Clear the SelectedAccounts collection
    End If
    
    Set Account = Nothing                       ' Release Account class object
    
    ' Archive process objects
    Set AccountFolder = Nothing                 ' Release the source folder for mails of the current account
    Set ArchiveFolder = Nothing                 ' Release the destination folder for mails of the current account
    Set DuplicateFolder = Nothing               ' Release
    
    If Not ValidItems Is Nothing Then
        Set ValidItems = Nothing                ' Clear the collection of valid items
    End If
    
    ' Forms
    If Not MailFolderForm Is Nothing Then
        Unload MailFolderForm                   ' Unload and release MailFolderForm instance
        Set MailFolderForm = Nothing
    End If

    If Not ArchiveProgressForm Is Nothing Then
        Unload ArchiveProgressForm              ' Unload and release ArchiveProcessForm instance
        Set ArchiveProgressForm = Nothing
    End If

    ' Counters
    i = 0                                       ' Reset counter i to 0
    j = 0                                       ' Reset counter j to 0
    
    ' Debug information
    Debug.Print "All global variables have been cleared and reset." & vbCrLf
End Sub

