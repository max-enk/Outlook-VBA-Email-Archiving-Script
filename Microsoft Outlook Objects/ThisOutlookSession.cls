VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' This runs at Outlook startup
Private Sub Application_Startup()
    ' Delay of 10s
    Delay 10000
    
    ' Send/recieve all folders
    SendReceiveAllMail
    
    ' Set up environment
    SetupEnvironment
    
    ' Ensure the config file exists
    If Dir(ConfigFile, vbDirectory) = "" Then
        ' DEBUG
        Debug.Print "AutoArchive is not configured."
        
        Exit Sub
    End If
    
    ' Create the folder if it doesn't exist
    If Dir(SourcePath, vbDirectory) = "" Then MkDir SourcePath
    
    ' Set up log file
    SetupLog
    
    ' DEBUG
    LogDebug "Starting AutoArchive Routine at Outlook startup"
    
    ' Read contents of main config file
    Set ConfigContents = ReadFile(ConfigFile)
    
    
    ' Check if AutoLaunch is enabled
    If Not LaunchAutomatically Then
        ' DEBUG
        Debug.Print "Aborting automatic AutoArchive launch, AutoLaunch not enabled."
        
        ' Delete log file
        Kill LogFile
        
        Exit Sub
    End If
    
    ' Check if last execution date exceeds set period
    If Not IsArchiveDue Then
        ' DEBUG
        Debug.Print "Aborting automatic AutoArchive launch, ExecutionPeriod not exceeded."
        
        ' Delete log file
        Kill LogFile
        
        Exit Sub
    End If
    
    ' Ask user to start archiving
    RunNow = False
    AutoLaunchPrompt.Show
    
    ' Check if user wants to run AutoArchive
    If RunNow Then
        ' launch main script
        AutoArchive
    Else
        ' DEBUG
        Debug.Print "Automatic AutoArchive launch aborted by user."
        
        ' Delete log file
        Kill LogFile
        
        Exit Sub
    End If
End Sub












