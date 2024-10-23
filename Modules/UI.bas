Attribute VB_Name = "UI"
' Launch AutoArchive through Outlook UI
Sub RunAutoArchiveNow()
    ' Set up environment
    SetupEnvironment
    
    ' Create the folder if it doesn't exist
    If Dir(SourcePath, vbDirectory) = "" Then MkDir SourcePath
    
    ' Set up log file
    SetupLog
    
    ' DEBUG
    LogDebug "Starting AutoArchive routine through user action." & vbCrLf

    ' Check if config file exists
    If Dir(ConfigFile, vbDirectory) = "" Then SetupMainConfig (ConfigFile)
    
    ' Read contents of main config file
    Set ConfigContents = ReadFile(ConfigFile)
    
    ' launch main script
    AutoArchive
End Sub



' Modify AutoArchive settings through Outlook UI
Sub ModifyAutoArchiveSettings()
    ' Set up environment
    SetupEnvironment
    
    ' Create the folder if it doesn't exist
    If Dir(SourcePath, vbDirectory) = "" Then MkDir SourcePath
    
    ' Set up log file
    SetupLog
    
    ' DEBUG
    LogDebug "Modifying AutoArchive settings." & vbCrLf
    
    ' Check if config file exists
    If Dir(ConfigFile, vbDirectory) = "" Then SetupMainConfig (ConfigFile)
    
    ' Read contents of main config file
    Set ConfigContents = ReadFile(ConfigFile)
    
    ' Show archive settings
    ArchiveSettings.Show
    
    ' Clear global variables
    ClearAll
End Sub
