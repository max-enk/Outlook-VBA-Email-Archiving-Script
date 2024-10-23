Attribute VB_Name = "Initialization"
' Set up environment
Sub SetupEnvironment()
    ' Set current date
    CurrentDate = Format(Date, "DD/MM/YYYY")
    
    ' Create a Shell object
    Dim shellApp As Object                                              ' Shell application object for accessing system paths
    Set shellApp = CreateObject("Shell.Application")
    
    ' Set the config paths
    SourcePath = Environ("LOCALAPPDATA") & "\Outlook AutoArchive\"      ' Source path
    ConfigFile = SourcePath & "autoarchive.conf"                        ' Main config file
End Sub



' Set up log file
Sub SetupLog()
    Dim LogPath As String
    Dim FileCount As Long
    Dim FileName As String
    Dim Contents As Collection
    Dim FileNum As Integer
    
    LogPath = SourcePath & "Logs\"
    
    ' Create the folder if it doesn't exist
    If Dir(LogPath, vbDirectory) = "" Then MkDir LogPath
    
    ' Count existing log files matching log*.log
    FileCount = 0
    FileName = Dir(LogPath & "log*.log")
    While FileName <> ""
        FileCount = FileCount + 1
        FileName = Dir()
    Wend
    
    ' Define the new log file path
    LogFile = LogPath & "log" & (FileCount + 1) & ".log"
    
    ' Add current date to contents
    Set Contents = New Collection
    Contents.Add "Current date: " & CurrentDate
    
    ' Create and open the log file
    FileNum = FreeFile
    Open LogFile For Output As FileNum
    
    ' Write contents to the file
    Dim i As Long
    For i = 1 To Contents.Count
        Print #FileNum, Contents(i)
    Next i
    
    ' Close the file
    Close FileNum
    
    ' Debug print
    Debug.Print "Log file created at: " & LogFile
    Debug.Print "Current date: " & CurrentDate & vbCrLf
End Sub



' Set up main config file
Sub SetupMainConfig(FilePath As String)
    Dim Contents As Collection
    
    ' Create a Shell object
    Dim shellApp As Object                                      ' Shell application object for accessing system paths
    Set shellApp = CreateObject("Shell.Application")
    
    ' Set default path
    Dim DocumentsPath As String
    
    ' Initialize default path
    DocumentsPath = shellApp.Namespace(&H5).Self.Path           ' &H5 corresponds to the Documents folder
    DefaultPath = DocumentsPath & "\Outlook Files\"             ' DEFAULT: Default location for storing backups
    
    ' Initialize the collection for config contents
    Set Contents = New Collection
    
    ' Add configuration items to the collection
    Contents.Add "AutoLaunch=True"                              ' DEFAULT: AutoLaunch setting is true
    Contents.Add "ExecutionDate=" & CurrentDate                 ' DEFAULT: execution date is current date
    Contents.Add "ExecutionPeriod=30"                           ' DEFAULT: execution period is 30
    Contents.Add "ArchivePath=" & DefaultPath
    Contents.Add "RetainmentPeriod=180"                         ' DEFAULT: retainment period is 6 months
    
    ' Write the contents to the file
    Call WriteFile(FilePath, Contents)
    
    ' DEBUG
    LogDebug "Created new config file " & FilePath & " with initial settings:"
    For i = 1 To Contents.Count
        LogDebug " - " & Contents(i)
    Next i
    Debug.Print ""
End Sub



' Check if AutoLaunch is enabled
Function LaunchAutomatically() As Boolean
    ' Extract AutoLaunch Setting
    Dim AutoLaunchString As String
    AutoLaunchString = GetSetting(ConfigContents, "AutoLaunch")
    
    ' Check AutoLaunch setting
    If AutoLaunchString = "True" Then
        LaunchAutomatically = True
    Else
        LaunchAutomatically = False
    End If
    
    ' DEBUG
    LogDebug "AutoLaunch: " & CStr(LaunchAutomatically) & vbCrLf
End Function



' Check if mail can be archived
Function IsArchiveDue() As Boolean
    Dim ExecutionDate As String              ' Date of the last archive execution
    Dim ExecutionPeriod As Long              ' Period in days for running the script
    
    ' Extract execution period and date
    ExecutionDate = GetSetting(ConfigContents, "ExecutionDate")
    ExecutionPeriod = GetSetting(ConfigContents, "ExecutionPeriod")
    
    ' Calculate days since execution
    DaysSinceExecution = DiffDates(ExecutionDate, CurrentDate)
    
    ' DEBUG
    LogDebug "Current Date: " & CurrentDate
    LogDebug "Execution Date: " & ExecutionDate
    LogDebug "Execution Period: " & ExecutionPeriod
    LogDebug "Days Since Execution: " & DaysSinceExecution & vbCrLf
    
    ' Check if more than the predefined days have passed
    If DaysSinceExecution > ExecutionPeriod Then
        IsArchiveDue = True
    Else
        IsArchiveDue = False
    End If
    
    ' DEBUG
    LogDebug "IsArchiveDue: " & CStr(IsArchiveDue) & vbCrLf
End Function



