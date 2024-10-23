VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArchiveSettings 
   Caption         =   "AutoArchive Settings"
   ClientHeight    =   4824
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6504
   OleObjectBlob   =   "ArchiveSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ArchiveSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Private variables
'' Parameters
Private AutoLaunch As Boolean                           ' AutoLaunch value
Private ExecutionPeriod As Integer                      ' Frequency of archive prompts
Private ArchivePath As String                           ' Default archive path
Private RetainmentPeriod As Integer                     ' Default period for retaining mail
'' Other
Private BrowseArchiveFolder As BrowseButtonHandler      ' Browse button handler
Private ExecDays As Collection                          ' Value of the execution period in days
Private ExecLabels As Collection                        ' Value of the execution period in user-readable form
Private ExecIndex As Integer                            ' Index of the execution period
Private RetainDays As Collection                        ' Value of the retain period in days
Private RetainLabels As Collection                      ' Value of the retain period in user-readable form
Private RetainIndex As Integer                          ' Index of the retain period



Private Sub LabelArchive_Click()

End Sub

' Initialize the UserForm
Private Sub UserForm_Initialize()
    ' Variable Declaration
    '' Loop Variables
    Dim i As Integer                            ' Loop counter for iterating through collections
    '' Parameters
    Dim AutoLaunchString As String              ' AutoLaunch from file
    Dim ExecutionPeriodString As String         ' ExecutionPeriod from file
    Dim ArchivePathString As String             ' ArchivePath from file
    Dim RetainmentPeriodString As String        ' RetainmentPeriod from file
    
    
    ' Variable definition
    '' ExecutionPeriod options
    Set ExecDays = New Collection
    Set ExecLabels = New Collection
    ''' Fill ExecDays with the execution period in days
    ExecDays.Add 1        ' Everyday
    ExecDays.Add 3        ' 3 Days
    ExecDays.Add 7        ' 1 Week
    ExecDays.Add 14       ' 2 Weeks
    ExecDays.Add 21       ' 3 Weeks
    ExecDays.Add 30       ' 1 Month
    ExecDays.Add 60       ' 2 Months
    ExecDays.Add 90       ' 3 Months
    ExecDays.Add 180      ' 6 Months
    ExecDays.Add 270      ' 9 Months
    ExecDays.Add 365      ' 1 Year
    ''' Fill ExecLabels with user-readable descriptions
    ExecLabels.Add "Everyday"
    ExecLabels.Add "Every 3 Days"
    ExecLabels.Add "1 Week"
    ExecLabels.Add "2 Weeks"
    ExecLabels.Add "3 Weeks"
    ExecLabels.Add "1 Month"
    ExecLabels.Add "2 Months"
    ExecLabels.Add "3 Months"
    ExecLabels.Add "6 Months"
    ExecLabels.Add "9 Months"
    ExecLabels.Add "1 Year"
    ''' Populate the ComboBoxExec with the execution period options
    For i = 1 To ExecLabels.Count
        ComboBoxExec.AddItem ExecLabels(i)
    Next i
    
    '' RetainmenPeriod options
    Set RetainDays = New Collection
    Set RetainLabels = New Collection
    ''' Fill RetainDays with the date value
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
    ''' Fill RetainLabels with the date
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
    ''' Populate the combobox with retain period options
    For i = 1 To RetainLabels.Count
        ComboBoxRetain.AddItem RetainLabels(i)
    Next i
    
    
    ' Extract and set parameters from ConfigContents
    '' AutoLaunch
    AutoLaunchString = GetSetting(ConfigContents, "AutoLaunch")
    If AutoLaunchString = "True" Then
        AutoLaunch = True
    ElseIf AutoLaunchString = "False" Then
        AutoLaunch = False
    Else
        ' Read error
        Call ReadError("AutoLaunch")
    End If
    
    '' ExecutionPeriod
    ExecutionPeriodString = GetSetting(ConfigContents, "ExecutionPeriod")
    If IsInt(ExecutionPeriodString) Then
        ExecutionPeriod = CInt(ExecutionPeriodString)
    Else
        ' Read error
        Call ReadError("ExecutionPeriod")
    End If

    '' ArchivePath
    ArchivePathString = GetSetting(ConfigContents, "ArchivePath")
    If IsValidPath(ArchivePathString) Then
        ArchivePath = ArchivePathString
        
        ' Check if the directory exists
        If Dir(ArchivePath, vbDirectory) = "" Then
            ' Create the directory if it does not exist
            MkDir ArchivePath
            
            LogDebug "Created directory " & ArchivePath & vbCrLf
        End If
    Else
        ' Read error
        Call ReadError("ArchivePath")
    End If
    
    '' RetainmentPeriod
    RetainmentPeriodString = GetSetting(ConfigContents, "RetainmentPeriod")
    If IsInt(RetainmentPeriodString) Then
        RetainmentPeriod = CInt(RetainmentPeriodString)
    Else
        ' Read error
        Call ReadError("RetainmentPeriod")
    End If


    ' Fill form
    '' AutoLaunch
    CheckBoxAutoLaunch.value = AutoLaunch
    
    '' ExecutionPeriod
    ExecIndex = GetIndex(ExecDays, ExecutionPeriod)
    If ExecIndex >= 0 Then
        ComboBoxExec.ListIndex = ExecIndex
    Else
        ' Set default to Everyday
        ComboBoxExec.ListIndex = 0
        LogDebug "ExecutionPeriod value from file is invalid. Using default value."
    End If
    
    '' ArchivePath
    TextBoxArchivePath.Text = ArchivePath
    
    '' RetainmentPeriod
    RetainIndex = GetIndex(RetainDays, RetainmentPeriod)
    If RetainIndex >= 0 Then
        ComboBoxRetain.ListIndex = RetainIndex
    Else
        ' Set default to 6 months
        ComboBoxRetain.ListIndex = 7
        LogDebug "RetainmentPeriod value from file is invalid. Using default value."
    End If
    
    
    ' BrowseButton functionality
    Set BrowseArchiveFolder = New BrowseButtonHandler
    Set BrowseArchiveFolder.BrowseButton = ButtonArchivePath
    Set BrowseArchiveFolder.TextBoxPath = TextBoxArchivePath
    
    
    ' Formatting
    LabelStartUp.Font.Bold = True
    LabelArchive.Font.Bold = True
End Sub



' Event handler for the Continue button
Private Sub ButtonSave_Click()
    ' Variable declaration
    '' Flags
    Dim Changed As Boolean          ' Track changes
    Dim Valid As Boolean            ' Valid settings
    '' Box values
    Dim BoxLaunch As Boolean        ' AutoLaunch from button
    Dim BoxExecIndex As Integer     ' ExecutionPeriod from combobox
    Dim BoxPath As String           ' ArchivePath from textbox
    Dim BoxRetainIndex As Integer   ' RetainmentPeriod from combobox

    ' Variable definition
    Changed = False
    Valid = True
    
    
    ' Check changes in settings
    '' AutoLaunch
    BoxLaunch = CheckBoxAutoLaunch.value
    If Not BoxLaunch = AutoLaunch Then
        Changed = True
        
        ' Update the AutoLaunch value in settings
        Call ReplaceSetting(ConfigContents, "AutoLaunch", BoxLaunch)
    End If
    
    '' ExecutionPeriod
    BoxExecIndex = ComboBoxExec.ListIndex
    If BoxExecIndex <> ExecIndex Then
        Changed = True
        
        ' Update ExecutionPeriod
        ExecutionPeriod = ExecDays(BoxExecIndex + 1)
        
        ' Update the ExecutionPeriod value in settings
        Call ReplaceSetting(ConfigContents, "ExecutionPeriod", ExecutionPeriod)
    End If

    '' ArchivePath
    BoxPath = TextBoxArchivePath.Text
    ''' Check if the path is valid
    If IsValidPath(BoxPath) Then
        ' Check if the value has changed
        If Not BoxPath = ArchivePath Then
            Changed = True
            
            If Right(BoxPath, 1) <> "\" Then
                BoxPath = BoxPath & "\"
            End If
            
            ' Update the ArchivePath in settings
            Call ReplaceSetting(ConfigContents, "ArchivePath", BoxPath)
            
            ' Check if the old directory is empty (contains neither files nor subfolders)
            If IsEmptyDirectory(ArchivePath) Then
                ' Delete the old directory if empty
                RmDir ArchivePath
                
                LogDebug "Deleted directory " & ArchivePath & vbCrLf
            End If
        End If
    Else
        Valid = False
        
        ' If conversion fails, show a message and reset the value
        MsgBox "Please enter a valid archive path.", vbExclamation, "Invalid Input."
        
        TextBoxArchivePath.Text = ArchivePath  ' Reset to previous value
    End If
    
    '' RetainmentPeriod
    BoxRetainIndex = ComboBoxRetain.ListIndex
    If BoxRetainIndex <> RetainIndex Then
        Changed = True
        
        ' Update RetainmentPeriod
        RetainmentPeriod = RetainDays(BoxRetainIndex + 1)
        
        ' Update the RetainmentPeriod value in settings
        Call ReplaceSetting(ConfigContents, "RetainmentPeriod", RetainmentPeriod)
    End If
    
    
    ' Update the config file
    If Changed And Valid Then
        Call WriteFile(ConfigFile, ConfigContents)
    End If
    
    
    ' Close form
    If Valid Then
        Me.Hide
    End If
End Sub



' Event handler for the Cancel button
Private Sub ButtonClose_Click()
    Me.Hide
End Sub



' Handler for config file read error
Private Sub ReadError(ByVal setting As String)
        MsgBox "Error reading setting '" & setting & "' in the configuration file:" & vbCrLf & _
            ConfigFile & vbCrLf & vbCrLf & _
            "Please check your setup and try again.", _
            vbExclamation, "Error Reading Configuration File"

        Me.Hide
End Sub

