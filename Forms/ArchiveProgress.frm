VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArchiveProgress 
   Caption         =   "Archive Progress"
   ClientHeight    =   3636
   ClientLeft      =   60
   ClientTop       =   276
   ClientWidth     =   5892
   OleObjectBlob   =   "ArchiveProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ArchiveProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Progress As Double ' Stores progress in %

' Initialize the UserForm
Private Sub UserForm_Initialize()
    ' Set UserForm dimensions
    Me.Width = 250
    Me.Height = 210
    
    ' Initialize LabelAccountName
    LabelAccountName.Font.Bold = True
    LabelAccountName.Caption = "Processing account: " & Account.GetAccountName

    ' Initialize LabelFolderName
    LabelFolderName.Caption = "Archiving mail in folder: " & AccountFolder.name
    
    ' Initialize LabelProgressBar
    LabelProgressBar.BackColor = RGB(0, 255, 0) ' Green color
    LabelProgressBar.Width = 0 ' Set initial width to 0
    
    ' Initialize LabelHeader
    LabelHeader.Font.Bold = True
    
    ' Initialize LabelProgress
    LabelProgress.Font.Bold = True
    
    ' Initialize LabelPercentage
    LabelPercentage.Font.Bold = True
    LabelPercentage.TextAlign = fmTextAlignRight
    LabelPercentage.Caption = "0%"
End Sub



' Subroutine to handle item processing and progress update
Public Sub ProcessItems()
    Dim TotalItems As Long
    Dim ProcessedItems As Long
    Dim Item As Object
    Dim ProgressBarWidth As Double
    Dim ItemSubject As String
    Dim ItemDate As String
    Dim ItemSender As String

    ' Initialize the progress
    TotalItems = ValidItems.Count
    ProcessedItems = 0
    ProgressBarWidth = FrameProgress.Width ' Maximum width for progress bar

    ' Show the form
    Me.Show vbModeless

    ' Iterate through ValidItems and move them to ArchiveFolder
    For Each Item In ValidItems
        Select Case Item.Class
            ' Mail item
            Case olMail
                ItemSubject = Replace(Replace(Replace(CStr(Item.Subject), vbCrLf, ""), vbCr, ""), vbLf, "")
                ItemSender = Item.SenderName
                ItemDate = Format(Item.ReceivedTime, "dd/mm/yyyy hh:mm")
            ' Report item
            Case olReport
                ItemSubject = Replace(Replace(Replace(CStr(Item.Subject), vbCrLf, ""), vbCr, ""), vbLf, "")
                ItemSender = "Postmaster"
                ItemDate = Format(Item.CreationTime, "dd/mm/yyyy hh:mm")
            ' Meeting item
            Case olMeetingRequest, olMeetingCancellation, olMeetingForwardNotification, _
                 olMeetingResponseNegative, olMeetingResponsePositive, olMeetingResponseTentative
                ItemSubject = Replace(Replace(Replace(CStr(Item.Subject), vbCrLf, ""), vbCr, ""), vbLf, "")
                ItemSender = Item.SenderName
                ItemDate = Format(Item.SentOn, "dd/mm/yyyy hh:mm")
        End Select
        
        ' Update processed items count
        ProcessedItems = ProcessedItems + 1
        
        ' Update the current header label
        LabelHeader.Caption = "Moving mail item (" & CStr(ProcessedItems) & "/" & CStr(TotalItems) & "):"
        
        ' Update the item label for mail or meeting items
        LabelItem.Caption = "Subject: " & ItemSubject & vbCrLf & _
                            "Sender: " & ItemSender & vbCrLf & _
                            "Date: " & ItemDate
        
        ' Check duplicates
        If IsDuplicate(Item) Then
            ' Move the item to the DuplicateFolder
            Item.Move DuplicateFolder
            LogDebug "   - Moved duplicate item: '" & ItemSubject & "' on " & ItemDate
        Else
            ' Move the item to the ArchiveFolder
            Item.Move ArchiveFolder
            LogDebug "   - Moved item: '" & ItemSubject & "' on " & ItemDate
        End If
        
        ' Update progress bar
        Progress = (ProcessedItems / TotalItems) * 100
        LabelProgressBar.Width = (Progress / 100) * ProgressBarWidth
        LabelPercentage.Caption = Format(Progress, "0") & "%"
        DoEvents
        Me.Repaint
        Delay 100

        If Not Continue Then Exit For
    Next Item

    Delay 1000
    Me.Hide
End Sub



' Event handler for the Cancel button
Private Sub ButtonCancel_Click()
    ' Set continue to false
    Continue = False
    
    ' Hide the UserForm
    Me.Hide
End Sub

' Event handler for closing the UserForm via the red "X" button
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Check if the form is being closed by the red "X" button
    If CloseMode = vbFormControlMenu Then
        ' Set Continue to False
        Continue = False
    End If
End Sub






