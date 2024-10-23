VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoLaunchPrompt 
   Caption         =   "AutoArchive Setup"
   ClientHeight    =   1440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5784
   OleObjectBlob   =   "AutoLaunchPrompt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AutoLaunchPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Initialize the UserForm
Private Sub UserForm_Initialize()
    StartLabel.Caption = "Mail has not been archived in " & DaysSinceExecution & " days. Run AutoArchive now?"
End Sub



' Event handler for the Continue button
Private Sub ButtonContinue_Click()
    RunNow = True
    Me.Hide
End Sub



' Event handler for the Cancel button
Private Sub ButtonCancel_Click()
    Me.Hide
End Sub

