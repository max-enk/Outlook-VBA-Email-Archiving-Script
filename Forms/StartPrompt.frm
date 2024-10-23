VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartPrompt 
   Caption         =   "Run AutoArchive"
   ClientHeight    =   1440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5784
   OleObjectBlob   =   "StartPrompt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StartPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Initialize the UserForm
Private Sub UserForm_Initialize()
    
End Sub



' Event handler for the Continue button
Private Sub ButtonContinue_Click()
    Me.Hide
End Sub



' Event handler for the Cancel button
Private Sub ButtonCancel_Click()
    ' Set continue to false
    Continue = False
    
    ' Hide the UserForm
    Me.Hide
End Sub


