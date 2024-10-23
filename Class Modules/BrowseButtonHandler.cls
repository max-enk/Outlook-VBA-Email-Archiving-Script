VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BrowseButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class module: BrowseButtonHandler



' Attributes
Public WithEvents BrowseButton As MSForms.CommandButton     ' Event intiation
Attribute BrowseButton.VB_VarHelpID = -1
Public TextBoxPath As MSForms.TextBox                       ' Path attribute



' Event handler for the browse button
Private Sub BrowseButton_Click()
    Dim FolderPath As String
    
    ' Get folder path from file dialog
    FolderPath = BrowseForFolder("Select Folder:")
    
    ' Check for valid new path (maintains old one otherwise)
    If FolderPath <> vbNullString Then
        TextBoxPath.Text = FolderPath
    End If
End Sub



' File Dialog
Private Function BrowseForFolder(ByVal dialogTitle As String) As String
    Dim shellApp As Object
    Dim Folder As Object
    
    ' Create a Shell Application object
    Set shellApp = CreateObject("Shell.Application")
    
    ' Open the folder picker dialog
    Set Folder = shellApp.BrowseForFolder(0, dialogTitle, &H1)
    
    ' If a folder was selected, return the folder's path, otherwise return an empty string
    If Not Folder Is Nothing Then
        BrowseForFolder = Folder.Self.Path
    Else
        BrowseForFolder = vbNullString
    End If
End Function

