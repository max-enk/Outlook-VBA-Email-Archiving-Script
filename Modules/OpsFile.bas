Attribute VB_Name = "OpsFile"
' Read specified file into collection
Function ReadFile(FilePath As String) As Collection
    ' Variable Declaration
    Dim FileNum As Integer
    Dim fileLine As String
    Dim fileContents As Collection

    ' Initialize Collection
    Set fileContents = New Collection

    ' Open the file
    FileNum = FreeFile
    Open FilePath For Input As FileNum

    ' Read each line from the file and add it to the Collection
    Do While Not EOF(FileNum)
        Line Input #FileNum, fileLine
        fileContents.Add fileLine
    Loop

    ' Close the file
    Close FileNum

    ' Return the Collection
    Set ReadFile = fileContents
    
    ' DEBUG
    LogDebug "Read from file " & FilePath & vbCrLf
End Function



' Write collection to file
Function WriteFile(FilePath As String, fileContents As Collection)
    ' Variable Declaration
    Dim FileNum As Integer
    Dim i As Long
    
    ' Open the file for writing
    FileNum = FreeFile
    Open FilePath For Output As FileNum

    ' Loop through the collection and write each item to the file
    For i = 1 To fileContents.Count
        Print #FileNum, fileContents(i)
    Next i

    ' Close the file
    Close FileNum
    
    ' DEBUG
    LogDebug "Wrote to file " & FilePath & vbCrLf
End Function



' Function to count files in a directory
Function CountFilesInDirectory(ByVal DirectoryPath As String) As Long
    Dim FileCount As Long
    Dim FileName As String
    
    FileName = Dir(DirectoryPath & "\*.*") ' Get the first file
    While FileName <> ""
        FileCount = FileCount + 1 ' Increment file count
        FileName = Dir() ' Get the next file
    Wend
    
    CountFilesInDirectory = FileCount
End Function



' Function to check if a file is in use by attempting to open it
Function IsFileInUse(FileName As String) As Boolean
    Dim FileNumber As Integer
    FileNumber = FreeFile ' Get a free file number

    On Error Resume Next
    Open FileName For Binary Access Read Lock Read As #FileNumber
    
    ' File is not in use
    If Err.Number = 0 Then
        Close #FileNumber
        IsFileInUse = False
        
        ' DEBUG
        ' Debug.Print "File " & FileName & " is not in use." & vbCrLf
    ' File is in use
    Else
        IsFileInUse = True
        
        ' DEBUG
        ' Debug.Print "File " & FileName & " is still in use. Err: " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0
End Function



' Append to log file
Sub LogDebug(DebugMessage As String)
    Dim FileNum As Integer
    
    ' Open the log file for appending
    FileNum = FreeFile
    Open LogFile For Append As FileNum
    
    ' Write the debug message to the log file
    If DebugMessage <> "" Then Print #FileNum, Replace(DebugMessage, vbCrLf, " ")
    
    ' Close the log file
    Close FileNum
    
    ' Print the debug message in the console
    Debug.Print DebugMessage
End Sub
