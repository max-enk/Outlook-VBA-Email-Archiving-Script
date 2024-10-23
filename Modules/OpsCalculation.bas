Attribute VB_Name = "OpsCalculation"
' Calculate the absolute difference between two dates
Function DiffDates(DateString1 As String, DateString2 As String) As Long
    Dim date1 As Date
    Dim date2 As Date
    Dim dayPart As String
    Dim monthPart As String
    Dim yearPart As String
    Dim parts1() As String
    Dim parts2() As String
    Dim daysDifference As Long
    
    ' Split the first DateString by the delimiter "/"
    parts1 = Split(DateString1, "/")
    
    ' Extract day, month, and year from parts1
    dayPart = parts1(0)
    monthPart = parts1(1)
    yearPart = parts1(2)
    
    ' Construct the first Date value
    date1 = DateSerial(CInt(yearPart), CInt(monthPart), CInt(dayPart))
    
    ' Split the second DateString by the delimiter "/"
    parts2 = Split(DateString2, "/")
    
    ' Extract day, month, and year from parts2
    dayPart = parts2(0)
    monthPart = parts2(1)
    yearPart = parts2(2)
    
    ' Construct the second Date value
    date2 = DateSerial(CInt(yearPart), CInt(monthPart), CInt(dayPart))
    
    ' Calculate the absolute difference in days
    daysDifference = Abs(DateDiff("d", date1, date2))
    
    ' Return the absolute difference in days
    DiffDates = daysDifference
End Function



' Delay program by set amount of time in milliseconds
Sub Delay(ms As Long)
    Dim startTime As Single
    Dim delayTime As Single
    
    ' Convert delay time from milliseconds to seconds
    delayTime = ms / 1000
    
    ' Get the current time in seconds since midnight
    startTime = Timer

    ' Loop until the delay time has passed
    Do While Timer < startTime + delayTime
        DoEvents
    Loop
End Sub



' Checks if the input string is a valid integer
Function IsInt(value As String) As Boolean
    Dim trimmedValue As String
    trimmedValue = Trim(value)  ' Remove any leading or trailing spaces
    
    ' Check if the string represents a valid number
    If IsNumeric(trimmedValue) Then
        ' Check if value contains decimal places
        If InStr(trimmedValue, ".") > 0 Or InStr(trimmedValue, ",") > 0 Then
            IsInt = False
        Else
            IsInt = True
        End If
    Else
        IsInt = False
    End If
End Function



' Checks if the input string is a valid path
Function IsValidPath(ByVal Path As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Regular expression pattern for a valid Windows path
    ' This checks for paths like: C:\Folder\Subfolder or \\Server\Folder\Subfolder
    With regEx
        .Pattern = "^[a-zA-Z]:\\(?:[^\\/:;*?""<>=|]+\\)*[^\\/:;*?""<>=|]*$"
        .IgnoreCase = True
        .Global = False
    End With
    
    ' Test the path against the pattern
    IsValidPath = regEx.Test(Path)
End Function



' Function to check if a directory is empty
Function IsEmptyDirectory(ByVal Path As String) As Boolean
    Dim FileOrFolder As String
    
    ' Get the first file or folder (includes both files and directories)
    FileOrFolder = Dir(Path & "*", vbDirectory)
    
    ' Loop through all files and folders in the directory
    Do While FileOrFolder <> ""
        ' Ignore the current and parent directory ('.' and '..')
        If FileOrFolder <> "." And FileOrFolder <> ".." Then
            ' Check if it's a directory
            If (GetAttr(Path & FileOrFolder) And vbDirectory) = vbDirectory Then
                ' It's a directory
                IsEmptyDirectory = False
                Exit Function
            Else
                ' It's a file
                IsEmptyDirectory = False
                Exit Function
            End If
        End If
        
        ' Get the next file or folder
        FileOrFolder = Dir
    Loop
    
    ' If no files or subfolders are found, the directory is empty
    IsEmptyDirectory = True
End Function




' Formats folder path by removing \\name from its start
Function FormatFolderPath(ByVal Path As String, ByVal name As String) As String
    FormatFolderPath = Replace(Path, "\\" & name, "", vbTextCompare)
End Function
