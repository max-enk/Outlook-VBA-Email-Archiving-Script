Attribute VB_Name = "OpsCollection"
' Find a line starting with a specified prefix
Function GetEntry(col As Collection, Prefix As String) As String
    Dim entry As Variant
    
    ' Initialize the return value as an empty string
    GetEntry = ""
    
    ' Iterate through each item in the collection
    For Each entry In col
        ' Check if the line starts with the specified prefix
        If Left(entry, Len(Prefix)) = Prefix Then
            GetEntry = entry
            Exit Function
        End If
    Next entry
End Function



' Get a setting for a parameter from a collection
Function GetSetting(col As Collection, Prefix As String) As String
    Dim entry As String
    Dim parts() As String
    
    ' Find the entry with the given prefix
    entry = GetEntry(col, Prefix)
    
    ' Check if the entry is not empty
    If entry <> "" Then
        ' Split the entry using the "=" delimiter
        parts = Split(entry, "=")
        
        ' Check if the split resulted in at least two parts (prefix and value)
        If UBound(parts) >= 1 Then
            ' Return the value part (the second part after the "=")
            GetSetting = parts(1)
        Else
            ' Return an empty string if the entry is malformed
            GetSetting = ""
        End If
    Else
        ' Return an empty string if the line is not found
        GetSetting = ""
    End If
End Function



' Get a folder setting for a parameter from a collection
Function GetFolderSettings(col As Collection, Prefix As String) As Collection
    Dim entry As String
    Dim parts() As String
    Dim result As Collection
    
    Set result = New Collection
    
    ' Find the entry with the given prefix
    entry = GetEntry(col, Prefix & "|")
    
    ' Check if the entry is not empty
    If entry <> "" Then
        ' Split the entry using the pipe delimiter
        parts = Split(entry, "|")
        
        ' Add all parts except the first one to the result collection
        Dim i As Integer
        For i = 1 To UBound(parts)
            result.Add parts(i)
        Next i
    End If
    
    ' Return the result collection
    Set GetFolderSettings = result
End Function



' Gets index from collection of entry matching string
Function GetIndex(ByVal col As Collection, ByVal searchString As String) As Integer
    Dim i As Integer
    
    ' Initialize the function to return -1 in case no match is found
    GetIndex = -1
    
    ' Loop through the collection to find the matching entry
    For i = 1 To col.Count
        If col(i) = searchString Then
            GetIndex = i - 1
            Exit Function
        End If
    Next i
End Function



' Replace a setting for a specific parameter
Function ReplaceSetting(ByRef col As Collection, ByVal Prefix As String, ByVal NewValue As String)
    Dim i As Integer
    Dim tempArray() As String
    Dim entry As String
    
    ' Allocate an array of the same size as the collection
    ReDim tempArray(1 To col.Count)
    
    ' Copy collection contents to array for manipulation
    For i = 1 To col.Count
        tempArray(i) = col(i)
    Next i
    
    ' Replace entry in the array if the prefix matches
    For i = 1 To UBound(tempArray)
        entry = tempArray(i)
        If Left(entry, Len(Prefix)) = Prefix Then
            tempArray(i) = Prefix & "=" & NewValue
            Exit For
        End If
    Next i
    
    ' Clear the collection and re-add the modified array contents
    Set col = New Collection
    For i = 1 To UBound(tempArray)
        col.Add tempArray(i)
    Next i
End Function

