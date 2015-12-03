
'功能：搜尋想要的檔案，在某個資料夾內。並且將搜尋到的結果放入一個Collection
'傳入：path 想要搜尋的資料夾
'      format 想要搜尋的副檔名
'      result 搜尋到的檔案
'回傳：null(透過byRef方式，直接存入result collection)
Sub findFiles(ByVal path As String, ByVal format As String, ByRef result As Collection)

    'save all the directory
    Dim findDirectory As Collection
    Set findDirectory = New Collection
    
    'use to find all files
    Dim strFile As String
    strFile = Dir(path & "\", vbDirectory Or vbNormal Or vbReadOnly)
    
    'find all files and directorys
    While strFile <> ""
        DoEvents
        'save all directorys into findDirectory
        If (GetAttr(path & "\" & strFile) = vbDirectory) And strFile <> "." And strFile <> ".." Then
            findDirectory.Add (path & "\" & strFile)
        ElseIf Right(strFile, Len(format)) = format Then
            result.Add path & "\" & strFile
            
        End If
        strFile = Dir
    Wend

    'find sub-directory
    If findDirectory.Count <> 0 Then
        Dim currectDir As Variant
        For Each currectDir In findDirectory
            Call findFiles(currectDir, format, result)
        Next
        'free the tmp collection
        Set currectDir = Nothing
    End If

    'free the collection
    Set findDirectory = Nothing

End Sub


'功能：選擇一個目錄，並且返回所選擇的完整路徑
'傳入：null
'回傳：完整路徑
Private Function choosePath() As String
    Dim fs As FileDialog
    Set fs = Application.FileDialog(msoFileDialogFolderPicker)
    fs.Show
    choosePath = fs.SelectedItems(1)
End Function
