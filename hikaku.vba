Option Explicit

'===========================================================
'【Entry Point】
' Perform the following checks on the selected folder:
' - Folder name check (recursive)
' - File content check (Tab delimited and UTF-8, line feed code must be CRLF)
' - File name check (matches the string in the A column of the CorrespondingSheet)
Sub RunBasicValidationChecks()
    Dim fd As FileDialog
    Dim folderPath As String
   
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Please select the folder to be checked. チェック対象のフォルダを選択してください"
    If fd.Show <> -1 Then
        MsgBox "Folder not selected. フォルダが選択されませんでした", vbExclamation
        Exit Sub
    End If
    folderPath = fd.SelectedItems(1)
   
    ' Folder name check
    ValidateFolderNames folderPath
   
    ' File content check & File name check
    ValidateFilesInFolderForFolder folderPath
End Sub

'=============================================
' Folder Name Validation Function
'=============================================
Function ValidateFolderName(folderName As String) As String
    Dim errorMsg As String
    errorMsg = ""
   
    ' Check if the folder name starts with "IF_"
    If Left(folderName, 3) <> "IF_" Then
        errorMsg = errorMsg & "Folder name must start with 'IF_'." & vbCrLf
    End If
   
    ' Split by underscore and check if it consists of three parts ("IF_", timestamp, ItemID)
    Dim parts() As String
    parts = Split(folderName, "_")
    If UBound(parts) <> 2 Then
        errorMsg = errorMsg & "Folder name must consist of three parts ('IF_', timestamp, ItemID) separated by underscores." & vbCrLf
        ValidateFolderName = errorMsg
        Exit Function
    End If
   
    Dim ts As String, itemId As String
    ts = parts(1)
    itemId = parts(2)
   
    ' Check if the timestamp part is 14 digits long
    If Len(ts) <> 14 Then
        errorMsg = errorMsg & "Timestamp must be 14 digits long (current length: " & Len(ts) & ")." & vbCrLf
    Else
        Dim yearPart As String, monthPart As String, dayPart As String
        Dim hourPart As String, minutePart As String, secondPart As String
       
        yearPart = Mid(ts, 1, 4)
        monthPart = Mid(ts, 5, 2)
        dayPart = Mid(ts, 7, 2)
        hourPart = Mid(ts, 9, 2)
        minutePart = Mid(ts, 11, 2)
        secondPart = Mid(ts, 13, 2)
       
        ' Check if each part is numeric
        If Not IsNumeric(yearPart) Then errorMsg = errorMsg & "Year part (" & yearPart & ") is not numeric." & vbCrLf
        If Not IsNumeric(monthPart) Then errorMsg = errorMsg & "Month part (" & monthPart & ") is not numeric." & vbCrLf
        If Not IsNumeric(dayPart) Then errorMsg = errorMsg & "Day part (" & dayPart & ") is not numeric." & vbCrLf
        If Not IsNumeric(hourPart) Then errorMsg = errorMsg & "Hour part (" & hourPart & ") is not numeric." & vbCrLf
        If Not IsNumeric(minutePart) Then errorMsg = errorMsg & "Minute part (" & minutePart & ") is not numeric." & vbCrLf
        If Not IsNumeric(secondPart) Then errorMsg = errorMsg & "Second part (" & secondPart & ") is not numeric." & vbCrLf
       
        ' Range checks
        Dim monthVal As Integer, dayVal As Integer, hourVal As Integer
        Dim minuteVal As Integer, secondVal As Integer
       
        monthVal = CInt(monthPart)
        dayVal = CInt(dayPart)
        hourVal = CInt(hourPart)
        minuteVal = CInt(minutePart)
        secondVal = CInt(secondPart)
       
        If monthVal < 1 Or monthVal > 12 Then
            errorMsg = errorMsg & "Month part (" & monthPart & ") is out of range (1-12)." & vbCrLf
        End If
        If dayVal < 1 Or dayVal > 31 Then
            errorMsg = errorMsg & "Day part (" & dayPart & ") is out of range (1-31)." & vbCrLf
        End If
        If hourVal < 0 Or hourVal > 23 Then
            errorMsg = errorMsg & "Hour part (" & hourPart & ") is out of range (0-23)." & vbCrLf
        End If
        If minuteVal < 0 Or minuteVal > 59 Then
            errorMsg = errorMsg & "Minute part (" & minutePart & ") is out of range (0-59)." & vbCrLf
        End If
        If secondVal < 0 Or secondVal > 59 Then
            errorMsg = errorMsg & "Second part (" & secondPart & ") is out of range (0-59, e.g., '73' is invalid)." & vbCrLf
        End If
    End If
   
    ' Validate the ItemID part: must start with 2-3 letters, followed by an 8-digit number
    Dim letterPart As String, numberPart As String
    Dim i As Integer
    For i = 1 To Len(itemId)
        Dim ch As String
        ch = Mid(itemId, i, 1)
        If ch Like "[A-Za-z]" Then
            letterPart = letterPart & ch
        Else
            Exit For
        End If
    Next i
   
    Dim letterCount As Integer
    letterCount = Len(letterPart)
    If letterCount <> 2 And letterCount <> 3 Then
        errorMsg = errorMsg & "ItemID alphabet part must be 2 or 3 letters (currently: " & letterCount & " letter(s))." & vbCrLf
    End If
   
    numberPart = Mid(itemId, letterCount + 1)
    If Len(numberPart) <> 8 Or Not IsNumeric(numberPart) Then
        errorMsg = errorMsg & "ItemID number part must be an 8-digit number (current: " & numberPart & ")." & vbCrLf
    End If
   
    ValidateFolderName = errorMsg
End Function

'=============================================
' Subroutine to record error information
'=============================================
Sub WriteFolderError(fullPath As String, folderName As String, errorDetails As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("FolderNameError")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "FolderNameError"
        ws.Cells(1, 1).value = "Full Path"
        ws.Cells(1, 2).value = "Folder Name"
        ws.Cells(1, 3).value = "Error Details"
    End If
   
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(lastRow, 1).value = fullPath
    ws.Cells(lastRow, 2).value = folderName
    ws.Cells(lastRow, 3).value = errorDetails
End Sub

'=============================================
' Process to recursively check folders and subfolders
'=============================================
Sub CheckFolderAndSubFolders(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
   
    Dim currentFolder As Object
    Set currentFolder = fso.GetFolder(folderPath)
   
    Dim currentFolderName As String
    currentFolderName = currentFolder.Name
   
    Dim errorMsg As String
    errorMsg = ValidateFolderName(currentFolderName)
   
    ' Record the error if any exist
    If errorMsg <> "" Then
        WriteFolderError currentFolder.Path, currentFolderName, errorMsg
    End If
   
    Dim subFolder As Object
    For Each subFolder In currentFolder.SubFolders
        CheckFolderAndSubFolders subFolder.Path
    Next subFolder
End Sub

'=============================================
' Validate all folder names starting from the specified root folder
' (Excludes the root folder itself)
'=============================================
Sub ValidateFolderNames(ByVal rootFolderPath As String)
   
    ' Get the root folder object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim rootFolder As Object
    Set rootFolder = fso.GetFolder(rootFolderPath)
   
    ' Process only the subfolders of the selected root folder
    Dim subFolder As Object
    For Each subFolder In rootFolder.SubFolders
        CheckFolderAndSubFolders subFolder.Path
    Next subFolder
   
    ' Adjust the column widths in the error log sheet to fit the contents
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("FolderNameError")
    On Error GoTo 0
    If Not ws Is Nothing Then
        ws.Columns("A:C").AutoFit
    End If
   
End Sub

'=============================================
' File content check & File name check
'=============================================
' Perform the following checks on the files in the selected folder
' and output the results to a dedicated sheet.
'
' ＜Check items＞
' 1. File content check (UTF-8(BOM), CRLF, tab delimited)
' If there is an error in any of the files, output as
' “File name : Error contents (comma delimited if there are multiple errors)”
' in the sheet “Contents check”.
' ・If all cases are OK, output as “File content check completed"
'
' 2. File name check (exact match with string in column A of CorrespondingSheet)
' ・If there are any discrepancies, a list is output to the sheet “FileNameError
' ・If all cases are OK, output “File name check completed"
Sub ValidateFilesInFolderForFolder(ByVal folderPath As String)
    Dim fso As Object, folderObj As Object, fileObj As Object
    Dim allowedNames As Collection
    Dim fileContentErrors As Collection
    Dim invalidFilenames As Collection
   
    ' Get the list of acceptable file names from column A of CorrespondingSheet
    Set allowedNames = GetAllowedFileNames()
    If allowedNames Is Nothing Then Exit Sub
   
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderObj = fso.GetFolder(folderPath)
   
    Set fileContentErrors = New Collection
    Set invalidFilenames = New Collection
   
    For Each fileObj In folderObj.Files
        Dim errMsg As String
        errMsg = ValidateFileContent(fileObj.Path)
        If errMsg <> "" Then
            fileContentErrors.Add fileObj.Name & " : " & errMsg
        End If
       
        If Not IsFileNameAllowed(fileObj.Name, allowedNames) Then
            invalidFilenames.Add fileObj.Name
        End If
    Next fileObj
   
    ' 1.Output file content check results to the “ContentCheck” sheet.
    Dim wsContent As Worksheet
    Set wsContent = GetOrCreateSheet("ContentCheck")
    wsContent.Cells.Clear
    Dim r As Long
    r = 1
    If fileContentErrors.Count = 0 Then
        wsContent.Cells(r, 1).value = "File content check completed. ファイル内容チェック完了"
    Else
        wsContent.Cells(r, 1).value = "List of file content errors. ファイル内容エラー一覧"
        r = r + 1
        Dim errItem As Variant
        For Each errItem In fileContentErrors
            wsContent.Cells(r, 1).value = errItem
            r = r + 1
        Next errItem
    End If
   
    '2.Output file name check results to “FileNameError” sheet
    Dim wsName As Worksheet
    Set wsName = GetOrCreateSheet("FileNameError")
    wsName.Cells.Clear
    r = 1
    If invalidFilenames.Count = 0 Then
        wsName.Cells(r, 1).value = "File name check completed. ファイル名チェック完了"
    Else
        wsName.Cells(r, 1).value = "List of incorrect file names. ファイル名間違い一覧"
        r = r + 1
        Dim fn As Variant
        For Each fn In invalidFilenames
            wsName.Cells(r, 1).value = fn
            r = r + 1
        Next fn
    End If
End Sub

'=============================================
' Helper function for sheet acquisition
'=============================================
' If the sheet with the specified name exists, it is returned; otherwise, a new sheet is created and returned.
Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

'=============================================
' Obtaining acceptable file names from CorrespondingSheet
'=============================================
' Returns the strings listed in column A of the sheet “CorrespondingSheet” as a collection.
Function GetAllowedFileNames() As Collection
    Dim allowed As New Collection
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
   
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CorrespondingSheet")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet 'CorrespondingSheet' not found", vbExclamation
        Set GetAllowedFileNames = Nothing
        Exit Function
    End If
   
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 1 To lastRow
        Dim cellVal As String
        cellVal = Trim(ws.Cells(r, "A").value)
        If cellVal <> "" Then
            allowed.Add cellVal
        End If
    Next r
    Set GetAllowedFileNames = allowed
End Function

' Check if the file name is included in the allowed list
Function IsFileNameAllowed(fileName As String, allowedNames As Collection) As Boolean
    Dim allowedName As Variant
    For Each allowedName In allowedNames
         If fileName = allowedName Then
             IsFileNameAllowed = True
             Exit Function
         End If
    Next allowedName
    IsFileNameAllowed = False
End Function

'=============================================
' File Content Check
'=============================================
' Performs each of the following checks on the specified file
' and returns a comma-separated list of error messages if any.
' (returns an empty string if all is OK)
Function ValidateFileContent(filePath As String) As String
    Dim errorsList() As String
    Dim n As Long: n = 0
    Dim tempMsg As String
   
    tempMsg = CheckUTF8Encoding(filePath)
    If tempMsg <> "" Then
        ReDim Preserve errorsList(n)
        errorsList(n) = tempMsg
        n = n + 1
    End If
   
    tempMsg = CheckCRLF(filePath)
    If tempMsg <> "" Then
        ReDim Preserve errorsList(n)
        errorsList(n) = tempMsg
        n = n + 1
    End If
   
    tempMsg = CheckTabDelimited(filePath)
    If tempMsg <> "" Then
        ReDim Preserve errorsList(n)
        errorsList(n) = tempMsg
        n = n + 1
    End If
   
    If n = 0 Then
        ValidateFileContent = ""
    Else
        ValidateFileContent = Join(errorsList, ",")
    End If
End Function

'-----------------------------------------------------------
' UTF-8 (BOM) check: check if the first 3 bytes of the file are EF BB BF
Function CheckUTF8Encoding(filePath As String) As String
    Dim stm As Object, bytes() As Byte
    On Error GoTo ErrHandler
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary mode
    stm.Open
    stm.LoadFromFile filePath
    If stm.size < 3 Then
         CheckUTF8Encoding = "BOM check not possible due to file size too small. ファイルサイズが小さすぎるためBOMチェック不可"
         stm.Close: Set stm = Nothing: Exit Function
    End If
    bytes = stm.Read(3)
    stm.Close: Set stm = Nothing
    If bytes(0) <> &HEF Or bytes(1) <> &HBB Or bytes(2) <> &HBF Then
         CheckUTF8Encoding = "UTF-8 BOM does not exist. UTF-8 BOMが存在しない"
    Else
         CheckUTF8Encoding = ""
    End If
    Exit Function
ErrHandler:
    CheckUTF8Encoding = "UTF-8 Check Error: " & Err.Description
End Function

'-----------------------------------------------------------
' CRLF check: Binary read to check if CR (13) immediately precedes each LF (10)
Function CheckCRLF(filePath As String) As String
    Dim stm As Object, bytes() As Byte, i As Long, size As Long
    On Error GoTo ErrHandler
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary mode
    stm.Open
    stm.LoadFromFile filePath
    size = stm.size
    If size = 0 Then
         CheckCRLF = "empty file"
         stm.Close: Set stm = Nothing: Exit Function
    End If
    bytes = stm.Read(size)
    stm.Close: Set stm = Nothing
    For i = 0 To UBound(bytes)
       If bytes(i) = 10 Then ' In case of LF
            If i = 0 Then
                CheckCRLF = "The beginning of the file is LF. ファイル先頭がLF"
                Exit Function
            Else
                If bytes(i - 1) <> 13 Then
                    CheckCRLF = "LF alone detected, not CRLF. LF単体を検出"
                    Exit Function
                End If
            End If
       End If
    Next i
    CheckCRLF = ""
    Exit Function
ErrHandler:
    CheckCRLF = "CRLF check error: " & Err.Description
End Function

'-----------------------------------------------------------
' Tab delimiter check: read in text mode and check if each line contains a tab character
' *In case of an error, it displays which line did not find a tab.
Function CheckTabDelimited(filePath As String) As String
    Dim stm As Object
    Dim fileText As String, lines() As String
    Dim i As Long
    On Error GoTo ErrHandler
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text mode
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    fileText = stm.ReadText(-1)
    stm.Close: Set stm = Nothing
   
    lines = Split(fileText, vbCrLf)
    For i = LBound(lines) To UBound(lines)
       If Trim(lines(i)) <> "" Then
          If InStr(lines(i), vbTab) = 0 Then
              ' The array starts at 0, so the row number for the user is (i + 1)
              CheckTabDelimited = "A non-tab-delimited line exists.タブ区切りでない行を検出 (Applicable lines: " & (i + 1)
              Exit Function
          End If
       End If
    Next i
    CheckTabDelimited = ""
    Exit Function
ErrHandler:
    CheckTabDelimited = "Tab delimiter check error: " & Err.Description
End Function

