Option Explicit

Sub CheckIFDataValue()
    Dim wsThisFile As Worksheet  ' ˆ—‘ÎÛ‚Ìƒ[ƒNƒV[ƒg
    Dim headerRow As Range       ' ƒwƒbƒ_[s‚Ì”ÍˆÍ
    Dim headerCell As Range      ' ƒwƒbƒ_[ƒZƒ‹
    Dim cellIFDataValue As Variant  ' ŠeƒZƒ‹‚Ì’l
    Dim headerType As String     ' ƒwƒbƒ_[‚ÌŒ^
    Dim lastRow As Long          ' ÅŒã‚Ìs”Ô†
    Dim lastCol As Long          ' ÅŒã‚Ì—ñ”Ô†
    Dim i As Long                ' sƒ‹[ƒv—p•Ï”
    Dim j As Long                ' —ñƒ‹[ƒv—p•Ï”
    Dim sheetHasMismatch As Boolean  ' ƒV[ƒg‚Éƒ~ƒXƒ}ƒbƒ`‚ª‚ ‚é‚©‚Ç‚¤‚©‚Ìƒtƒ‰ƒO
    Dim maxLength As Long        ' Å‘å’·
    Dim lengthExceeds As Boolean ' ’·‚³’´‰ß‚Ìƒtƒ‰ƒO
    Dim lovSheet As Worksheet    ' LOVƒV[ƒg
    Dim lovValue As String       ' LOV’l
    Dim lovDict As Object        ' DictionaryƒIƒuƒWƒFƒNƒg
    Dim lovRange As Range        ' LOV‚Ì”ÍˆÍ
    Dim firstAddress As String   ' LOVŒŸõ‚ÌÅ‰‚ÌƒAƒhƒŒƒX
    Dim cellComment As String       'ƒZƒ‹‚É‹LÚ‚³‚ê‚½ƒƒ‚‚ð•Û‘¶‚·‚é—p‚Ì•Ï”

    'ˆ—ŠJŽn‚ÌƒƒOo—Í
    Debug.Print vbCrLf
    Debug.Print Now & " " & "CheckIFDataValue Start"

    ' ˆ—ŠJŽn‚Ìƒ|ƒbƒvƒAƒbƒv‚ð•\Ž¦
    MsgBox "‚±‚ÌExcel“à‚ÌƒV[ƒg‚É’Ç‰Á‚³‚ê‚½‚à‚Ì‚É‚Â‚¢‚ÄTenkai DB IFƒtƒ@ƒCƒ‹‚Ìƒf[ƒ^ƒ`ƒFƒbƒNˆ—‚ðŠJŽn‚µ‚Ü‚·B", vbOKOnly, "ŠJŽn"

    ' ‰æ–ÊXV‚ÆŽ©“®ŒvŽZ‚ð’âŽ~‚µA•s—v‚ÈƒCƒxƒ“ƒg‚ð–³Œø‰»
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False


    ' ‚±‚Ìƒtƒ@ƒCƒ‹‚ÌŠeƒ[ƒNƒV[ƒg‚ðƒ‹[ƒv
    For Each wsThisFile In ThisWorkbook.Worksheets
    
        Debug.Print Now & " " & "Processing sheet:" & wsThisFile.Name
        
        ' ‘ÎÛŠO‚ÌƒV[ƒg‚ðƒXƒLƒbƒv
        If wsThisFile.Name = "Corresponding Sheets" Or wsThisFile.Name = "ƒtƒ@ƒCƒ‹–¼ŠÔˆá‚¢" Or wsThisFile.Name = "LOV_Entity_datamodel" Or wsThisFile.Name = "LOV_Entity_classfn" Then
            
            Debug.Print Now & " " & "Skip"
            GoTo SkipSheet
        
        End If

        sheetHasMismatch = False
        lengthExceeds = False

        With wsThisFile
        
            ' ‚±‚Ìƒtƒ@ƒCƒ‹‚ÌŠeƒ[ƒNƒV[ƒg‚ÌÅŒã‚Ìs‚Æ—ñ‚ðŽæ“¾
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column

            ' ‚±‚Ìƒtƒ@ƒCƒ‹‚ÌŠeƒ[ƒNƒV[ƒg‚Ì3`5s–Ú‚Ìƒwƒbƒ_[s‚ðŽæ“¾
            Set headerRow = .Range(.Cells(3, 1), .Cells(5, lastCol))

            'Šeƒ[ƒNƒV[ƒg‚ÌŒ…”ãŒÀ’l‚ð”z—ñ‚ÉŠi”[
            Dim maxLengths() As Long
            ReDim maxLengths(1 To lastCol)
            For j = 1 To lastCol
                
                '5s–Ú‚ÌŒ…”‚ðŠi”[
                If IsNumeric(.Cells(5, j).value) And Not IsEmpty(.Cells(5, j).value) Then
                    maxLengths(j) = CLng(.Cells(5, j).value)
                
                '5s–Ú‚ª”Žš‚Å‚Í‚È‚¢ê‡A4s–Ú‚ð”Žš‚ðŠi”[@¦Classification.txt‚È‚Ç‚ðl—¶
                ElseIf IsNumeric(.Cells(4, j).value) And Not IsEmpty(.Cells(4, j).value) Then
                    maxLengths(j) = CLng(.Cells(4, j).value)
                    
                '‚»‚êˆÈŠO‚Ìê‡‚Í0‚ðŠi”[
                Else
                    maxLengths(j) = 0
                End If
            Next j

            ' ‚±‚Ìƒtƒ@ƒCƒ‹‚ÌŠeƒ[ƒNƒV[ƒg–¼‚Ìæ“ª‚É"(PtCl)"‚Ü‚½‚Í"(DcCl)"‚ªŠÜ‚Ü‚ê‚é‚©Šm”F
            If Left(wsThisFile.Name, 6) = "(PtCl)" Or Left(wsThisFile.Name, 6) = "(DcCl)" Then
                ' LOV_Entity_classfnƒV[ƒg‚ðŽQÆ
                Set lovSheet = ThisWorkbook.Worksheets("LOV_Entity_classfn")

                ' ‚±‚Ìƒtƒ@ƒCƒ‹‚ÌŠeƒ[ƒNƒV[ƒg‚Ì6s–Ú‚ÌŠeƒJƒ‰ƒ€‚ðƒ`ƒFƒbƒN
                For j = 1 To lastCol
                    ' Dictionary‚ð‰Šú‰»
                    Set lovDict = CreateObject("Scripting.Dictionary")
                    
                    lovValue = .Cells(6, j).value
                    ' lovValue‚ª‹ó”’‚Å‚È‚­A‚©‚Â"LOV"‚ðŠÜ‚ÝA"No LOV"‚ðŠÜ‚Ü‚È‚¢ê‡‚É‚Ì‚Ýˆ—‚ðŽÀs
                    If Not IsEmpty(lovValue) And InStr(lovValue, "LOV") > 0 And InStr(lovValue, "No LOV") = 0 Then
                        ' "LOV:"‚Ü‚½‚Í"LOV :"‚Ì•¶Žš—ñ‚ðŠÜ‚ÞƒZƒ‹‚Ì’l‚ðŽæ“¾‚µA"LOV:"‚Ü‚½‚Í"LOV :"‚ÌŒã‚Ì•¶Žš—ñ‚ð’Šo
                        If InStr(lovValue, "LOV:") > 0 Then
                            lovValue = Trim(Mid(lovValue, InStr(lovValue, "LOV:") + 4))
                        ElseIf InStr(lovValue, "LOV :") > 0 Then
                            lovValue = Trim(Mid(lovValue, InStr(lovValue, "LOV :") + 5))
                        End If
                        
                        ' LOVƒV[ƒg‚Ì2—ñ–Ú‚Å’Šo‚µ‚½•¶Žš—ñ‚ðŒŸõ
                        Set lovRange = lovSheet.Columns(2).Find(lovValue)
                        If Not lovRange Is Nothing Then
                            ' ŒŸõ‚ÌÅ‰‚ÌƒAƒhƒŒƒX‚ð‹L‰¯
                            firstAddress = lovRange.Address
                            
                            ' LOVƒV[ƒg‚ÌD—ñ‚ÆE—ñ‚Ì’l‚ðÅ‰‚©‚çÅŒã‚Ü‚ÅŽæ“¾‚µADictionary‚ÉŠi”[
                            Do
                                lovDict(lovRange.Offset(0, 2).value) = True ' D—ñ‚ðDictionary‚É’Ç‰Á
                                lovDict(lovRange.Offset(0, 3).value) = True ' E—ñ‚ðDictionary‚É’Ç‰Á
                                ' ŽŸ‚ÌŒŸõŒ‹‰Ê‚ðŽæ“¾
                                Set lovRange = lovSheet.Columns(2).FindNext(lovRange)
                                ' ŽŸ‚ÌŒŸõŒ‹‰Ê‚ªNothing‚Å‚È‚¢A‚©‚ÂÅ‰‚ÌƒAƒhƒŒƒX‚ÆˆÙ‚È‚éŠÔƒ‹[ƒv‚ðŒp‘±
                            Loop While Not lovRange Is Nothing And lovRange.Address <> firstAddress
                            
                            ' ‚±‚Ìƒtƒ@ƒCƒ‹‚ÌŠeƒ[ƒNƒV[ƒg‚Ì7s–Ú‚Ü‚½‚Í8s–ÚˆÈ~‚ðƒ`ƒFƒbƒN
                            For i = 7 To lastRow
                                If Left(wsThisFile.Name, 5) = "(DcCl)" Then
                                    i = 8
                                End If
                                cellIFDataValue = .Cells(i, j).value
                                ' ‚±‚Ìƒtƒ@ƒCƒ‹‚ÌŠeƒ[ƒNƒV[ƒg‚Ì7s–Ú‚Ü‚½‚Í8s–ÚˆÈ~‚ÌƒZƒ‹‚Ì’l‚ªDictionary‚É‘¶Ý‚µ‚È‚¢ê‡AƒZƒ‹‚ÌƒtƒHƒ“ƒg‚ð‘¾Žš‚É‚·‚é
                                If Not IsEmpty(cellIFDataValue) And Not IsNull(cellIFDataValue) And Len(Trim(cellIFDataValue)) > 0 And Not lovDict.Exists(cellIFDataValue) Then
                                    
                                    'ƒZƒ‹‚ð‰©F‚É•ÏX
                                    .Cells(i, j).Interior.Color = vbYellow
                                    
                                    'ƒJƒ‰ƒ€ƒwƒbƒ_[‚ð‰©F‚É•ÏX
                                    .Cells(1, j).Interior.Color = vbYellow
                                    .Cells(2, j).Interior.Color = vbYellow
                                    .Cells(3, j).Interior.Color = vbYellow
                                    
                                    'ƒRƒƒ“ƒg•Û‘¶—p‚Ì•Ï”‚ð‰Šú‰»
                                    cellComment = ""
                                    
                                    '‚·‚Å‚ÉƒRƒƒ“ƒg‚ª‹L“ü‚³‚ê‚Ä‚¢‚È‚¢‚©‚ðŠm”F
                                    If TypeName(.Cells(i, j).Comment) = "Comment" Then
                                        
                                        '‚·‚Å‚ÉƒRƒƒ“ƒg‚ª‚ ‚Á‚½ê‡‚ÍƒRƒƒ“ƒg‚ÌƒoƒbƒNƒAƒbƒv‚ð•Û‘¶
                                        cellComment = .Cells(i, j).Comment.Text
                                        
                                        'ƒRƒƒ“ƒg‚ðƒNƒŠƒA
                                        .Cells(i, j).ClearComments
                                        
                                        'ƒoƒbƒNƒAƒbƒv‚µ‚½ƒRƒƒ“ƒg‚Æ‚Æ‚à‚É‹LÚ
                                        .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LOVnotFound(classification)" & vbCrLf & cellComment
                                        
                                    'ƒRƒƒ“ƒg‚ª‚È‚©‚Á‚½ê‡
                                    Else
                                        
                                        '‚»‚Ì‚Ü‚ÜƒRƒƒ“ƒg‚ð‹LÚ
                                        .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LOVnotFound(classification)"
                                        
                                    End If
                                    
                                    'ƒRƒƒ“ƒg‚ÌƒTƒCƒY‚ðŽ©“®’²®
                                    .Cells(i, j).Comment.Shape.TextFrame.AutoSize = True
                                    
                                    'ƒV[ƒgƒ^ƒu‚ÌF‚ð‰©F‚É‚·‚éƒtƒ‰ƒO‚ð—§‚Ä‚é
                                    sheetHasMismatch = True
                                    
                                End If
                            Next i
                        End If
                    End If
                Next j
            Else
                ' LOV_Entity_datamodelƒV[ƒg‚ðŽQÆ
                Set lovSheet = ThisWorkbook.Worksheets("LOV_Entity_datamodel")

                ' ‚±‚Ìƒtƒ@ƒCƒ‹‚ÌŠeƒ[ƒNƒV[ƒg‚Ì7s–Ú‚ÌŠeƒJƒ‰ƒ€‚ðƒ`ƒFƒbƒN
                For j = 1 To lastCol
                    ' Dictionary‚ð‰Šú‰»
                    Set lovDict = CreateObject("Scripting.Dictionary")
                    
                    lovValue = .Cells(7, j).value
                    ' lovValue‚ª‹ó”’‚Å‚È‚­A‚©‚Â"LOV"‚ðŠÜ‚ÝA"No LOV"‚ðŠÜ‚Ü‚È‚¢ê‡‚É‚Ì‚Ýˆ—‚ðŽÀs
                    If Not IsEmpty(lovValue) And InStr(lovValue, "LOV") > 0 And InStr(lovValue, "No LOV") = 0 Then
                        ' "LOV:"‚Ü‚½‚Í"LOV :"‚Ì•¶Žš—ñ‚ðŠÜ‚ÞƒZƒ‹‚Ì’l‚ðŽæ“¾‚µA"LOV:"‚Ü‚½‚Í"LOV :"‚ÌŒã‚Ì•¶Žš—ñ‚ð’Šo
                        If InStr(lovValue, "LOV:") > 0 Then
                            lovValue = Trim(Mid(lovValue, InStr(lovValue, "LOV:") + 4))
                        ElseIf InStr(lovValue, "LOV :") > 0 Then
                            lovValue = Trim(Mid(lovValue, InStr(lovValue, "LOV :") + 5))
                        End If
                        
                        ' LOVƒV[ƒg‚Ì2—ñ–Ú‚Å’Šo‚µ‚½•¶Žš—ñ‚ðŒŸõ
                        Set lovRange = lovSheet.Columns(2).Find(lovValue)
                        If Not lovRange Is Nothing Then
                            ' ŒŸõ‚ÌÅ‰‚ÌƒAƒhƒŒƒX‚ð‹L‰¯
                            firstAddress = lovRange.Address
                            
                            ' LOVƒV[ƒg‚ÌD—ñ‚ÆE—ñ‚Ì’l‚ðÅ‰‚©‚çÅŒã‚Ü‚ÅŽæ“¾‚µADictionary‚ÉŠi”[
                            Do
                                lovDict(Trim(lovRange.Offset(0, 2).value)) = True ' D—ñ‚ðDictionary‚É’Ç‰Á
                                lovDict(Trim(lovRange.Offset(0, 3).value)) = True ' E—ñ‚ðDictionary‚É’Ç‰Á
                                ' ŽŸ‚ÌŒŸõŒ‹‰Ê‚ðŽæ“¾
                                Set lovRange = lovSheet.Columns(2).FindNext(lovRange)
                                ' ŽŸ‚ÌŒŸõŒ‹‰Ê‚ªNothing‚Å‚È‚¢A‚©‚ÂÅ‰‚ÌƒAƒhƒŒƒX‚ÆˆÙ‚È‚éŠÔƒ‹[ƒv‚ðŒp‘±
                            Loop While Not lovRange Is Nothing And lovRange.Address <> firstAddress
                            
                            ' ‚±‚Ìƒtƒ@ƒCƒ‹‚ÌŠeƒ[ƒNƒV[ƒg‚Ì8s–ÚˆÈ~‚ðƒ`ƒFƒbƒN
                            For i = 8 To lastRow
                                cellIFDataValue = .Cells(i, j).value
                                ' ‚±‚Ìƒtƒ@ƒCƒ‹‚ÌŠeƒ[ƒNƒV[ƒg‚Ì8s–ÚˆÈ~‚ÌƒZƒ‹‚Ì’l‚ªDictionary‚É‘¶Ý‚µ‚È‚¢ê‡AƒZƒ‹‚ÌƒtƒHƒ“ƒg‚ð‘¾Žš‚É‚·‚é
                                If Not IsEmpty(cellIFDataValue) And Not IsNull(cellIFDataValue) And Len(Trim(cellIFDataValue)) > 0 And Not lovDict.Exists(cellIFDataValue) Then
                                    
                                    'ƒZƒ‹‚ð‰©F‚É•ÏX
                                    .Cells(i, j).Interior.Color = vbYellow
                                    
                                    'ƒJƒ‰ƒ€ƒwƒbƒ_[‚ð‰©F‚É•ÏX
                                    .Cells(1, j).Interior.Color = vbYellow
                                    .Cells(2, j).Interior.Color = vbYellow
                                    .Cells(3, j).Interior.Color = vbYellow
                                    
                                    'ƒRƒƒ“ƒg•Û‘¶—p‚Ì•Ï”‚ð‰Šú‰»
                                    cellComment = ""
                                    
                                    '‚·‚Å‚ÉƒRƒƒ“ƒg‚ª‹L“ü‚³‚ê‚Ä‚¢‚È‚¢‚©‚ðŠm”F
                                    If TypeName(.Cells(i, j).Comment) = "Comment" Then
                                        
                                        '‚·‚Å‚ÉƒRƒƒ“ƒg‚ª‚ ‚Á‚½ê‡‚ÍƒRƒƒ“ƒg‚ÌƒoƒbƒNƒAƒbƒv‚ð•Û‘¶
                                        cellComment = .Cells(i, j).Comment.Text
                                        
                                        'ƒRƒƒ“ƒg‚ðƒNƒŠƒA
                                        .Cells(i, j).ClearComments
                                        
                                        'ƒoƒbƒNƒAƒbƒv‚µ‚½ƒRƒƒ“ƒg‚Æ‚Æ‚à‚É‹LÚ
                                        .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LOVnotFound(datamodel)" & vbCrLf & cellComment
                                        
                                    'ƒRƒƒ“ƒg‚ª‚È‚©‚Á‚½ê‡
                                    Else
                                        
                                        '‚»‚Ì‚Ü‚ÜƒRƒƒ“ƒg‚ð‹LÚ
                                        .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LOVnotFound(datamodel)"
                                        
                                    End If
                                    
                                    'ƒRƒƒ“ƒg‚ÌƒTƒCƒY‚ðŽ©“®’²®
                                    .Cells(i, j).Comment.Shape.TextFrame.AutoSize = True
                                    
                                    'ƒV[ƒgƒ^ƒu‚ÌF‚ð‰©F‚É‚·‚éƒtƒ‰ƒO‚ð—§‚Ä‚é
                                    sheetHasMismatch = True
                                
                                End If
                            Next i
                        End If
                    End If
                Next j
            End If

            ' ‚±‚Ìƒtƒ@ƒCƒ‹‚ÌŠeƒ[ƒNƒV[ƒg‚Ì7s–ÚˆÈ~‚ÌŠeƒZƒ‹‚ðƒ`ƒFƒbƒN
            For i = 7 To lastRow
                For j = 1 To lastCol
                
                     ' ŠeƒZƒ‹‚Ì’l‚ðŽæ“¾
                    cellIFDataValue = .Cells(i, j).value
                    
                    ' ‹ó”’ƒZƒ‹‚ÍƒXƒLƒbƒv
                    If cellIFDataValue = "" Or IsEmpty(cellIFDataValue) Then
                        GoTo NextCell
                    End If

                    ' ŠeƒZƒ‹‚Ìƒwƒbƒ_[Fƒf[ƒ^Œ^‚ðŽæ“¾
                    Set headerCell = headerRow.Cells(2, j)
                    
                    '‘OŒã‚Ì‹ó”’‚ðíœ
                    headerCell = Trim(headerCell)
                    
                    'Žæ“¾‚µ‚½ƒwƒbƒ_[Fƒf[ƒ^Œ^‚ªString‚È‚Ç‚Ì•¶Žš—ñ‚Å‚ ‚Á‚½ê‡‚É‚Í“Á‚Éˆ—‚È‚µ
                    If headerCell = "Integer" Or headerCell = "Double" Or headerCell = "Decimal" Or headerCell = "Date" Or headerCell = "Boolean" Or headerCell = "String" Or headerCell = "Numeric" Or headerCell = "" Then
                        
                    'Žæ“¾‚µ‚½ƒwƒbƒ_[Fƒf[ƒ^Œ^‚ªString‚È‚Ç‚Ì•¶Žš—ñ‚Å‚Í‚È‚©‚Á‚½ê‡
                    Else
                    
                        '1s–Ú‚ðŽæ“¾@¦Classification.txt‚È‚Ç‚ðl—¶
                        Set headerCell = headerRow.Cells(1, j)
                    
                    End If

                    ' ƒwƒbƒ_[‚ÌŒ^‚ðŽæ“¾
                    headerType = CStr(headerCell.value)

                    ' ƒf[ƒ^‚ÌŒ^‚ªˆê’v‚µ‚È‚¢ê‡
                    If Not CheckDataType(cellIFDataValue, headerType) Then
                        
                        'ƒZƒ‹‚ð‰©F‚É•ÏX
                        .Cells(i, j).Interior.Color = vbYellow
                        
                        'ƒJƒ‰ƒ€ƒwƒbƒ_[‚ð‰©F‚É•ÏX
                        .Cells(1, j).Interior.Color = vbYellow
                        .Cells(2, j).Interior.Color = vbYellow
                        .Cells(3, j).Interior.Color = vbYellow
                        
                        'ƒRƒƒ“ƒg•Û‘¶—p‚Ì•Ï”‚ð‰Šú‰»
                        cellComment = ""
                        
                        '‚·‚Å‚ÉƒRƒƒ“ƒg‚ª‹L“ü‚³‚ê‚Ä‚¢‚È‚¢‚©‚ðŠm”F
                        If TypeName(.Cells(i, j).Comment) = "Comment" Then
                            
                            '‚·‚Å‚ÉƒRƒƒ“ƒg‚ª‚ ‚Á‚½ê‡‚ÍƒRƒƒ“ƒg‚ÌƒoƒbƒNƒAƒbƒv‚ð•Û‘¶
                            cellComment = .Cells(i, j).Comment.Text
                            
                            'ƒRƒƒ“ƒg‚ðƒNƒŠƒA
                            .Cells(i, j).ClearComments
                            
                            'ƒoƒbƒNƒAƒbƒv‚µ‚½ƒRƒƒ“ƒg‚Æ‚Æ‚à‚É‹LÚ
                            .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "DataTypeUnmatch" & vbCrLf & cellComment
                            
                        'ƒRƒƒ“ƒg‚ª‚È‚©‚Á‚½ê‡
                        Else
                            
                            '‚»‚Ì‚Ü‚ÜƒRƒƒ“ƒg‚ð‹LÚ
                            .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "DataTypeUnmatch"
                            
                        End If
                        
                        'ƒRƒƒ“ƒg‚ÌƒTƒCƒY‚ðŽ©“®’²®
                        .Cells(i, j).Comment.Shape.TextFrame.AutoSize = True
                        
                        'ƒV[ƒgƒ^ƒu‚ÌF‚ð‰©F‚É‚·‚éƒtƒ‰ƒO‚ð—§‚Ä‚é
                        sheetHasMismatch = True
                    
                    End If

                    ' ƒoƒCƒg”§ŒÀ‚ðƒ`ƒFƒbƒN
                    If maxLengths(j) > 0 And GetByteLength(CStr(cellIFDataValue)) > maxLengths(j) Then
                        
                        'ƒZƒ‹‚ð‰©F‚É•ÏX
                        .Cells(i, j).Interior.Color = vbYellow
                        
                        'ƒJƒ‰ƒ€ƒwƒbƒ_[‚ð‰©F‚É•ÏX
                        .Cells(1, j).Interior.Color = vbYellow
                        .Cells(2, j).Interior.Color = vbYellow
                        .Cells(3, j).Interior.Color = vbYellow
                        
                        'ƒRƒƒ“ƒg•Û‘¶—p‚Ì•Ï”‚ð‰Šú‰»
                        cellComment = ""
                        
                        '‚·‚Å‚ÉƒRƒƒ“ƒg‚ª‹L“ü‚³‚ê‚Ä‚¢‚È‚¢‚©‚ðŠm”F
                        If TypeName(.Cells(i, j).Comment) = "Comment" Then
                            
                            '‚·‚Å‚ÉƒRƒƒ“ƒg‚ª‚ ‚Á‚½ê‡‚ÍƒRƒƒ“ƒg‚ÌƒoƒbƒNƒAƒbƒv‚ð•Û‘¶
                            cellComment = .Cells(i, j).Comment.Text
                            
                            'ƒRƒƒ“ƒg‚ðƒNƒŠƒA
                            .Cells(i, j).ClearComments
                            
                            'ƒoƒbƒNƒAƒbƒv‚µ‚½ƒRƒƒ“ƒg‚Æ‚Æ‚à‚É‹LÚ
                            .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LengthExceed" & vbCrLf & cellComment
                            
                        'ƒRƒƒ“ƒg‚ª‚È‚©‚Á‚½ê‡
                        Else
                            
                            '‚»‚Ì‚Ü‚ÜƒRƒƒ“ƒg‚ð‹LÚ
                            .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LengthExceed"
                            
                        End If
                        
                        'ƒRƒƒ“ƒg‚ÌƒTƒCƒY‚ðŽ©“®’²®
                        .Cells(i, j).Comment.Shape.TextFrame.AutoSize = True
                        
                        'ƒV[ƒgƒ^ƒu‚ÌF‚ð‰©F‚É‚·‚éƒtƒ‰ƒO‚ð—§‚Ä‚é
                        lengthExceeds = True
                        
                    End If
NextCell:
                Next j
            Next i

            ' ƒGƒ‰[‚ªŠÜ‚Ü‚ê‚éƒV[ƒgƒ^ƒu‚ÌF‚ð‰©F‚É‚·‚é
            If sheetHasMismatch Or lengthExceeds Then
                wsThisFile.Tab.Color = vbYellow
            End If
        End With
SkipSheet:
    Next wsThisFile

    ' ‰æ–ÊXV‚ÆŽ©“®ŒvŽZ‚ðÄŠJ‚µA•s—v‚ÈƒCƒxƒ“ƒg‚Ì–³Œø‰»‚ð‰ðœ
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    'ˆ—ŠJŽn‚ÌƒƒOo—Í
    Debug.Print Now & " " & "CheckIFDataValue Finish"
    Debug.Print vbCrLf

    'ˆê”Ôæ“ª‚ÌƒV[ƒg‚ÖˆÚ“®
    Worksheets(1).Select
    
    ' ˆ—I—¹‚Ìƒ|ƒbƒvƒAƒbƒv‚ð•\Ž¦
    MsgBox "TenkaiDB IFƒtƒ@ƒCƒ‹‚Ìƒf[ƒ^ƒ`ƒFƒbƒNˆ—‚ªŠ®—¹‚µ‚Ü‚µ‚½B" & vbCrLf & "Žd—l‚ÆˆÙ‚È‚Á‚Ä‚¢‚éƒf[ƒ^‚É‚Â‚¢‚Ä‚ÍƒZƒ‹‚ª‰©F‚Å•\Ž¦‚³‚ê‚Ä‚¨‚èAƒƒ‚‚É‚»‚Ì——R‚ª‹LÚ‚³‚ê‚Ä‚¢‚Ü‚·B‚²Šm”F‰º‚³‚¢B", vbOKOnly, "Š®—¹"
End Sub

' Œ^‚Ìˆê’v‚ðƒ`ƒFƒbƒN‚·‚éŠÖ”
Function CheckDataType(ByVal value As Variant, ByVal dataType As String) As Boolean
    On Error GoTo ErrorHandler

    ' ’l‚ª‹ó‚Ìê‡‚Íˆê’v‚Æ‚Ý‚È‚·
    If IsEmpty(value) Then
        CheckDataType = True
    Else
        ' ƒf[ƒ^Œ^‚É‰ž‚¶‚½”»’è‚ðs‚¤
        Select Case dataType
            Case "Integer"
                ' ®”Œ^‚Ì”»’è
                CheckDataType = IsNumeric(value) And (CLng(value) = value) And (InStr(CStr(value), ".") = 0)
            Case "Double"
                CheckDataType = IsNumeric(value) And (CDbl(value) = value)
            Case "Decimal"
                ' ¬”Œ^‚Ì”»’è
                CheckDataType = IsNumeric(value) And (CDbl(value) = value)
            Case "Date"
                ' “ú•tŒ^‚Ì”»’è
                If IsNumeric(value) = True Then
                    CheckDataType = False
                Else
                    CheckDataType = IsDate(value)
                End If
            Case "Boolean"
                ' ƒu[ƒ‹Œ^‚Ì”»’è
                CheckDataType = (LCase(CStr(value)) = "true" Or LCase(CStr(value)) = "false")
            Case "String"
                ' •¶Žš—ñŒ^‚Ì”»’è
                ' ‚à‚µ’l‚ª•¶Žš—ñ‚Å "+" ‚Æ "." ‚ðŠÜ‚ñ‚Å‚¢‚ÄA+‚Æ.‚ðœ‹Ž‚µ‚½ê‡‚É”Žš”»’è‚³‚ê‚éê‡A”Ž®‚Ìˆê•”‚ÆŒ©‚È‚³‚ê‚é‚½‚ßA•¶Žš—ñ‚Æ‚µ‚Ä‚Í•s³‚Æ‚·‚éB
                If InStr(CStr(value), "+") > 0 And InStr(CStr(value), ".") > 0 And IsNumeric(Replace(Replace(value, "+", ""), ".", "")) = True Then
                    CheckDataType = False
                ' ‚à‚µ’l‚ª“ú•tŒ`Ž®‚Å "/" ‚ðŠÜ‚ñ‚Å‚¢‚éê‡A“ú•t‚Æ‚µ‚Ä”FŽ¯‚³‚ê‚é‚½‚ßA•¶Žš—ñ‚Æ‚µ‚Ä‚Í•s³‚Æ‚·‚éB
                ElseIf IsDate(CStr(value)) = True And InStr(CStr(value), "/") > 0 Then
                    CheckDataType = False
                ' ‚à‚µ’l‚ª "true" ‚Ü‚½‚Í "false" ‚Ì•¶Žš—ñ‚Å‚ ‚éê‡Aƒu[ƒ‹Œ^‚ÆŒ©‚È‚³‚ê‚é‚½‚ßA•¶Žš—ñ‚Æ‚µ‚Ä‚Í•s³‚Æ‚·‚éB
                ElseIf (LCase(CStr(value)) = "true" Or LCase(CStr(value)) = "false") = True Then
                    CheckDataType = False
                ' ‚à‚µ’l‚ª”’l‚Å‚ ‚éê‡A”’lŒ^‚ÆŒ©‚È‚³‚ê‚é‚½‚ßA•¶Žš—ñ‚Æ‚µ‚Ä‚Í•s³‚Æ‚·‚éB@¦ƒ`ƒFƒbƒN‚ÌƒmƒCƒY‚ª‘½‚­‚È‚é‚½‚ßA‚±‚ê‚ðOFF‚É‚µ‚Ä‚à‚æ‚¢‚©‚àc—vŠm”Fš
                'ElseIf IsNumeric(value) = True Then
                    'CheckDataType = False
                ' ã‹L‚Ì‚Ç‚ÌðŒ‚É‚àŠY“–‚µ‚È‚¢ê‡A’l‚Í•¶Žš—ñ‚Æ‚µ‚Ä—LŒø‚Æ‚·‚éB
                Else
                    CheckDataType = True
                End If
            Case "Numeric"
                ' ”’lŒ^‚Ì”»’è
                CheckDataType = IsNumeric(value)
            Case Else
                ' ‚»‚Ì‘¼‚ÌŒ^‚Ìê‡‚Íˆê’v‚Æ‚Ý‚È‚·
                CheckDataType = True
        End Select
    End If

    Exit Function

ErrorHandler:
    ' ƒGƒ‰[‚ª”­¶‚µ‚½ê‡‚Íˆê’v‚µ‚È‚¢‚Æ‚Ý‚È‚·
    CheckDataType = False
End Function

' ƒoƒCƒg”‚ðŒvŽZ‚·‚éŠÖ”
Function GetByteLength(str As String) As Long
    Dim i As Long
    Dim byteLength As Long
    Dim charCode As Long

    byteLength = 0

    ' •¶Žš—ñ‚ÌŠe•¶Žš‚É‚Â‚¢‚ÄƒoƒCƒg”‚ðŒvŽZ
    For i = 1 To Len(str)
        charCode = AscW(Mid(str, i, 1))
        ' •¶Žš‚ÌUnicode’l‚ðŽæ“¾
        If charCode <= 127 Then
            ' 1ƒoƒCƒg•¶ŽšiASCII•¶Žšj‚Ìê‡
            byteLength = byteLength + 1
        Else
            ' 3ƒoƒCƒg•¶Žši”ñASCII•¶Žšj‚Ìê‡
            byteLength = byteLength + 3
        End If
    Next i

    GetByteLength = byteLength
End Function
