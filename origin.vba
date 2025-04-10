Option Explicit

Sub CheckIFDataValue()
    Dim wsThisFile As Worksheet  ' 処理対象のワークシート
    Dim headerRow As Range       ' ヘッダー行の範囲
    Dim headerCell As Range      ' ヘッダーセル
    Dim cellIFDataValue As Variant  ' 各セルの値
    Dim headerType As String     ' ヘッダーの型
    Dim lastRow As Long          ' 最後の行番号
    Dim lastCol As Long          ' 最後の列番号
    Dim i As Long                ' 行ループ用変数
    Dim j As Long                ' 列ループ用変数
    Dim sheetHasMismatch As Boolean  ' シートにミスマッチがあるかどうかのフラグ
    Dim maxLength As Long        ' 最大長
    Dim lengthExceeds As Boolean ' 長さ超過のフラグ
    Dim lovSheet As Worksheet    ' LOVシート
    Dim lovValue As String       ' LOV値
    Dim lovDict As Object        ' Dictionaryオブジェクト
    Dim lovRange As Range        ' LOVの範囲
    Dim firstAddress As String   ' LOV検索の最初のアドレス
    Dim cellComment As String       'セルに記載されたメモを保存する用の変数

    '処理開始のログ出力
    Debug.Print vbCrLf
    Debug.Print Now & " " & "CheckIFDataValue Start"

    ' 処理開始のポップアップを表示
    MsgBox "このExcel内のシートに追加されたものについてTenkai DB IFファイルのデータチェック処理を開始します。", vbOKOnly, "開始"

    ' 画面更新と自動計算を停止し、不要なイベントを無効化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False


    ' このファイルの各ワークシートをループ
    For Each wsThisFile In ThisWorkbook.Worksheets
    
        Debug.Print Now & " " & "Processing sheet:" & wsThisFile.Name
        
        ' 対象外のシートをスキップ
        If wsThisFile.Name = "Corresponding Sheets" Or wsThisFile.Name = "ファイル名間違い" Or wsThisFile.Name = "LOV_Entity_datamodel" Or wsThisFile.Name = "LOV_Entity_classfn" Then
            
            Debug.Print Now & " " & "Skip"
            GoTo SkipSheet
        
        End If

        sheetHasMismatch = False
        lengthExceeds = False

        With wsThisFile
        
            ' このファイルの各ワークシートの最後の行と列を取得
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column

            ' このファイルの各ワークシートの3〜5行目のヘッダー行を取得
            Set headerRow = .Range(.Cells(3, 1), .Cells(5, lastCol))

            '各ワークシートの桁数上限値を配列に格納
            Dim maxLengths() As Long
            ReDim maxLengths(1 To lastCol)
            For j = 1 To lastCol
                
                '5行目の桁数を格納
                If IsNumeric(.Cells(5, j).value) And Not IsEmpty(.Cells(5, j).value) Then
                    maxLengths(j) = CLng(.Cells(5, j).value)
                
                '5行目が数字ではない場合、4行目を数字を格納　※Classification.txtなどを考慮
                ElseIf IsNumeric(.Cells(4, j).value) And Not IsEmpty(.Cells(4, j).value) Then
                    maxLengths(j) = CLng(.Cells(4, j).value)
                    
                'それ以外の場合は0を格納
                Else
                    maxLengths(j) = 0
                End If
            Next j

            ' このファイルの各ワークシート名の先頭に"(PtCl)"または"(DcCl)"が含まれるか確認
            If Left(wsThisFile.Name, 6) = "(PtCl)" Or Left(wsThisFile.Name, 6) = "(DcCl)" Then
                ' LOV_Entity_classfnシートを参照
                Set lovSheet = ThisWorkbook.Worksheets("LOV_Entity_classfn")

                ' このファイルの各ワークシートの6行目の各カラムをチェック
                For j = 1 To lastCol
                    ' Dictionaryを初期化
                    Set lovDict = CreateObject("Scripting.Dictionary")
                    
                    lovValue = .Cells(6, j).value
                    ' lovValueが空白でなく、かつ"LOV"を含み、"No LOV"を含まない場合にのみ処理を実行
                    If Not IsEmpty(lovValue) And InStr(lovValue, "LOV") > 0 And InStr(lovValue, "No LOV") = 0 Then
                        ' "LOV:"または"LOV :"の文字列を含むセルの値を取得し、"LOV:"または"LOV :"の後の文字列を抽出
                        If InStr(lovValue, "LOV:") > 0 Then
                            lovValue = Trim(Mid(lovValue, InStr(lovValue, "LOV:") + 4))
                        ElseIf InStr(lovValue, "LOV :") > 0 Then
                            lovValue = Trim(Mid(lovValue, InStr(lovValue, "LOV :") + 5))
                        End If
                        
                        ' LOVシートの2列目で抽出した文字列を検索
                        Set lovRange = lovSheet.Columns(2).Find(lovValue)
                        If Not lovRange Is Nothing Then
                            ' 検索の最初のアドレスを記憶
                            firstAddress = lovRange.Address
                            
                            ' LOVシートのD列とE列の値を最初から最後まで取得し、Dictionaryに格納
                            Do
                                lovDict(lovRange.Offset(0, 2).value) = True ' D列をDictionaryに追加
                                lovDict(lovRange.Offset(0, 3).value) = True ' E列をDictionaryに追加
                                ' 次の検索結果を取得
                                Set lovRange = lovSheet.Columns(2).FindNext(lovRange)
                                ' 次の検索結果がNothingでない、かつ最初のアドレスと異なる間ループを継続
                            Loop While Not lovRange Is Nothing And lovRange.Address <> firstAddress
                            
                            ' このファイルの各ワークシートの7行目または8行目以降をチェック
                            For i = 7 To lastRow
                                If Left(wsThisFile.Name, 5) = "(DcCl)" Then
                                    i = 8
                                End If
                                cellIFDataValue = .Cells(i, j).value
                                ' このファイルの各ワークシートの7行目または8行目以降のセルの値がDictionaryに存在しない場合、セルのフォントを太字にする
                                If Not IsEmpty(cellIFDataValue) And Not IsNull(cellIFDataValue) And Len(Trim(cellIFDataValue)) > 0 And Not lovDict.Exists(cellIFDataValue) Then
                                    
                                    'セルを黄色に変更
                                    .Cells(i, j).Interior.Color = vbYellow
                                    
                                    'カラムヘッダーを黄色に変更
                                    .Cells(1, j).Interior.Color = vbYellow
                                    .Cells(2, j).Interior.Color = vbYellow
                                    .Cells(3, j).Interior.Color = vbYellow
                                    
                                    'コメント保存用の変数を初期化
                                    cellComment = ""
                                    
                                    'すでにコメントが記入されていないかを確認
                                    If TypeName(.Cells(i, j).Comment) = "Comment" Then
                                        
                                        'すでにコメントがあった場合はコメントのバックアップを保存
                                        cellComment = .Cells(i, j).Comment.Text
                                        
                                        'コメントをクリア
                                        .Cells(i, j).ClearComments
                                        
                                        'バックアップしたコメントとともに記載
                                        .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LOVnotFound(classification)" & vbCrLf & cellComment
                                        
                                    'コメントがなかった場合
                                    Else
                                        
                                        'そのままコメントを記載
                                        .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LOVnotFound(classification)"
                                        
                                    End If
                                    
                                    'コメントのサイズを自動調整
                                    .Cells(i, j).Comment.Shape.TextFrame.AutoSize = True
                                    
                                    'シートタブの色を黄色にするフラグを立てる
                                    sheetHasMismatch = True
                                    
                                End If
                            Next i
                        End If
                    End If
                Next j
            Else
                ' LOV_Entity_datamodelシートを参照
                Set lovSheet = ThisWorkbook.Worksheets("LOV_Entity_datamodel")

                ' このファイルの各ワークシートの7行目の各カラムをチェック
                For j = 1 To lastCol
                    ' Dictionaryを初期化
                    Set lovDict = CreateObject("Scripting.Dictionary")
                    
                    lovValue = .Cells(7, j).value
                    ' lovValueが空白でなく、かつ"LOV"を含み、"No LOV"を含まない場合にのみ処理を実行
                    If Not IsEmpty(lovValue) And InStr(lovValue, "LOV") > 0 And InStr(lovValue, "No LOV") = 0 Then
                        ' "LOV:"または"LOV :"の文字列を含むセルの値を取得し、"LOV:"または"LOV :"の後の文字列を抽出
                        If InStr(lovValue, "LOV:") > 0 Then
                            lovValue = Trim(Mid(lovValue, InStr(lovValue, "LOV:") + 4))
                        ElseIf InStr(lovValue, "LOV :") > 0 Then
                            lovValue = Trim(Mid(lovValue, InStr(lovValue, "LOV :") + 5))
                        End If
                        
                        ' LOVシートの2列目で抽出した文字列を検索
                        Set lovRange = lovSheet.Columns(2).Find(lovValue)
                        If Not lovRange Is Nothing Then
                            ' 検索の最初のアドレスを記憶
                            firstAddress = lovRange.Address
                            
                            ' LOVシートのD列とE列の値を最初から最後まで取得し、Dictionaryに格納
                            Do
                                lovDict(Trim(lovRange.Offset(0, 2).value)) = True ' D列をDictionaryに追加
                                lovDict(Trim(lovRange.Offset(0, 3).value)) = True ' E列をDictionaryに追加
                                ' 次の検索結果を取得
                                Set lovRange = lovSheet.Columns(2).FindNext(lovRange)
                                ' 次の検索結果がNothingでない、かつ最初のアドレスと異なる間ループを継続
                            Loop While Not lovRange Is Nothing And lovRange.Address <> firstAddress
                            
                            ' このファイルの各ワークシートの8行目以降をチェック
                            For i = 8 To lastRow
                                cellIFDataValue = .Cells(i, j).value
                                ' このファイルの各ワークシートの8行目以降のセルの値がDictionaryに存在しない場合、セルのフォントを太字にする
                                If Not IsEmpty(cellIFDataValue) And Not IsNull(cellIFDataValue) And Len(Trim(cellIFDataValue)) > 0 And Not lovDict.Exists(cellIFDataValue) Then
                                    
                                    'セルを黄色に変更
                                    .Cells(i, j).Interior.Color = vbYellow
                                    
                                    'カラムヘッダーを黄色に変更
                                    .Cells(1, j).Interior.Color = vbYellow
                                    .Cells(2, j).Interior.Color = vbYellow
                                    .Cells(3, j).Interior.Color = vbYellow
                                    
                                    'コメント保存用の変数を初期化
                                    cellComment = ""
                                    
                                    'すでにコメントが記入されていないかを確認
                                    If TypeName(.Cells(i, j).Comment) = "Comment" Then
                                        
                                        'すでにコメントがあった場合はコメントのバックアップを保存
                                        cellComment = .Cells(i, j).Comment.Text
                                        
                                        'コメントをクリア
                                        .Cells(i, j).ClearComments
                                        
                                        'バックアップしたコメントとともに記載
                                        .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LOVnotFound(datamodel)" & vbCrLf & cellComment
                                        
                                    'コメントがなかった場合
                                    Else
                                        
                                        'そのままコメントを記載
                                        .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LOVnotFound(datamodel)"
                                        
                                    End If
                                    
                                    'コメントのサイズを自動調整
                                    .Cells(i, j).Comment.Shape.TextFrame.AutoSize = True
                                    
                                    'シートタブの色を黄色にするフラグを立てる
                                    sheetHasMismatch = True
                                
                                End If
                            Next i
                        End If
                    End If
                Next j
            End If

            ' このファイルの各ワークシートの7行目以降の各セルをチェック
            For i = 7 To lastRow
                For j = 1 To lastCol
                
                     ' 各セルの値を取得
                    cellIFDataValue = .Cells(i, j).value
                    
                    ' 空白セルはスキップ
                    If cellIFDataValue = "" Or IsEmpty(cellIFDataValue) Then
                        GoTo NextCell
                    End If

                    ' 各セルのヘッダー：データ型を取得
                    Set headerCell = headerRow.Cells(2, j)
                    
                    '前後の空白を削除
                    headerCell = Trim(headerCell)
                    
                    '取得したヘッダー：データ型がStringなどの文字列であった場合には特に処理なし
                    If headerCell = "Integer" Or headerCell = "Double" Or headerCell = "Decimal" Or headerCell = "Date" Or headerCell = "Boolean" Or headerCell = "String" Or headerCell = "Numeric" Or headerCell = "" Then
                        
                    '取得したヘッダー：データ型がStringなどの文字列ではなかった場合
                    Else
                    
                        '1行目を取得　※Classification.txtなどを考慮
                        Set headerCell = headerRow.Cells(1, j)
                    
                    End If

                    ' ヘッダーの型を取得
                    headerType = CStr(headerCell.value)

                    ' データの型が一致しない場合
                    If Not CheckDataType(cellIFDataValue, headerType) Then
                        
                        'セルを黄色に変更
                        .Cells(i, j).Interior.Color = vbYellow
                        
                        'カラムヘッダーを黄色に変更
                        .Cells(1, j).Interior.Color = vbYellow
                        .Cells(2, j).Interior.Color = vbYellow
                        .Cells(3, j).Interior.Color = vbYellow
                        
                        'コメント保存用の変数を初期化
                        cellComment = ""
                        
                        'すでにコメントが記入されていないかを確認
                        If TypeName(.Cells(i, j).Comment) = "Comment" Then
                            
                            'すでにコメントがあった場合はコメントのバックアップを保存
                            cellComment = .Cells(i, j).Comment.Text
                            
                            'コメントをクリア
                            .Cells(i, j).ClearComments
                            
                            'バックアップしたコメントとともに記載
                            .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "DataTypeUnmatch" & vbCrLf & cellComment
                            
                        'コメントがなかった場合
                        Else
                            
                            'そのままコメントを記載
                            .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "DataTypeUnmatch"
                            
                        End If
                        
                        'コメントのサイズを自動調整
                        .Cells(i, j).Comment.Shape.TextFrame.AutoSize = True
                        
                        'シートタブの色を黄色にするフラグを立てる
                        sheetHasMismatch = True
                    
                    End If

                    ' バイト数制限をチェック
                    If maxLengths(j) > 0 And GetByteLength(CStr(cellIFDataValue)) > maxLengths(j) Then
                        
                        'セルを黄色に変更
                        .Cells(i, j).Interior.Color = vbYellow
                        
                        'カラムヘッダーを黄色に変更
                        .Cells(1, j).Interior.Color = vbYellow
                        .Cells(2, j).Interior.Color = vbYellow
                        .Cells(3, j).Interior.Color = vbYellow
                        
                        'コメント保存用の変数を初期化
                        cellComment = ""
                        
                        'すでにコメントが記入されていないかを確認
                        If TypeName(.Cells(i, j).Comment) = "Comment" Then
                            
                            'すでにコメントがあった場合はコメントのバックアップを保存
                            cellComment = .Cells(i, j).Comment.Text
                            
                            'コメントをクリア
                            .Cells(i, j).ClearComments
                            
                            'バックアップしたコメントとともに記載
                            .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LengthExceed" & vbCrLf & cellComment
                            
                        'コメントがなかった場合
                        Else
                            
                            'そのままコメントを記載
                            .Cells(i, j).AddComment Date & " " & Time & vbCrLf & "LengthExceed"
                            
                        End If
                        
                        'コメントのサイズを自動調整
                        .Cells(i, j).Comment.Shape.TextFrame.AutoSize = True
                        
                        'シートタブの色を黄色にするフラグを立てる
                        lengthExceeds = True
                        
                    End If
NextCell:
                Next j
            Next i

            ' エラーが含まれるシートタブの色を黄色にする
            If sheetHasMismatch Or lengthExceeds Then
                wsThisFile.Tab.Color = vbYellow
            End If
        End With
SkipSheet:
    Next wsThisFile

    ' 画面更新と自動計算を再開し、不要なイベントの無効化を解除
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    '処理開始のログ出力
    Debug.Print Now & " " & "CheckIFDataValue Finish"
    Debug.Print vbCrLf

    '一番先頭のシートへ移動
    Worksheets(1).Select
    
    ' 処理終了のポップアップを表示
    MsgBox "TenkaiDB IFファイルのデータチェック処理が完了しました。" & vbCrLf & "仕様と異なっているデータについてはセルが黄色で表示されており、メモにその理由が記載されています。ご確認下さい。", vbOKOnly, "完了"
End Sub

' 型の一致をチェックする関数
Function CheckDataType(ByVal value As Variant, ByVal dataType As String) As Boolean
    On Error GoTo ErrorHandler

    ' 値が空の場合は一致とみなす
    If IsEmpty(value) Then
        CheckDataType = True
    Else
        ' データ型に応じた判定を行う
        Select Case dataType
            Case "Integer"
                ' 整数型の判定
                CheckDataType = IsNumeric(value) And (CLng(value) = value) And (InStr(CStr(value), ".") = 0)
            Case "Double"
                CheckDataType = IsNumeric(value) And (CDbl(value) = value)
            Case "Decimal"
                ' 小数型の判定
                CheckDataType = IsNumeric(value) And (CDbl(value) = value)
            Case "Date"
                ' 日付型の判定
                If IsNumeric(value) = True Then
                    CheckDataType = False
                Else
                    Dim regEx As Object
                    Set regEx = CreateObject("VBScript.RegExp")
                    ' Format Check e.g.2024/05/12 17:30:51
                    regEx.Pattern = "^\d{4}/(0[1-9]|1[0-2])/(0[1-9]|[12]\d|3[01]) ([01]\d|2[0-3]):[0-5]\d:[0-5]\d$"
                    regEx.IgnoreCase = True
                    regEx.Global = False
                    If IsDate(value) And regEx.Test(value) Then
                        CheckDataType = True
                    Else
                        CheckDataType = False
                    End If
                End If
            Case "Boolean"
                ' ブール型の判定
                CheckDataType = (LCase(CStr(value)) = "true" Or LCase(CStr(value)) = "false")
            Case "String"
                ' 文字列型の判定
                ' もし値が文字列で "+" と "." を含んでいて、+と.を除去した場合に数字判定される場合、数式の一部と見なされるため、文字列としては不正とする。
                If InStr(CStr(value), "+") > 0 And InStr(CStr(value), ".") > 0 And IsNumeric(Replace(Replace(value, "+", ""), ".", "")) = True Then
                    CheckDataType = False
                ' もし値が日付形式で "/" を含んでいる場合、日付として認識されるため、文字列としては不正とする。
                ElseIf IsDate(CStr(value)) = True And InStr(CStr(value), "/") > 0 Then
                    CheckDataType = False
                ' もし値が "true" または "false" の文字列である場合、ブール型と見なされるため、文字列としては不正とする。
                ElseIf (LCase(CStr(value)) = "true" Or LCase(CStr(value)) = "false") = True Then
                    CheckDataType = False
                ' もし値が数値である場合、数値型と見なされるため、文字列としては不正とする。　※チェックのノイズが多くなるため、これをOFFにしてもよいかも…要確認★
                'ElseIf IsNumeric(value) = True Then
                    'CheckDataType = False
                ' 上記のどの条件にも該当しない場合、値は文字列として有効とする。
                Else
                    CheckDataType = True
                End If
            Case "Numeric"
                ' 数値型の判定
                CheckDataType = IsNumeric(value)
            Case Else
                ' その他の型の場合は一致とみなす
                CheckDataType = True
        End Select
    End If

    Exit Function

ErrorHandler:
    ' エラーが発生した場合は一致しないとみなす
    CheckDataType = False
End Function

' バイト数を計算する関数
Function GetByteLength(str As String) As Long
    Dim i As Long
    Dim byteLength As Long
    Dim charCode As Long

    byteLength = 0

    ' 文字列の各文字についてバイト数を計算
    For i = 1 To Len(str)
        charCode = AscW(Mid(str, i, 1))
        ' 文字のUnicode値を取得
        If charCode <= 127 Then
            ' 1バイト文字（ASCII文字）の場合
            byteLength = byteLength + 1
        Else
            ' 3バイト文字（非ASCII文字）の場合
            byteLength = byteLength + 3
        End If
    Next i

    GetByteLength = byteLength
End Function
