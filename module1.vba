Option Explicit

' フォルダ選択ダイアログを表示し、選択されたIFファイルが格納されたフォルダのパスを返す関数
Function SelectFolder() As String
    Dim fd As FileDialog
    MsgBox "TenkaiDB IFファイルが格納されているフォルダを選択してください。" & vbCrLf & "指定されたフォルダ内のファイルのデータをこのファイルに取り込みます。", vbOKOnly
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "TenkaiDB IFファイルが格納されているフォルダを選択してください。"
    If fd.Show = -1 Then
        SelectFolder = fd.SelectedItems(1)
    Else
        SelectFolder = ""
    End If
    Debug.Print Now & " " & "Selected Folder: " & SelectFolder
End Function

' メイン処理
Sub ImportIFData()
    Dim selectedFolder As String         ' 選択されたIFファイルが格納されたフォルダのパス
    Dim folderPath As String             ' 選択されたフォルダのパス（末尾にを追加）j
    Dim fileName As String               ' フォルダ内の各IFファイルの名前
    Dim filePath As String               ' フォルダ内の各IFファイルのフルパス
    Dim wsLoopSheet As Worksheet         ' 各シートをループするためのWorksheetオブジェクト
    Dim wsFileNameError As Worksheet     ' ファイル名間違いシートのWorksheetオブジェクト
    Dim newSheet As Worksheet            ' 新しく作成するシートのWorksheetオブジェクト
    Dim Answer As Long                   ' メッセージボックスの戻り値
    Dim outputSheets As Object           ' 出力用のシートを管理するコレクションオブジェクト
    Dim outputSheet As Worksheet         ' 出力用のシートのWorksheetオブジェクト
    Dim fso As Object                    ' FileSystemObjectオブジェクト
    Dim ts As Object                     ' TextStreamオブジェクト
    Dim textFile As String              'テキストファイルのすべてのデータを格納する変数
    Dim textFileline As Variant           ' テキストファイルの各行を格納する変数
    Dim textFilelineArray As Variant     ' 各行をタブで分割した配列
    Dim rowCounter As Long               ' 行カウンタ
    Dim colCounter As Long               ' 列カウンタ
    Dim targetSheetName As String        ' シート名として使用するファイル名（最大30文字）
    Dim prefix As String                 ' シート名の先頭に付ける文字列

    '処理開始のログ出力
    Debug.Print vbCrLf
    Debug.Print Now & " " & "ImportIFData Start"
    
    ' 画面更新と自動計算を停止
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' フォルダ選択ダイアログを表示してIFファイルが格納されたフォルダを選択
    selectedFolder = SelectFolder()
    If selectedFolder = "" Then GoTo Cleanup
    
    ' データチェック開始する前に、既存のシートを削除してよいかのメッセージを表示
    Answer = MsgBox("いま存在するSheetついては削除されますが、よろしいですか？", vbYesNo)
        
    ' Noボタンを推された場合は、Sub処理を終了させる
    If Answer = vbNo Then
        Exit Sub
    End If
        
    ' TenkaiDB IFファイルのデータのシート出力開始のメッセージを表示
    MsgBox "TenkaiDB IFファイルのデータを個別のシートに出力します。", vbOKOnly

    ' 前回実行した際に残っていたシートを削除
    For Each wsLoopSheet In ThisWorkbook.Sheets
        If wsLoopSheet.Name <> "Corresponding Sheets" And wsLoopSheet.Name <> "ファイル名間違い" Then
            Application.DisplayAlerts = False
            wsLoopSheet.Delete
            Application.DisplayAlerts = True
        End If
    Next wsLoopSheet

    ' ファイル名間違いシートを取得または作成。ファイル名間違いがあった場合も処理を継続させるためのエラーハンドリング
    On Error Resume Next
    Set wsFileNameError = ThisWorkbook.Sheets("ファイル名間違い")
    On Error GoTo 0
    If wsFileNameError Is Nothing Then
        Set wsFileNameError = ThisWorkbook.Sheets.Add
        wsFileNameError.Name = "ファイル名間違い"
    End If

    ' 出力用のシートを管理するコレクションを作成
    Set outputSheets = CreateObject("Scripting.Dictionary")

    ' 選択したフォルダ内のIFファイルを取得
    folderPath = selectedFolder & """"
    fileName = Dir(folderPath & "*.*")

    ' FileSystemObjectを作成
    Set fso = CreateObject("Scripting.FileSystemObject")

    Do While fileName <> ""
        ' 選択したフォルダ内の現在のIFファイルのフルパスを取得
        filePath = folderPath & fileName
        Debug.Print Now & " " & "Processing File: " & filePath

        ' シート名の先頭に付ける文字列を決定
        If InStr(fileName, "ptc_") > 0 Then
            prefix = "(PtCl)"
        ElseIf InStr(fileName, "dcc_") > 0 Then
            prefix = "(DcCl)"
        Else
            prefix = "(dm)"
        End If

        ' シート名として使用するファイル名を取得（最大30文字）
        targetSheetName = prefix & Left(fileName, 30 - Len(prefix))
        Debug.Print "Target Sheet Name: " & targetSheetName

        ' 出力用のシートが既に存在するか確認
        If Not outputSheets.Exists(targetSheetName) Then
            ' 出力用のシート存在しない場合、新しいシートを作成
            Set newSheet = ThisWorkbook.Sheets.Add
            newSheet.Name = targetSheetName
            outputSheets.Add targetSheetName, newSheet
            Debug.Print Now & " " & "New output sheet created: " & targetSheetName
        End If

        ' 出力先のシートを取得
        Set outputSheet = outputSheets(targetSheetName)

        '出力先のシート全体をテキスト形式に設定
        outputSheet.Cells.NumberFormat = "@"

       'テキスト形式のIFファイルをUTF-8形式で開く
        With CreateObject("ADODB.Stream")
            .Charset = "UTF-8"
            .Open
            .LoadFromFile filePath
            textFile = .ReadText
            .Close
            
            '改行コードをvbCrLfに統一置換
            textFile = Replace(textFile, vbCrLf, vbCr)
            textFile = Replace(textFile, vbLf, vbCr)
            textFile = Replace(textFile, vbCr, vbCrLf)
                        
            'データを改行ごとに分割
            textFileline = Split(textFile, vbCrLf)
            
            '1行ごとにループ処理
            For rowCounter = 0 To UBound(textFileline)
                
                '1行をタブごとに分割
                textFilelineArray = Split(textFileline(rowCounter), vbTab)
                               
                '1個目のデータから、最終のデータまでループ
                For colCounter = LBound(textFilelineArray) To UBound(textFilelineArray)
                    outputSheet.Cells(rowCounter + 1, colCounter + 1).value = textFilelineArray(colCounter)
                Next colCounter
            
            Next rowCounter
            
        End With



        ' 日付が正しく表示されるように列幅を自動調整
        outputSheet.Columns.AutoFit

        ' 次のIFファイルを取得
        fileName = Dir()
        
    Loop

    '処理完了のログ出力
    Debug.Print Now & " " & "ImportIFData Finish"
    Debug.Print vbCrLf

    ' シート出力終了のメッセージを表示
    MsgBox "TenkaiDB IFファイルのデータの取り込みが終了しました。", vbOKOnly

    '一番先頭のシートへ移動
    Worksheets(1).Select

Cleanup:
    ' 画面更新と自動計算を再開
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
