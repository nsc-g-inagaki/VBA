Attribute VB_Name = "Import"
Option Explicit

'客先タイムシートのファイル名のテンプレート
'客先タイムシート.csvで終わるものすべて当てはまる
Private Const CtsFileNameFormat = "*客先タイムシート.csv"


'Sociaタイムシートのファイル名のテンプレート
'Socia.csvで終わるものすべて当てはまる
Private Const SocFileNameFormat = "*Socia.csv"


'インポートの処理
Sub ImportFiles()

    'フォルダのパスを保持
    Dim FolderPath As String
    
    '最新の客先タイムシートファイル名を保持
    Dim LatestCtsFile As String
    
    '最新のSociaタイムシートファイル名を保持
    Dim LatestSocFile As String

    '注意メッセージを表示
    Call MsgBox("インポートを行う前に、以下のフォーマットでCSVファイルがフォルダ内にあることをご確認ください。" _
        & vbNewLine & vbNewLine _
        & "「客先タイムシート」と書かれたCSVファイルが存在すること" _
        & vbNewLine & "「Socia」と書かれたCSVファイルが存在すること" _
        & vbNewLine & vbNewLine & "＊ファイルが複数ある場合は最新のものを使用します。", _
        vbOKOnly + vbInformation, _
        "ファイルインポート確認事項" _
    )
    
    'インポートファイルが入っているフォルダのパスを取得する
    FolderPath = AskFolderPath()

    'キャンセルなどの対応　処理を止める
    If FolderPath = "" Then GoTo Finish
    
    'フォルダ内最新の客先タイムシートのファイル名を取得
    LatestCtsFile = GetLatestFile(FolderPath, CtsFileNameFormat)
    
    'フォルダ内最新のSociaタイムシートのファイル名を取得
    LatestSocFile = GetLatestFile(FolderPath, SocFileNameFormat)
    
    'ファイルがあること、同じファイルを選択していないことチェック
    If LatestCtsFile = "" Or LatestSocFile = "" Or LatestCtsFile = LatestSocFile Then
        Call MsgBox("条件を満たすファイルがフォルダ内に見つかりませんでした。" _
            & vbNewLine & "フォルダの内容及びパスを確認し、もう一度やり直してください。", _
            vbOKOnly + vbCritical, _
            "ファイルが見つかりませんでした。" _
        )
        GoTo Finish
    End If
    
    '客先タイムシートデータをインポート
    Call ImportCtsValues(FolderPath, LatestCtsFile)
    
    'Sociaタイムシートデータをインポート
    Call ImportSocValues(FolderPath, LatestSocFile)
        
Finish:

End Sub


'フォルダのパスを取得する処理
'
'戻り値 ユーザーが選択したフォルダへのパス
'
Private Function AskFolderPath() As String
    
    'フォルダーパスを保持
    Dim FolderPath As String
    
    'フォルダーピッカーを生成
    With Application.FileDialog(msoFileDialogFolderPicker)
        '複数選択を無効
        .AllowMultiSelect = False
        '開いたとき最初に表示される場所を指定　ユーザーのフォルダ(環境変数より)
        .InitialFileName = Environ("USERPROFILE") & Application.PathSeparator
        
        '選択=>OKが選択されたときの処理
        If .Show = -1 Then FolderPath = .SelectedItems(1)
    End With

    '取得したパスを返す
    AskFolderPath = FolderPath
End Function


'最新のファイルを取得する
'
'引数 Folder：フォルダへのパス FileName：検索をするファイル名(今回はファイル名テンプレートを使用)
'
'戻り値 最新ファイル名
'
Private Function GetLatestFile(ByVal Folder As String, FileName As String) As String

    '最新日時を保持
    Dim LatestDate As Date
    
    '最新ファイル
    Dim LatestFile As String
    
    'ファイル名
    Dim File As String
    
    'パスの最後にセパレーターを追加(Windowsは￥Mac,Linuxでは/)
    Folder = Folder & Application.PathSeparator
    
    'Dirで指定した、パスのものをとってくる
    File = Dir(Folder & FileName)
    
    'フォルダ内のテンプレートに該当するファイルを一つ一つループ
    Do While File <> ""
        
        'ファイルの日時が最新より新しいかをチェック
        If FileDateTime(Folder & File) > LatestDate Then
            
            '最新ファイルにファイルを設定
            LatestFile = File
            
            '最新日時を更新
            LatestDate = FileDateTime(Folder & File)
                
        End If
        
        '次のファイルを取得
        File = Dir
    Loop
    
    '最新のファイル名を返す
    GetLatestFile = LatestFile
End Function


'客先タイムシートの値をインポートする
'
'引数 Path：ファイルまでのパス File：ファイル名
'
Private Sub ImportCtsValues(Path As String, File As String)
        
    '行を保持
    Dim Row As Long
    
    'CSVの行を配列にしたものを保持
    Dim ArrLine As Variant
    
    'ファイルを読み込むのに必要なオブジェクト
    Dim Fso As New Scripting.FileSystemObject
    
    '行にデータが始まる行を設定
    Row = main.FirstDataRow
    
    'ファイルを開く ForReading=>読み込み専用
    With Fso.OpenTextFile(Path & Application.PathSeparator & File, ForReading)
            
        '最終行ではない場合一行飛ばす
        'ファイルお開いた直後なので、1行目を飛ばしている
        If Not .AtEndOfLine Then .SkipLine
        
        '最終行にならない間処理を続行
        Do Until .AtEndOfLine
            '1行読み込んで「,」で区切る
            ArrLine = Split(.ReadLine, ",")
            
            '必要なデータを読み込む
            Cells(Row, main.CtsEmployeeNumColumn).Value = ArrLine(1)
            Cells(Row, main.CtsEmployeeNameColumn).Value = ArrLine(2)
            '時間の計算をする
            Cells(Row, main.CtsWorkingHoursColumn).Value = CalculateCtsHours(ArrLine)
            
            '行を一つずらす
            Row = Row + 1
        Loop
        
        'ファイルを閉じる
        .Close
    End With
End Sub


'Sociaタイムシートの値をインポートする
'
'引数 Path：ファイルまでのパス File：ファイル名
'
Private Sub ImportSocValues(Path As String, File As String)

    '行を保持
    Dim Row As Long
    
    'CSVの行を配列にしたものを保持
    Dim ArrLine As Variant
    
    'ファイルを読み込むのに必要なオブジェクト
    Dim Result As Variant
    
    '客先社員番号の最後の行を保持
    Dim LastCtsNumRow As Long
    
    'ファイルを読み込むのに必要なオブジェクト
    Dim Fso As New Scripting.FileSystemObject
    
    '行にデータが始まる行を設定
    Row = main.FirstDataRow
    
    '客先社員番号の最後の曲を取得
    LastCtsNumRow = main.GetLastRowInColumn(main.CtsEmployeeNumColumn)
    
    'ファイルを開く ForReading=>読み込み専用
    With Fso.OpenTextFile(Path & Application.PathSeparator & File, ForReading)
     
        '最終行ではない場合一行飛ばす
        'ファイルお開いた直後なので、1行目を飛ばしている
        If Not .AtEndOfStream Then .SkipLine
            
        '最終行にならない間処理を続行
        Do Until .AtEndOfStream
            
            '1行読み込んで「,」で区切る
            ArrLine = Split(.ReadLine, ",")
                
            '社員番号が客先でインポートしたものと一致しているかを探す
            Result = Application.Match( _
                CLng(ArrLine(0)), _
                Range( _
                    Cells(main.FirstDataRow, main.CtsEmployeeNumColumn), _
                    Cells(LastCtsNumRow, main.CtsEmployeeNumColumn)), _
                0)
            
            '検索の結果エラーでなければデータを読み込む
            If Not IsError(Result) Then
                Cells(Result + 2, main.SocEmployeeNumColumn).Value = ArrLine(0)
                Cells(Result + 2, main.SocEmployeeNameColumn).Value = ArrLine(1)
                '時間を計算して読み込む
                Cells(Result + 2, main.SocWorkingHoursColumn).Value = CalculateSocHours(ArrLine)
            End If

        Loop
    
    End With
    
End Sub


'客先タイムシートの時間を計算
'
'引数 Data：CSV1行分のデータが入った配列
'
'戻り値 総労働時間
'
Private Function CalculateCtsHours(Data As Variant) As Double

    '労働時間を保持
    Dim TotHours As Double
    
    'カウンター
    Dim I As Integer
    
    For I = 5 To 9
    
        '6の場合スキップ
        If I = 6 Then GoTo Continue
        
        '文字列から数字に変換して値を足す
        TotHours = TotHours + CDbl(Data(I))
Continue:
    Next
    
    '合計を戻す
    CalculateCtsHours = TotHours
    
End Function


'Socタイムシートの時間を計算
'
'引数 Data：CSV1行分のデータが入った配列
'
'戻り値 総労働時間
'
Private Function CalculateSocHours(Data As Variant) As Double
    
    '結果を保持
    Dim Result As Double
    
    'カウンター
    Dim I As Integer
    
    '6個目のCSVデータ時間を変換
    Result = ConvertTimeStrToDbl(Data(6))
    
    '残りのデータも変換して計算する
    For I = 7 To 9
        Result = Result - ConvertTimeStrToDbl(Data(I))
        
    Next
    
    '合計を戻す
    CalculateSocHours = Result
End Function

'時間を小数に変換(Excelのフォーマット)
'
'引数 Data：CSV1行分のデータが入った配列
'
'戻り値 小数化された時間
'
Private Function ConvertTimeStrToDbl(ByVal Value As String) As Double
    
    '結果を保持
    Dim Result As Double
    
    'ループで使用するオブジェクト
    Dim Data As Variant
    
    '計算に使う値を保持
    Dim tmp As Double
    
    'Excelは24時間を1として保持するので、まず24を入れる
    tmp = 24
    
    '結果の初期値
    Result = 0
    
    '時間の値を「:」で区切って一つずつDataに入れて順に処理
    '時間　計算は　時間/24
    '分　　計算は　分/24*60
    '秒　　計算は　秒/24*60*60
    For Each Data In Split(Value, ":")
        
        '時間を計算して合計に足す
        Result = Result + Data / tmp
        
        '計算に使う値を更新
        tmp = tmp * 60
    Next
        
    '結果を返す
    ConvertTimeStrToDbl = Result
End Function
