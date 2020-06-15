Attribute VB_Name = "Validation"
Option Explicit
'正規表現で値のチェックを行うので、パターンを固定にしておく
'社員番号の正規表現　半角数字0-9が最低５最高６桁

Private Const EmployeeNumberPattern As String = "^[0-9]{5,6}$"

'名前の正規表現 _
\u3040-\u309Fー ひらがなを表している　あ-んー _
\u30A0-\u30FF カタカナを表している　ァ-ヾ _
\u4E00-\u9FFF 漢字を表している　一-龠々 _
A-Za-Z 半角アルファベット大文字小文字を表している _
+ 最低１文字を表す
Private Const NamePattern As String = "^[\u3040-\u309Fー\u30A0-\u30FF\u4E00-\u9FFFA-Za-z]+"

'客先タイムシートの時間正規表現
'0-9 1から3桁まで(必須)　小数点と0-9　1から2桁まで(任意) _
この表現でバリデーションできている範囲
Private Const CtsWorkingHoursPattern As String = "^\d{1,3}(\.\d{1,2})?$"

'Sociaの時間正規表現
'0-9 1から3桁まで:0-5と0-9(必須) :0-5 0-9(任意)
Private Const SocWorkingHoursPattern As String = "^[0-9]{1,3}:[0-5][0-9](:[0-5][0-9])?$"


'データのフォーマットをチェックする処理
'
'戻り値 True：データ不備なし False：データに不備あり
'
Function ValidateData() As Boolean

    '行を保持する変数
    Dim Row As Long
    'エラーの数を保持する
    Dim ErrorCount As Long: ErrorCount = 0
    
    '有効な社員番号かどうかを保持する
    Dim IsValidEmployeeNum As Boolean
    
    '有効な名前かどうかを保持する
    Dim IsValidName As Boolean
    
    '有効な労働時間かどうかを保持する
    Dim IsValidWorkingHours As Boolean
    
    '有効なデータかどうかを保持する
    Dim IsValidData As Boolean
    
    'Mainモジュールよりデータが開始する行と _
     シート上で最後の行を取得しループする
     
    For Row = main.FirstDataRow To main.GetLastRowInSheet
    
        '社員番号のフォーマットをチェック
        IsValidEmployeeNum = ValidateEmployeeNumber(Row)
        
        '社員名のフォーマットをチェック
        IsValidName = ValidateEmployeeName(Row)
        
        '労働時間のフォーマットをチェック
        IsValidWorkingHours = ValidateWorkingHoursFormat(Row)
        
        '社員番号、氏名、時間のフォーマットが正しいことをチェック
        If IsValidEmployeeNum And _
           IsValidName And _
           IsValidWorkingHours Then
           '社員番号、氏名、時間のフォーマットが正しいとき
           '客先側とSocia側の一致を判定
           IsValidData = CompareDatas(Row)
        End If
        
        'データがすべて正しいことをチェック
        If Not IsValidEmployeeNum Or _
           Not IsValidName Or _
           Not IsValidWorkingHours Or _
           Not IsValidData Then
            '入力されているデータが一つでも正しくない場合
            'エラーカウンターを一つ増やす
            ErrorCount = ErrorCount + 1
        End If
    Next
    
    'エラーの数が0かどうかを返す
    ValidateData = ErrorCount = 0
End Function


'社員番号のフォーマットをチェックする処理
'
'引数 Row 処理を実行する行
'
'戻り値 True：一致 False：不一致
'
Private Function ValidateEmployeeNumber(Row As Long) As Boolean
    
    'エラーの数を保持する
    Dim ErrorCount As Integer: ErrorCount = 0
    
    '客先側の社員番号のセルを選択 Rowは今いる行
    Cells(Row, main.CtsEmployeeNumColumn).Select
    
    'ValidatePatternを呼び出して、セルの値のフォーマットをチェック
    If Not ValidatePattern(EmployeeNumberPattern, ActiveCell.Value) Then
        
        'エラーカウンターの増やす
        ErrorCount = ErrorCount + 1
        
        '選択されているセルに、エラーのスタイルを適用する
        Call main.ApplyErrorStyle(ActiveCell)
        
    End If
    
    'Socia側の社員番号のセルを選択 Rowは今いる行
    Cells(Row, main.SocEmployeeNumColumn).Select
    
    'ValidatePatternを呼び出して、セルの値のフォーマットをチェック
    If Not ValidatePattern(EmployeeNumberPattern, ActiveCell.Value) Then
        
        'エラーカウンターの増やす
        ErrorCount = ErrorCount + 1
        '選択されているセルに、エラーのスタイルを適用する
        Call main.ApplyErrorStyle(ActiveCell)
        
    End If
    
    'エラーの数が0かどうかを返す
    ValidateEmployeeNumber = ErrorCount = 0
End Function


'氏名のフォーマットをチェックする処理
'
'引数 Row 処理を実行する行
'
'戻り値 True：一致 False：不一致
'
Private Function ValidateEmployeeName(Row As Long) As Boolean

    'エラーの数を保持する
    Dim ErrorCount As Integer: ErrorCount = 0
    
    '客先側の氏名のセルを選択 Rowは今いる行
    Cells(Row, main.CtsEmployeeNameColumn).Select
    
    'ValidatePatternを呼び出して、セルの値のフォーマットをチェック
    If Not ValidatePattern(NamePattern, ActiveCell.Value) Then
        
        'エラーカウンターの増やす
        ErrorCount = ErrorCount + 1
        
        '選択されているセルに、エラーのスタイルを適用する
        Call main.ApplyErrorStyle(ActiveCell)
    End If
    
    'Socia側の氏名のセルを選択 Rowは今いる行
    Cells(Row, main.SocEmployeeNameColumn).Select
    
    'ValidatePatternを呼び出して、セルの値のフォーマットをチェック
    If Not ValidatePattern(NamePattern, ActiveCell.Value) Then
        
        'エラーカウンターの増やす
        ErrorCount = ErrorCount + 1
        
        '選択されているセルに、エラーのスタイルを適用する
        Call main.ApplyErrorStyle(ActiveCell)
    End If
    
    'エラーの数が0かどうかを返す
    ValidateEmployeeName = ErrorCount = 0

End Function


'労働時間のフォーマットをチェックする処理
'
'引数 Row 処理を実行する行
'
'戻り値 True：一致 False：不一致
'
Private Function ValidateWorkingHoursFormat(Row As Long) As Boolean
    
    'エラーの数を保持する
    Dim ErrorCount As Integer: ErrorCount = 0
    
    '客先側の時間のセルを選択 Rowは今いる行
    Cells(Row, main.CtsWorkingHoursColumn).Select
    
    'ValidatePatternを呼び出して、セルの値のフォーマットをチェック
    If Not ValidatePattern(CtsWorkingHoursPattern, ActiveCell.Value) Then
        
        'エラーカウンターの増やす
        ErrorCount = ErrorCount + 1
        
        '選択されているセルに、エラーのスタイルを適用する
        Call main.ApplyErrorStyle(ActiveCell)
        
    End If
    
    'Socia側の時間のセルを選択 Rowは今いる行
    Cells(Row, main.SocWorkingHoursColumn).Select
    
    'ValidatePatternを呼び出して、セルの値のフォーマットをチェック
    If Not ValidatePattern(SocWorkingHoursPattern, ActiveCell.Text) Then
        
        'エラーカウンターの増やす
        ErrorCount = ErrorCount + 1
        
        '選択されているセルに、エラーのスタイルを適用する
        Call main.ApplyErrorStyle(ActiveCell)
        
    End If
    
    'エラーの数が0かどうかを返す
    ValidateWorkingHoursFormat = ErrorCount = 0
    
End Function


'客先タイムシートとSociaのデータの一致をチェック(社員番号と氏名のみ)
'
'引数 Row 処理を実行する行
'
'戻り値 True：一致 False：不一致
'
Private Function CompareDatas(Row As Long) As Boolean
    
    'データの一致状態を保持
    Dim IsDataMatch As Boolean: IsDataMatch = True
    
    '社員番号と氏名の一致をチェック
    If (Cells(Row, main.CtsEmployeeNumColumn) <> Cells(Row, main.SocEmployeeNumColumn)) Or _
        (Cells(Row, main.CtsEmployeeNameColumn) <> Cells(Row, main.SocEmployeeNameColumn)) Then
        'どちらか一つでも一致しない場合
        
        '不一致を保持
        IsDataMatch = False
        
        '行にエラーのスタイルを適用
        Call main.ApplyErrorStyle(main.GetDataRowRange(Row))
        
    End If
    
    'データの一致状態を返す　True：一致　False：不一致
    CompareDatas = IsDataMatch
    
End Function


'指定された正規表現を使ってデータの妥当性をチェック
'
'引数　ValidatePattern：正規表現のパターン(String型)
'引数　Data：チェックをするデータ(String型)
'
'戻り値　True：一致　False：不一致
'
Function ValidatePattern(ByVal ValidationPattern As String, ByVal Data As String) As Boolean
    
    '正規表現が使用できるように設定
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    '正規表現オブジェクトの設定
    With regex
        .Global = True '文字列全体で検索するように設定
        .IgnoreCase = False  '大文字と小文字を区別するように設定
        .Pattern = ValidationPattern '正規表現のパターンを設定
    End With
    
    '正規表現をマッチング
    ValidatePattern = regex.Test(Data)
    
End Function

'変換後の労働時間が一致していることをチェックする処理
'
'引数 Row: 比較処理を実行する行を表す
'
'戻り値　True：一致　False:不一致
Function ValidateWorkingHours(Row As Long) As Boolean

    '客先タイムシートの時間と変更後列の時間が一致しているかどうかを返す
    ValidateWorkingHours = Cells(Row, main.CtsWorkingHoursColumn) = Cells(Row, main.ConvertedColumn)
    
End Function
