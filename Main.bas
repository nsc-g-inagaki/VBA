Attribute VB_Name = "Main"
Option Explicit
'データが始まる行
Public Const FirstDataRow = 3

'客先タイムシート - 社員番号
Public Const CtsEmployeeNumColumn = 1

'客先タイムシート - 氏名
Public Const CtsEmployeeNameColumn = 2

'客先タイムシート - 時間
Public Const CtsWorkingHoursColumn = 3

'Socia - 社員番号
Public Const SocEmployeeNumColumn = 4

'Socia - 氏名
Public Const SocEmployeeNameColumn = 5

'Socia - 時間
Public Const SocWorkingHoursColumn = 6

'変換後の列番号
Public Const ConvertedColumn = 7

'チェックの列番号
Public Const CheckColumn = 8


'タイムシートチェックボタンのクリック処理
Sub BtnTimeSheetClick()
    
    '画面更新を無効にする
    Application.ScreenUpdating = False
  
    'ClearFormatsを呼び出して、シートのフォーマットを初期状態に戻す
    Call ClearFormats
    
    'データのチェックを行う
    If Not Validation.ValidateData Then
        '不正な値がある場合メッセージボックスを表示する
        MsgBox _
            "入力内容に誤りがります。" & vbNewLine & "ハイライトされたセル・行の入力内容をご確認ください。", _
            vbOKOnly + vbCritical, _
            "バリデーションエラー"
            
        'ストップオペレーションへ行く
        GoTo StopOp
    End If
    
    '時間のチェック処理を呼び出す
    Call CheckWorkingHours
    
'ストップオペレーションラベル
'処理を止めるときにここより下を実行
StopOp:

    '画面更新を有効にする
    Application.ScreenUpdating = True
    
    'シートの上が表示されるようにA1のセルを選択する
    Cells(1, 1).Select
    
    End Sub
    

'データ削除ボタンのクリック
Sub BtnDeleteClick()
    
    'G3からHの最終行までのデータを消す
    Range(Cells(FirstDataRow, ConvertedColumn), _
        Cells(GetLastRowInColumn(CheckColumn), CheckColumn)).ClearContents
        
    'グラフ関係のクリア処理を呼ぶ
    Call ClearGraph
End Sub


'エラースタイルを適用する
'
'引数 Cell スタイルを適用するセルもしくは範囲(Range)
'
Sub ApplyErrorStyle(Cell As Range)

    With Cell
        '背景色を赤っぽい色に変更
        .Interior.color = RGB(255, 71, 71)
        
        '文字の色を白に変更
        .Font.color = RGB(255, 255, 255)
        
        '文字を太文字に変更
        .Font.Bold = True
    End With

End Sub


'フォーマットの初期化を行う
Sub ClearFormats()

    'A3からH列シートの最終行までのフォーマットを消す
    Range("A3", Cells(Rows.Count, CheckColumn)).ClearFormats
       
    '客先側時間列(C列)と変換後列(G列)の書式設定を行う
    Application.Union( _
        Columns(CtsWorkingHoursColumn), _
        Columns(ConvertedColumn)).NumberFormat = "0.00"
        
    'Socia側時間列(F列の書式設定を行う
    Columns(SocWorkingHoursColumn).NumberFormat = "[h]:mm"
    
    'チェック列(H列)を中央揃えに設定
    Columns(CheckColumn).HorizontalAlignment = xlCenter
    
    '客先側時間列(C列)のデータが始まる行(3)から同列とシートの最終行までと
    'Socia側時間列(F列)のデータが始まる行(3)から同列とシートの最終行まで
    'までの範囲で以下処理を実行
    With Application.Union( _
        Range( _
            Cells(FirstDataRow, CtsWorkingHoursColumn), _
            Cells(GetLastRowInSheet, CtsWorkingHoursColumn)), _
        Range( _
            Cells(FirstDataRow, SocWorkingHoursColumn), _
            Cells(GetLastRowInSheet, SocWorkingHoursColumn)))
        
        '文字の色を白色に設定
        .Font.color = RGB(255, 255, 255)
        
        '文字を太文字に設定
        .Font.Bold = True
        
        '背景色を青っぽい色に変更
        .Interior.color = RGB(155, 194, 230)
        
    End With
    
End Sub


'労働時間の一致をチェックする処理
Private Sub CheckWorkingHours()

    'セルに表示する値を保持
    Dim Msg As String
    
    '一致数を保持
    Dim CorrectCount As Long: CorrectCount = 0
    
    '不一致数を保持
    Dim WrongCount As Long: WrongCount = 0
    
    'Socia側時間列とデータが始まる行のセルを選択(F3)
    Cells(FirstDataRow, SocWorkingHoursColumn).Select
    
    '選択されているセルが空でない間以下の処理を実行
    Do While ActiveCell <> ""
        
        'セルに表示する値を空にする
        Msg = ""
    
        '選択されているセルと同じ行にある変換後列(G列)に変換した労働時間を設定
        Cells(ActiveCell.Row, ConvertedColumn).Value = ConvertHoursToDecimal(ActiveCell.Value)
        
        '変換後の労働時間が客先側と一致していることをチェック
        If Validation.ValidateWorkingHours(ActiveCell.Row) Then
            
            'セルに表示する値を設定
            Msg = "〇"
            
            '一致数を増やす
            CorrectCount = CorrectCount + 1
        Else
            
            'セルに表示する値を設定
            Msg = "×"
            
            '一致数を増やす
            WrongCount = WrongCount + 1
        End If
        
        'チェック列に値を反映
        Cells(ActiveCell.Row, CheckColumn).Value = Msg
        
        '選択されているセルの一つ下のセルを選択
        ActiveCell.Offset(1, 0).Select
    Loop
    
    'グラフ用のデータを反映する処理を呼び出す
    Call PopulateGraphData(CorrectCount, WrongCount)
    
End Sub


'グラフにデータを入れる処理を行う
'
'引数 CorrectCount 時間に不備がなかった数
'引数 WrongCount 時間に不備があった数
'
Private Sub PopulateGraphData(CorrectCount, WrongCount)
    
    'それぞれのセルに値を入れる
    Cells(3, "N").Value = CorrectCount
    Cells(4, "N").Value = WrongCount
    
    'TODO　グラフの作成処理
End Sub


'グラフのクリア処理
Private Sub ClearGraph()

    'それぞれのセルの値を0に設定
    Cells(3, "N").Value = 0
    Cells(4, "N").Value = 0
    
    'TODO　グラフのデータをクリア
End Sub


'列で使用されている最後の行を取得する処理
'
'引数 ColumnIndex 列の番号
'
'戻り値 指定された列で使用されている最終行もしくはデータが始まる行
'
Function GetLastRowInColumn(ColumnIndex) As Long
    
    '最終行を保持
    Dim LastRow As Long
    
    '列の最終行を取得
    LastRow = Cells(Rows.Count, ColumnIndex).End(xlUp).Row
    
    '最終行がデータの開始行よりも小さい場合はデータが始まる行を返す
    'それ以外は取得した値を返す
    GetLastRowInColumn = IIf(LastRow < FirstDataRow, FirstDataRow, LastRow)
End Function


'シートで使用されている最後の行を取得する処理
'
'戻り値 シートで使用されている最終行
'
Function GetLastRowInSheet() As Long
    'シートで使用されている最終行を取得して返す
    GetLastRowInSheet = ActiveSheet.UsedRange.Rows.Count
End Function



'データの行の範囲を取得する 客先タイムシートとSociaの部分のみ
'
'引数 Row 範囲を取得したい行
'
'戻り値 指定された行の客先側社員番号列からSocia側時間列までの範囲
'
Function GetDataRowRange(ByVal Row As Long) As Range
    '範囲を指定された行の客先側社員番号からSocia側時間までに指定して返す
    Set GetDataRowRange = Range(Cells(Row, CtsEmployeeNumColumn), Cells(Row, SocWorkingHoursColumn))

End Function


'時間を小数形式に変あkンする処理
'
'引数 Data 変換する時間
'
'戻り値 引数で渡された時間を小数化したもの
'
Private Function ConvertHoursToDecimal(Data As Variant) As Double
    
    '時間を小数に変換して返す
    ConvertHoursToDecimal = Int(Data) * 24 + Hour(Data) + Round(Minute(Data) / 60, 2)
    
End Function



