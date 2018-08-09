Sub LogReadWrite_1(ByVal DirectoryPath As String)

    ''ファイルの読み出し・書き込みを行う関数
    
    ''変数定義
    Dim OpenFileName As String
    Dim buf As String
    Dim n As Long
    
    ''画面の更新をオフにすることで処理速度を上げる
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ''エクセルシートの行を指定
    n = 5
    
    OpenFileName = DirectoryPath & "/log1.txt"
    Open OpenFileName For Input As #1
    
        Do Until EOF(1)
            Line Input #1, buf
            n = n + 1
            Cells(n, 1) = buf
        Loop
    Close #1
    
    ''データ文字列をカンマで区切る関数をコール
    divideByComma
    
End Sub


Sub divideByComma()

    ''データ文字列をカンマで区切る関数
    
    ''変数定義
    Dim ws As Worksheet
    
    Set ws = Worksheets("1")
    
    ''カンマでデータを区切る
    ws.Columns("A").TextToColumns Comma:=True

End Sub
