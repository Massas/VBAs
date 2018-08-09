Sub ExecuteMacros()

''ログ書き出しマクロのメイン関数

    ''画面の更新をオフにすることで処理速度を上げる
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ''変数の定義と初期化
    Dim StrPath As String
    Dim DirectoryPath As String
    
    ''対象のログが格納されているディレクトリを指定する
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = True Then
            MsgBox .SelectedItems(1)
        End If
        
        StrPath = .SelectedItems(1)
    End With
    
    ''サブ関数をコールする
    Call Sheet2.LogReadWrite_1(DirectoryPath)
    
    ''画面の更新をオンに戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    

End Sub

