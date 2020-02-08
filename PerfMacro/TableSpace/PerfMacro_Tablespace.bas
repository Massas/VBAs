Attribute VB_Name = "PerfMacro_TableSpace"
'============================
'  性能情報解析プログラム
'============================
' 【機能】
'  1：csvファイルを読み込み、セルにプロットする
'  2：プロットしたデータを基に計算を行い、別シートにその値をプロットする

Sub PerfMacro()

    ' Application.ScreenUpdating = False
    ' Application.Calculation = xlCalculationManual

    '=========================
    '  変数定義
    '=========================
    Dim dirPath As String
    Dim filePathBuff As String
    Dim openFileName As String
    Dim buf As String
    Dim bufTmp As Variant
    Dim bufSplited As Variant
    Dim n As Long
    Dim i As Long
    Dim dataSheetName As String
    Dim calcSheetName As String
    Dim resultSheetName As String

    ' ワークシート名設定
    dataSheetName = "【データ】mon_get_tablespace"
    calcSheetName = "【計算】mon_get_tablespace"
    resultSheetName = "【結果】mon_get_tablespace"

    Worksheets(1).Name = dataSheetName

    ' ワークシートの初期化処理
    ' 前提：ワークシートを事前に手動で作成してあること
    ' TODO：ワークシート作成関数を実装する

    '=========================
    '  変数定義
    '=========================
    Dim ws As Worksheet

    Application.DisplayAlerts = False

    For Each ws In Worksheets
        If ws.Name = dataSheetName Then
            Worksheets(dataSheetName).Activate
            Worksheets(dataSheetName).Cells.Clear
        ElseIf ws.Name = calcSheetName Then
            Worksheets(calcSheetName).Activate
            Worksheets(calcSheetName).Cells.Clear
        ElseIf ws.Name = resultSheetName Then
            Worksheets(resultSheetName).Activate
            Worksheets(resultSheetName).Cells.Clear
        Else
        End If
    Next ws

    Application.DisplayAlerts = True

    Worksheets(dataSheetName).Activate

    ' csvファイル読込処理
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            MsgBox .SelectedItems(1)
        End If
        
        dirPath = .SelectedItems(1)
    End With

    ' csvファイル取得処理
    filePathBuff = Dir(dirPath & "\\" & "mon_get_tablespace*.csv")
    
    n = 1
    
    Do While filePathBuff <> ""
        openFileName = dirPath & "\" & filePathBuff
        
        Open openFileName For Input As #1
        
        ' ファイル末尾までループ処理
        Do Until EOF(1)
        
            ' 行全体を変数に格納
            ' 改行コードLFはエクセルでは改行コードと認識されない為、改行されずに格納される挙動となる
            Line Input #1, buf
            
            ' 改行コードLFを区切り文字として行を分割し、配列に格納する
            bufTmp = Split(buf, vbLf)
            
            For i = 0 To UBound(bufTmp) - 1
                ' LFで区切った行を更にカンマで区切って配列に格納する
                bufSplited = Split(bufTmp(i), ",")
                
                Cells(n, 1).Resize(1, UBound(bufSplited) + 1).Value = bufSplited
                
                n = n + 1
            Next i
        Loop
            
        Close #1
        
        filePathBuff = Dir()
    
    Loop

    ' 不要データ削除処理
    ' 各行のA列のセルに"TIMESTAMP"というヘッダ情報が含まれるか文字列検索を行う
    ' この処理はA1セルから始めてワークシートのデータが存在する行数分ループ処理する
    ' "TIMESTAMP"が発見された場合、その行全体を削除してセルを上に詰める
    ' 但し、1度目に発見されたものは必要なヘッダ情報なので削除するスキップする
    Dim maxRownum As Long
    Dim headerFlag As Boolean
    Dim j As Long
    Dim headerRange As Range
    Dim deleteKey As String
    
    deleteKey = "TIMESTAMP"
    headerFlag = True
    maxRownum = Cells(Rows.Count, 1).End(xlUp).row

    For j = 1 To maxRownum
        Set headerRange = Rows(j).Find(What:=deleteKey, LookIn:=xlValues, Lookat:=xlWhole)

        If Not (headerRange Is Nothing) And headerFlag = True Then
            headerFlag = False
        ElseIf Not (headerRange Is Nothing) And headerFlag = False Then
            Rows(j).EntireRow.Delete
        End If
    Next j

    ' 必要な情報だけを別のワークシートにコピーする処理
    Dim dataArray As Variant                     ' コピーするデータ項目を持つ配列。配列に設定した順番でコピーする
    Dim columnNum As Long
    Dim arrayNum As Long
    Dim rowNum As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long

    ' データ項目を配列にセットする
    dataArray = Array("TIMESTAMP", "TBSP_NAME", "TBSP_USED_PAGES", "TBSP_TOTAL_PAGES")

    arrayNum = UBound(dataArray, 1) + 1

    ' Copyメソッドでワークシートにコピーする
    Worksheets(dataSheetName).Cells.Select
    Selection.Copy
    Worksheets(calcSheetName).Cells.PasteSpecial (xlPasteFormulas)
    Worksheets(calcSheetName).Activate

    rowNum = Cells(Rows.Count, 1).End(xlUp).row
    
    columnNum = Cells(1, 1).End(xlToRight).column

    ' 左の列から順番にヘッダ情報を走査して、配列に要素が存在するか確認する
    For k = 1 To arrayNum
        For l = 1 To columnNum

            ' 配列の要素とデータ項目を突き合わせる
            ' 等しい場合にはワークシートのk番目の列に新しい列を挿入してから、列の内容をコピペする
            If dataArray(k - 1) = Worksheets(calcSheetName).Cells(1, l).Value Then
                Worksheets(calcSheetName).Columns(k).Insert
                Worksheets(calcSheetName).Columns(l + 1).Copy
                Worksheets(calcSheetName).Columns(k).PasteSpecial (xlPasteFormulas)
                Worksheets(calcSheetName).Columns(l + 1).Delete
            End If
        Next l
    Next k

    ' 不要なデータを削除する
    For m = arrayNum To columnNum
        Worksheets(calcSheetName).Columns(arrayNum + 1).Delete
    Next m

    ' 文字列データを数値データに変換する処理
    Dim str As String
    Dim num As Double
    Dim maxCol As Long
    Dim cellRange As Range
    Dim findKey As String                        ' FIND関数の検索文字列

    ' 最大列番号を取得する
    maxCol = Worksheets(calcSheetName).Cells(1, Columns.Count).End(xlToLeft).column

    ' 検索文字列を設定する
    ' ASCIIコードでダブルクオーテーションは「34」
    findKey = Chr(34)

    For j = 1 To rowNum
        If j = 1 Then
            j = j + 1
        End If

        For i = 1 To maxCol
            ' 検索文字列を部分一致検索してRangeオブジェクトを取得する
            Set cellRange = Worksheets(calcSheetName).Cells(j, i).Find(What:=findKey, LookIn:=xlValues, Lookat:=xlPart)

            If (cellRange Is Nothing) = True Then
                ' 文字列データを数値データに変換する
                str = Worksheets(calcSheetName).Cells(j, i).Value
                num = Val(str)
                Worksheets(calcSheetName).Cells(j, i).Value = num
            End If

            Set cellRange = Nothing
        Next i

    Next j

' データ計算用関数をコールする
calcTableSpaceFreeRatio

' TODO グラフ作成関数を実装する

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

' テーブルスペース空き領域率計算関数
Sub calcTableSpaceFreeRatio()

    Dim dataHeaderArray As Variant               ' 結果シートの列ヘッダーを格納する配列
    Dim timeStampArray As Variant                ' TIMESTAMPの値を要素とする配列
    Dim dataResultArray As Variant               ' 結果シートの各セルを表す2次元配列
    Dim calcSheetName As String
    Dim resultSheetName As String
    Dim arrayNum As Long
    Dim rowNum As Long
    Dim i As Long                                ' TIMESTAMPの値を格納する配列を制御する為の変数
    Dim j As Long                                ' TIMESTAMP配列の最大値
    Dim n As Long                                ' TIMESTAMP配列の要素番号
    Dim k As Long                                ' 結果ワークシートのヘッダ列数
    Dim l As Long                                ' 結果ワークシートに反映するTIMESTAMPの行数
    Dim m As Long                                ' 結果ワークシートのレコード数分ループ処理する為の変数


    ' 結果ワークシートのヘッダを設定する
    dataHeaderArray = Array("TIMESTAMP", "SYSCATSPACE")

    calcSheetName = "【計算】mon_get_tablespace"
    resultSheetName = "【結果】mon_get_tablespace"

    arrayNum = UBound(dataHeaderArray, 1)
    rowNum = Cells(Rows.Count, 1).End(xlUp).row

    ' 結果ワークシートのヘッダ部分を作成する
    ' TIMESTAMPの列からデータ取得時間を抜き出して配列に格納する
    ' ワークシート上ではTIMESTAMPは2行目からプロットされている
    i = 2
    j = rowNum                         ' TODO ここの値を変数化する
    ReDim timeStampArray(j)

    For n = 1 To j
        timeStampArray(n) = Cells(i, 1).Value
        ' i = i + 2
    Next n

    Worksheets(resultSheetName).Activate

    ' 結果ワークシートにヘッダ情報を作る
    For k = 1 To arrayNum + 1
        Worksheets(resultSheetName).Cells(1, k).Value = dataHeaderArray(k - 1)
    Next k

    ' TIMESTAMP列から結果ワークシートの結果列にプロットする
    For l = 1 To j - 1
        Worksheets(resultSheetName).Cells(l + 1, 1).Value = timeStampArray(l)
    Next l

    Worksheets(calcSheetName).Activate

    ' 計算結果を格納する2次元配列の領域を変更する
    ReDim dataResultArray(j, (UBound(dataHeaderArray, 1)))

    ' データ項目の計算を行う
    Dim row As Long
    Dim column As Long
    Dim numerator As Double
    Dim denominator As Double

    For m = 1 To rowNum - 1

        ' 2次元配列の列分処理した後、行の値を2次元配列に格納する
        If m = 1 Then
            row = 1
            column = 1
        ElseIf ((m - 1) Mod 2) <> 0 Then
            column = column + 1
        Else
            row = row + 1
            column = 1
        End If

        If m >= 1 Then
            ' 分母・分子の計算
            numerator = Worksheets(calcSheetName).Cells(m + 1, 3).Value
            denominator = Worksheets(calcSheetName).Cells(m + 1, 4).Value

            If numerator = 0 Or denominator = 0 Then
                dataResultArray(row, column) = 100
            Else
                dataResultArray(row, column) = Round((1 - (numerator / denominator)) * 100, 3)
            End If
        End If
    Next m

        Worksheets(resultSheetName).Activate

        ' 計算した結果を2次元配列に格納する
        For l = 1 To j - 1
            For k = 1 To UBound(dataHeaderArray, 1)
                Worksheets(resultSheetName).Cells(l + 1, k + 1) = dataResultArray(l, k)
            Next k
        Next l

    End Sub

