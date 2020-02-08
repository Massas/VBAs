Attribute VB_Name = "PerfMacro_TableSpace"
'============================
'  ���\����̓v���O����
'============================
' �y�@�\�z
'  1�Fcsv�t�@�C����ǂݍ��݁A�Z���Ƀv���b�g����
'  2�F�v���b�g�����f�[�^����Ɍv�Z���s���A�ʃV�[�g�ɂ��̒l���v���b�g����

Sub PerfMacro()

    ' Application.ScreenUpdating = False
    ' Application.Calculation = xlCalculationManual

    '=========================
    '  �ϐ���`
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

    ' ���[�N�V�[�g���ݒ�
    dataSheetName = "�y�f�[�^�zmon_get_tablespace"
    calcSheetName = "�y�v�Z�zmon_get_tablespace"
    resultSheetName = "�y���ʁzmon_get_tablespace"

    Worksheets(1).Name = dataSheetName

    ' ���[�N�V�[�g�̏���������
    ' �O��F���[�N�V�[�g�����O�Ɏ蓮�ō쐬���Ă��邱��
    ' TODO�F���[�N�V�[�g�쐬�֐�����������

    '=========================
    '  �ϐ���`
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

    ' csv�t�@�C���Ǎ�����
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            MsgBox .SelectedItems(1)
        End If
        
        dirPath = .SelectedItems(1)
    End With

    ' csv�t�@�C���擾����
    filePathBuff = Dir(dirPath & "\\" & "mon_get_tablespace*.csv")
    
    n = 1
    
    Do While filePathBuff <> ""
        openFileName = dirPath & "\" & filePathBuff
        
        Open openFileName For Input As #1
        
        ' �t�@�C�������܂Ń��[�v����
        Do Until EOF(1)
        
            ' �s�S�̂�ϐ��Ɋi�[
            ' ���s�R�[�hLF�̓G�N�Z���ł͉��s�R�[�h�ƔF������Ȃ��ׁA���s���ꂸ�Ɋi�[����鋓���ƂȂ�
            Line Input #1, buf
            
            ' ���s�R�[�hLF����؂蕶���Ƃ��čs�𕪊����A�z��Ɋi�[����
            bufTmp = Split(buf, vbLf)
            
            For i = 0 To UBound(bufTmp) - 1
                ' LF�ŋ�؂����s���X�ɃJ���}�ŋ�؂��Ĕz��Ɋi�[����
                bufSplited = Split(bufTmp(i), ",")
                
                Cells(n, 1).Resize(1, UBound(bufSplited) + 1).Value = bufSplited
                
                n = n + 1
            Next i
        Loop
            
        Close #1
        
        filePathBuff = Dir()
    
    Loop

    ' �s�v�f�[�^�폜����
    ' �e�s��A��̃Z����"TIMESTAMP"�Ƃ����w�b�_��񂪊܂܂�邩�����񌟍����s��
    ' ���̏�����A1�Z������n�߂ă��[�N�V�[�g�̃f�[�^�����݂���s�������[�v��������
    ' "TIMESTAMP"���������ꂽ�ꍇ�A���̍s�S�̂��폜���ăZ������ɋl�߂�
    ' �A���A1�x�ڂɔ������ꂽ���͕̂K�v�ȃw�b�_���Ȃ̂ō폜����X�L�b�v����
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

    ' �K�v�ȏ�񂾂���ʂ̃��[�N�V�[�g�ɃR�s�[���鏈��
    Dim dataArray As Variant                     ' �R�s�[����f�[�^���ڂ����z��B�z��ɐݒ肵�����ԂŃR�s�[����
    Dim columnNum As Long
    Dim arrayNum As Long
    Dim rowNum As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long

    ' �f�[�^���ڂ�z��ɃZ�b�g����
    dataArray = Array("TIMESTAMP", "TBSP_NAME", "TBSP_USED_PAGES", "TBSP_TOTAL_PAGES")

    arrayNum = UBound(dataArray, 1) + 1

    ' Copy���\�b�h�Ń��[�N�V�[�g�ɃR�s�[����
    Worksheets(dataSheetName).Cells.Select
    Selection.Copy
    Worksheets(calcSheetName).Cells.PasteSpecial (xlPasteFormulas)
    Worksheets(calcSheetName).Activate

    rowNum = Cells(Rows.Count, 1).End(xlUp).row
    
    columnNum = Cells(1, 1).End(xlToRight).column

    ' ���̗񂩂珇�ԂɃw�b�_���𑖍����āA�z��ɗv�f�����݂��邩�m�F����
    For k = 1 To arrayNum
        For l = 1 To columnNum

            ' �z��̗v�f�ƃf�[�^���ڂ�˂����킹��
            ' �������ꍇ�ɂ̓��[�N�V�[�g��k�Ԗڂ̗�ɐV�������}�����Ă���A��̓��e���R�s�y����
            If dataArray(k - 1) = Worksheets(calcSheetName).Cells(1, l).Value Then
                Worksheets(calcSheetName).Columns(k).Insert
                Worksheets(calcSheetName).Columns(l + 1).Copy
                Worksheets(calcSheetName).Columns(k).PasteSpecial (xlPasteFormulas)
                Worksheets(calcSheetName).Columns(l + 1).Delete
            End If
        Next l
    Next k

    ' �s�v�ȃf�[�^���폜����
    For m = arrayNum To columnNum
        Worksheets(calcSheetName).Columns(arrayNum + 1).Delete
    Next m

    ' ������f�[�^�𐔒l�f�[�^�ɕϊ����鏈��
    Dim str As String
    Dim num As Double
    Dim maxCol As Long
    Dim cellRange As Range
    Dim findKey As String                        ' FIND�֐��̌���������

    ' �ő��ԍ����擾����
    maxCol = Worksheets(calcSheetName).Cells(1, Columns.Count).End(xlToLeft).column

    ' �����������ݒ肷��
    ' ASCII�R�[�h�Ń_�u���N�I�[�e�[�V�����́u34�v
    findKey = Chr(34)

    For j = 1 To rowNum
        If j = 1 Then
            j = j + 1
        End If

        For i = 1 To maxCol
            ' ����������𕔕���v��������Range�I�u�W�F�N�g���擾����
            Set cellRange = Worksheets(calcSheetName).Cells(j, i).Find(What:=findKey, LookIn:=xlValues, Lookat:=xlPart)

            If (cellRange Is Nothing) = True Then
                ' ������f�[�^�𐔒l�f�[�^�ɕϊ�����
                str = Worksheets(calcSheetName).Cells(j, i).Value
                num = Val(str)
                Worksheets(calcSheetName).Cells(j, i).Value = num
            End If

            Set cellRange = Nothing
        Next i

    Next j

' �f�[�^�v�Z�p�֐����R�[������
calcTableSpaceFreeRatio

' TODO �O���t�쐬�֐�����������

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

' �e�[�u���X�y�[�X�󂫗̈旦�v�Z�֐�
Sub calcTableSpaceFreeRatio()

    Dim dataHeaderArray As Variant               ' ���ʃV�[�g�̗�w�b�_�[���i�[����z��
    Dim timeStampArray As Variant                ' TIMESTAMP�̒l��v�f�Ƃ���z��
    Dim dataResultArray As Variant               ' ���ʃV�[�g�̊e�Z����\��2�����z��
    Dim calcSheetName As String
    Dim resultSheetName As String
    Dim arrayNum As Long
    Dim rowNum As Long
    Dim i As Long                                ' TIMESTAMP�̒l���i�[����z��𐧌䂷��ׂ̕ϐ�
    Dim j As Long                                ' TIMESTAMP�z��̍ő�l
    Dim n As Long                                ' TIMESTAMP�z��̗v�f�ԍ�
    Dim k As Long                                ' ���ʃ��[�N�V�[�g�̃w�b�_��
    Dim l As Long                                ' ���ʃ��[�N�V�[�g�ɔ��f����TIMESTAMP�̍s��
    Dim m As Long                                ' ���ʃ��[�N�V�[�g�̃��R�[�h�������[�v��������ׂ̕ϐ�


    ' ���ʃ��[�N�V�[�g�̃w�b�_��ݒ肷��
    dataHeaderArray = Array("TIMESTAMP", "SYSCATSPACE")

    calcSheetName = "�y�v�Z�zmon_get_tablespace"
    resultSheetName = "�y���ʁzmon_get_tablespace"

    arrayNum = UBound(dataHeaderArray, 1)
    rowNum = Cells(Rows.Count, 1).End(xlUp).row

    ' ���ʃ��[�N�V�[�g�̃w�b�_�������쐬����
    ' TIMESTAMP�̗񂩂�f�[�^�擾���Ԃ𔲂��o���Ĕz��Ɋi�[����
    ' ���[�N�V�[�g��ł�TIMESTAMP��2�s�ڂ���v���b�g����Ă���
    i = 2
    j = rowNum                         ' TODO �����̒l��ϐ�������
    ReDim timeStampArray(j)

    For n = 1 To j
        timeStampArray(n) = Cells(i, 1).Value
        ' i = i + 2
    Next n

    Worksheets(resultSheetName).Activate

    ' ���ʃ��[�N�V�[�g�Ƀw�b�_�������
    For k = 1 To arrayNum + 1
        Worksheets(resultSheetName).Cells(1, k).Value = dataHeaderArray(k - 1)
    Next k

    ' TIMESTAMP�񂩂猋�ʃ��[�N�V�[�g�̌��ʗ�Ƀv���b�g����
    For l = 1 To j - 1
        Worksheets(resultSheetName).Cells(l + 1, 1).Value = timeStampArray(l)
    Next l

    Worksheets(calcSheetName).Activate

    ' �v�Z���ʂ��i�[����2�����z��̗̈��ύX����
    ReDim dataResultArray(j, (UBound(dataHeaderArray, 1)))

    ' �f�[�^���ڂ̌v�Z���s��
    Dim row As Long
    Dim column As Long
    Dim numerator As Double
    Dim denominator As Double

    For m = 1 To rowNum - 1

        ' 2�����z��̗񕪏���������A�s�̒l��2�����z��Ɋi�[����
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
            ' ����E���q�̌v�Z
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

        ' �v�Z�������ʂ�2�����z��Ɋi�[����
        For l = 1 To j - 1
            For k = 1 To UBound(dataHeaderArray, 1)
                Worksheets(resultSheetName).Cells(l + 1, k + 1) = dataResultArray(l, k)
            Next k
        Next l

    End Sub

