Sub LogReadWrite_1(ByVal DirectoryPath As String)

    ''�t�@�C���̓ǂݏo���E�������݂��s���֐�
    
    ''�ϐ���`
    Dim OpenFileName As String
    Dim buf As String
    Dim n As Long
    
    ''��ʂ̍X�V���I�t�ɂ��邱�Ƃŏ������x���グ��
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ''�G�N�Z���V�[�g�̍s���w��
    n = 5
    
    OpenFileName = DirectoryPath & "/log1.txt"
    Open OpenFileName For Input As #1
    
        Do Until EOF(1)
            Line Input #1, buf
            n = n + 1
            Cells(n, 1) = buf
        Loop
    Close #1
    
    ''�f�[�^��������J���}�ŋ�؂�֐����R�[��
    divideByComma
    
End Sub


Sub divideByComma()

    ''�f�[�^��������J���}�ŋ�؂�֐�
    
    ''�ϐ���`
    Dim ws As Worksheet
    
    Set ws = Worksheets("1")
    
    ''�J���}�Ńf�[�^����؂�
    ws.Columns("A").TextToColumns Comma:=True

End Sub
