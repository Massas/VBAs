Sub ExecuteMacros()

''���O�����o���}�N���̃��C���֐�

    ''��ʂ̍X�V���I�t�ɂ��邱�Ƃŏ������x���グ��
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ''�ϐ��̒�`�Ə�����
    Dim StrPath As String
    Dim DirectoryPath As String
    
    ''�Ώۂ̃��O���i�[����Ă���f�B���N�g�����w�肷��
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = True Then
            MsgBox .SelectedItems(1)
        End If
        
        StrPath = .SelectedItems(1)
    End With
    
    ''�T�u�֐����R�[������
    Call Sheet2.LogReadWrite_1(DirectoryPath)
    
    ''��ʂ̍X�V���I���ɖ߂�
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    

End Sub

