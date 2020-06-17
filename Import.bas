Attribute VB_Name = "Import"
Option Explicit

'�q��^�C���V�[�g�̃t�@�C�����̃e���v���[�g
'�q��^�C���V�[�g.csv�ŏI�����̂��ׂē��Ă͂܂�
Private Const CtsFileNameFormat = "*�q��^�C���V�[�g.csv"


'Socia�^�C���V�[�g�̃t�@�C�����̃e���v���[�g
'Socia.csv�ŏI�����̂��ׂē��Ă͂܂�
Private Const SocFileNameFormat = "*Socia.csv"


'�C���|�[�g�̏���
Sub ImportFiles()

    '�t�H���_�̃p�X��ێ�
    Dim FolderPath As String
    
    '�ŐV�̋q��^�C���V�[�g�t�@�C������ێ�
    Dim LatestCtsFile As String
    
    '�ŐV��Socia�^�C���V�[�g�t�@�C������ێ�
    Dim LatestSocFile As String

    '���Ӄ��b�Z�[�W��\��
    Call MsgBox("�C���|�[�g���s���O�ɁA�ȉ��̃t�H�[�}�b�g��CSV�t�@�C�����t�H���_���ɂ��邱�Ƃ����m�F���������B" _
        & vbNewLine & vbNewLine _
        & "�u�q��^�C���V�[�g�v�Ə����ꂽCSV�t�@�C�������݂��邱��" _
        & vbNewLine & "�uSocia�v�Ə����ꂽCSV�t�@�C�������݂��邱��" _
        & vbNewLine & vbNewLine & "���t�@�C������������ꍇ�͍ŐV�̂��̂��g�p���܂��B", _
        vbOKOnly + vbInformation, _
        "�t�@�C���C���|�[�g�m�F����" _
    )
    
    '�C���|�[�g�t�@�C���������Ă���t�H���_�̃p�X���擾����
    FolderPath = AskFolderPath()

    '�L�����Z���Ȃǂ̑Ή��@�������~�߂�
    If FolderPath = "" Then GoTo Finish
    
    '�t�H���_���ŐV�̋q��^�C���V�[�g�̃t�@�C�������擾
    LatestCtsFile = GetLatestFile(FolderPath, CtsFileNameFormat)
    
    '�t�H���_���ŐV��Socia�^�C���V�[�g�̃t�@�C�������擾
    LatestSocFile = GetLatestFile(FolderPath, SocFileNameFormat)
    
    '�t�@�C�������邱�ƁA�����t�@�C����I�����Ă��Ȃ����ƃ`�F�b�N
    If LatestCtsFile = "" Or LatestSocFile = "" Or LatestCtsFile = LatestSocFile Then
        Call MsgBox("�����𖞂����t�@�C�����t�H���_���Ɍ�����܂���ł����B" _
            & vbNewLine & "�t�H���_�̓��e�y�уp�X���m�F���A������x��蒼���Ă��������B", _
            vbOKOnly + vbCritical, _
            "�t�@�C����������܂���ł����B" _
        )
        GoTo Finish
    End If
    
    '�q��^�C���V�[�g�f�[�^���C���|�[�g
    Call ImportCtsValues(FolderPath, LatestCtsFile)
    
    'Socia�^�C���V�[�g�f�[�^���C���|�[�g
    Call ImportSocValues(FolderPath, LatestSocFile)
        
Finish:

End Sub


'�t�H���_�̃p�X���擾���鏈��
'
'�߂�l ���[�U�[���I�������t�H���_�ւ̃p�X
'
Private Function AskFolderPath() As String
    
    '�t�H���_�[�p�X��ێ�
    Dim FolderPath As String
    
    '�t�H���_�[�s�b�J�[�𐶐�
    With Application.FileDialog(msoFileDialogFolderPicker)
        '�����I���𖳌�
        .AllowMultiSelect = False
        '�J�����Ƃ��ŏ��ɕ\�������ꏊ���w��@���[�U�[�̃t�H���_(���ϐ����)
        .InitialFileName = Environ("USERPROFILE") & Application.PathSeparator
        
        '�I��=>OK���I�����ꂽ�Ƃ��̏���
        If .Show = -1 Then FolderPath = .SelectedItems(1)
    End With

    '�擾�����p�X��Ԃ�
    AskFolderPath = FolderPath
End Function


'�ŐV�̃t�@�C�����擾����
'
'���� Folder�F�t�H���_�ւ̃p�X FileName�F����������t�@�C����(����̓t�@�C�����e���v���[�g���g�p)
'
'�߂�l �ŐV�t�@�C����
'
Private Function GetLatestFile(ByVal Folder As String, FileName As String) As String

    '�ŐV������ێ�
    Dim LatestDate As Date
    
    '�ŐV�t�@�C��
    Dim LatestFile As String
    
    '�t�@�C����
    Dim File As String
    
    '�p�X�̍Ō�ɃZ�p���[�^�[��ǉ�(Windows�́�Mac,Linux�ł�/)
    Folder = Folder & Application.PathSeparator
    
    'Dir�Ŏw�肵���A�p�X�̂��̂��Ƃ��Ă���
    File = Dir(Folder & FileName)
    
    '�t�H���_���̃e���v���[�g�ɊY������t�@�C���������[�v
    Do While File <> ""
        
        '�t�@�C���̓������ŐV���V���������`�F�b�N
        If FileDateTime(Folder & File) > LatestDate Then
            
            '�ŐV�t�@�C���Ƀt�@�C����ݒ�
            LatestFile = File
            
            '�ŐV�������X�V
            LatestDate = FileDateTime(Folder & File)
                
        End If
        
        '���̃t�@�C�����擾
        File = Dir
    Loop
    
    '�ŐV�̃t�@�C������Ԃ�
    GetLatestFile = LatestFile
End Function


'�q��^�C���V�[�g�̒l���C���|�[�g����
'
'���� Path�F�t�@�C���܂ł̃p�X File�F�t�@�C����
'
Private Sub ImportCtsValues(Path As String, File As String)
        
    '�s��ێ�
    Dim Row As Long
    
    'CSV�̍s��z��ɂ������̂�ێ�
    Dim ArrLine As Variant
    
    '�t�@�C����ǂݍ��ނ̂ɕK�v�ȃI�u�W�F�N�g
    Dim Fso As New Scripting.FileSystemObject
    
    '�s�Ƀf�[�^���n�܂�s��ݒ�
    Row = main.FirstDataRow
    
    '�t�@�C�����J�� ForReading=>�ǂݍ��ݐ�p
    With Fso.OpenTextFile(Path & Application.PathSeparator & File, ForReading)
            
        '�ŏI�s�ł͂Ȃ��ꍇ��s��΂�
        '�t�@�C�����J��������Ȃ̂ŁA1�s�ڂ��΂��Ă���
        If Not .AtEndOfLine Then .SkipLine
        
        '�ŏI�s�ɂȂ�Ȃ��ԏ����𑱍s
        Do Until .AtEndOfLine
            '1�s�ǂݍ���Łu,�v�ŋ�؂�
            ArrLine = Split(.ReadLine, ",")
            
            '�K�v�ȃf�[�^��ǂݍ���
            Cells(Row, main.CtsEmployeeNumColumn).Value = ArrLine(1)
            Cells(Row, main.CtsEmployeeNameColumn).Value = ArrLine(2)
            '���Ԃ̌v�Z������
            Cells(Row, main.CtsWorkingHoursColumn).Value = CalculateCtsHours(ArrLine)
            
            '�s������炷
            Row = Row + 1
        Loop
        
        '�t�@�C�������
        .Close
    End With
End Sub


'Socia�^�C���V�[�g�̒l���C���|�[�g����
'
'���� Path�F�t�@�C���܂ł̃p�X File�F�t�@�C����
'
Private Sub ImportSocValues(Path As String, File As String)

    '�s��ێ�
    Dim Row As Long
    
    'CSV�̍s��z��ɂ������̂�ێ�
    Dim ArrLine As Variant
    
    '�t�@�C����ǂݍ��ނ̂ɕK�v�ȃI�u�W�F�N�g
    Dim Result As Variant
    
    '�q��Ј��ԍ��̍Ō�̍s��ێ�
    Dim LastCtsNumRow As Long
    
    '�t�@�C����ǂݍ��ނ̂ɕK�v�ȃI�u�W�F�N�g
    Dim Fso As New Scripting.FileSystemObject
    
    '�s�Ƀf�[�^���n�܂�s��ݒ�
    Row = main.FirstDataRow
    
    '�q��Ј��ԍ��̍Ō�̋Ȃ��擾
    LastCtsNumRow = main.GetLastRowInColumn(main.CtsEmployeeNumColumn)
    
    '�t�@�C�����J�� ForReading=>�ǂݍ��ݐ�p
    With Fso.OpenTextFile(Path & Application.PathSeparator & File, ForReading)
     
        '�ŏI�s�ł͂Ȃ��ꍇ��s��΂�
        '�t�@�C�����J��������Ȃ̂ŁA1�s�ڂ��΂��Ă���
        If Not .AtEndOfStream Then .SkipLine
            
        '�ŏI�s�ɂȂ�Ȃ��ԏ����𑱍s
        Do Until .AtEndOfStream
            
            '1�s�ǂݍ���Łu,�v�ŋ�؂�
            ArrLine = Split(.ReadLine, ",")
                
            '�Ј��ԍ����q��ŃC���|�[�g�������̂ƈ�v���Ă��邩��T��
            Result = Application.Match( _
                CLng(ArrLine(0)), _
                Range( _
                    Cells(main.FirstDataRow, main.CtsEmployeeNumColumn), _
                    Cells(LastCtsNumRow, main.CtsEmployeeNumColumn)), _
                0)
            
            '�����̌��ʃG���[�łȂ���΃f�[�^��ǂݍ���
            If Not IsError(Result) Then
                Cells(Result + 2, main.SocEmployeeNumColumn).Value = ArrLine(0)
                Cells(Result + 2, main.SocEmployeeNameColumn).Value = ArrLine(1)
                '���Ԃ��v�Z���ēǂݍ���
                Cells(Result + 2, main.SocWorkingHoursColumn).Value = CalculateSocHours(ArrLine)
            End If

        Loop
    
    End With
    
End Sub


'�q��^�C���V�[�g�̎��Ԃ��v�Z
'
'���� Data�FCSV1�s���̃f�[�^���������z��
'
'�߂�l ���J������
'
Private Function CalculateCtsHours(Data As Variant) As Double

    '�J�����Ԃ�ێ�
    Dim TotHours As Double
    
    '�J�E���^�[
    Dim I As Integer
    
    For I = 5 To 9
    
        '6�̏ꍇ�X�L�b�v
        If I = 6 Then GoTo Continue
        
        '�����񂩂琔���ɕϊ����Ēl�𑫂�
        TotHours = TotHours + CDbl(Data(I))
Continue:
    Next
    
    '���v��߂�
    CalculateCtsHours = TotHours
    
End Function


'Soc�^�C���V�[�g�̎��Ԃ��v�Z
'
'���� Data�FCSV1�s���̃f�[�^���������z��
'
'�߂�l ���J������
'
Private Function CalculateSocHours(Data As Variant) As Double
    
    '���ʂ�ێ�
    Dim Result As Double
    
    '�J�E���^�[
    Dim I As Integer
    
    '6�ڂ�CSV�f�[�^���Ԃ�ϊ�
    Result = ConvertTimeStrToDbl(Data(6))
    
    '�c��̃f�[�^���ϊ����Čv�Z����
    For I = 7 To 9
        Result = Result - ConvertTimeStrToDbl(Data(I))
        
    Next
    
    '���v��߂�
    CalculateSocHours = Result
End Function

'���Ԃ������ɕϊ�(Excel�̃t�H�[�}�b�g)
'
'���� Data�FCSV1�s���̃f�[�^���������z��
'
'�߂�l ���������ꂽ����
'
Private Function ConvertTimeStrToDbl(ByVal Value As String) As Double
    
    '���ʂ�ێ�
    Dim Result As Double
    
    '���[�v�Ŏg�p����I�u�W�F�N�g
    Dim Data As Variant
    
    '�v�Z�Ɏg���l��ێ�
    Dim tmp As Double
    
    'Excel��24���Ԃ�1�Ƃ��ĕێ�����̂ŁA�܂�24������
    tmp = 24
    
    '���ʂ̏����l
    Result = 0
    
    '���Ԃ̒l���u:�v�ŋ�؂��Ĉ����Data�ɓ���ď��ɏ���
    '���ԁ@�v�Z�́@����/24
    '���@�@�v�Z�́@��/24*60
    '�b�@�@�v�Z�́@�b/24*60*60
    For Each Data In Split(Value, ":")
        
        '���Ԃ��v�Z���č��v�ɑ���
        Result = Result + Data / tmp
        
        '�v�Z�Ɏg���l���X�V
        tmp = tmp * 60
    Next
        
    '���ʂ�Ԃ�
    ConvertTimeStrToDbl = Result
End Function
