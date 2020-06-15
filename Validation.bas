Attribute VB_Name = "Validation"
Option Explicit
'���K�\���Œl�̃`�F�b�N���s���̂ŁA�p�^�[�����Œ�ɂ��Ă���
'�Ј��ԍ��̐��K�\���@���p����0-9���Œ�T�ō��U��

Private Const EmployeeNumberPattern As String = "^[0-9]{5,6}$"

'���O�̐��K�\�� _
\u3040-\u309F�[ �Ђ炪�Ȃ�\���Ă���@��-��[ _
\u30A0-\u30FF �J�^�J�i��\���Ă���@�@-�S _
\u4E00-\u9FFF ������\���Ă���@��-ꞁX _
A-Za-Z ���p�A���t�@�x�b�g�啶����������\���Ă��� _
+ �Œ�P������\��
Private Const NamePattern As String = "^[\u3040-\u309F�[\u30A0-\u30FF\u4E00-\u9FFFA-Za-z]+"

'�q��^�C���V�[�g�̎��Ԑ��K�\��
'0-9 1����3���܂�(�K�{)�@�����_��0-9�@1����2���܂�(�C��) _
���̕\���Ńo���f�[�V�����ł��Ă���͈�
Private Const CtsWorkingHoursPattern As String = "^\d{1,3}(\.\d{1,2})?$"

'Socia�̎��Ԑ��K�\��
'0-9 1����3���܂�:0-5��0-9(�K�{) :0-5 0-9(�C��)
Private Const SocWorkingHoursPattern As String = "^[0-9]{1,3}:[0-5][0-9](:[0-5][0-9])?$"


'�f�[�^�̃t�H�[�}�b�g���`�F�b�N���鏈��
'
'�߂�l True�F�f�[�^�s���Ȃ� False�F�f�[�^�ɕs������
'
Function ValidateData() As Boolean

    '�s��ێ�����ϐ�
    Dim Row As Long
    '�G���[�̐���ێ�����
    Dim ErrorCount As Long: ErrorCount = 0
    
    '�L���ȎЈ��ԍ����ǂ�����ێ�����
    Dim IsValidEmployeeNum As Boolean
    
    '�L���Ȗ��O���ǂ�����ێ�����
    Dim IsValidName As Boolean
    
    '�L���ȘJ�����Ԃ��ǂ�����ێ�����
    Dim IsValidWorkingHours As Boolean
    
    '�L���ȃf�[�^���ǂ�����ێ�����
    Dim IsValidData As Boolean
    
    'Main���W���[�����f�[�^���J�n����s�� _
     �V�[�g��ōŌ�̍s���擾�����[�v����
     
    For Row = main.FirstDataRow To main.GetLastRowInSheet
    
        '�Ј��ԍ��̃t�H�[�}�b�g���`�F�b�N
        IsValidEmployeeNum = ValidateEmployeeNumber(Row)
        
        '�Ј����̃t�H�[�}�b�g���`�F�b�N
        IsValidName = ValidateEmployeeName(Row)
        
        '�J�����Ԃ̃t�H�[�}�b�g���`�F�b�N
        IsValidWorkingHours = ValidateWorkingHoursFormat(Row)
        
        '�Ј��ԍ��A�����A���Ԃ̃t�H�[�}�b�g�����������Ƃ��`�F�b�N
        If IsValidEmployeeNum And _
           IsValidName And _
           IsValidWorkingHours Then
           '�Ј��ԍ��A�����A���Ԃ̃t�H�[�}�b�g���������Ƃ�
           '�q�摤��Socia���̈�v�𔻒�
           IsValidData = CompareDatas(Row)
        End If
        
        '�f�[�^�����ׂĐ��������Ƃ��`�F�b�N
        If Not IsValidEmployeeNum Or _
           Not IsValidName Or _
           Not IsValidWorkingHours Or _
           Not IsValidData Then
            '���͂���Ă���f�[�^����ł��������Ȃ��ꍇ
            '�G���[�J�E���^�[������₷
            ErrorCount = ErrorCount + 1
        End If
    Next
    
    '�G���[�̐���0���ǂ�����Ԃ�
    ValidateData = ErrorCount = 0
End Function


'�Ј��ԍ��̃t�H�[�}�b�g���`�F�b�N���鏈��
'
'���� Row ���������s����s
'
'�߂�l True�F��v False�F�s��v
'
Private Function ValidateEmployeeNumber(Row As Long) As Boolean
    
    '�G���[�̐���ێ�����
    Dim ErrorCount As Integer: ErrorCount = 0
    
    '�q�摤�̎Ј��ԍ��̃Z����I�� Row�͍�����s
    Cells(Row, main.CtsEmployeeNumColumn).Select
    
    'ValidatePattern���Ăяo���āA�Z���̒l�̃t�H�[�}�b�g���`�F�b�N
    If Not ValidatePattern(EmployeeNumberPattern, ActiveCell.Value) Then
        
        '�G���[�J�E���^�[�̑��₷
        ErrorCount = ErrorCount + 1
        
        '�I������Ă���Z���ɁA�G���[�̃X�^�C����K�p����
        Call main.ApplyErrorStyle(ActiveCell)
        
    End If
    
    'Socia���̎Ј��ԍ��̃Z����I�� Row�͍�����s
    Cells(Row, main.SocEmployeeNumColumn).Select
    
    'ValidatePattern���Ăяo���āA�Z���̒l�̃t�H�[�}�b�g���`�F�b�N
    If Not ValidatePattern(EmployeeNumberPattern, ActiveCell.Value) Then
        
        '�G���[�J�E���^�[�̑��₷
        ErrorCount = ErrorCount + 1
        '�I������Ă���Z���ɁA�G���[�̃X�^�C����K�p����
        Call main.ApplyErrorStyle(ActiveCell)
        
    End If
    
    '�G���[�̐���0���ǂ�����Ԃ�
    ValidateEmployeeNumber = ErrorCount = 0
End Function


'�����̃t�H�[�}�b�g���`�F�b�N���鏈��
'
'���� Row ���������s����s
'
'�߂�l True�F��v False�F�s��v
'
Private Function ValidateEmployeeName(Row As Long) As Boolean

    '�G���[�̐���ێ�����
    Dim ErrorCount As Integer: ErrorCount = 0
    
    '�q�摤�̎����̃Z����I�� Row�͍�����s
    Cells(Row, main.CtsEmployeeNameColumn).Select
    
    'ValidatePattern���Ăяo���āA�Z���̒l�̃t�H�[�}�b�g���`�F�b�N
    If Not ValidatePattern(NamePattern, ActiveCell.Value) Then
        
        '�G���[�J�E���^�[�̑��₷
        ErrorCount = ErrorCount + 1
        
        '�I������Ă���Z���ɁA�G���[�̃X�^�C����K�p����
        Call main.ApplyErrorStyle(ActiveCell)
    End If
    
    'Socia���̎����̃Z����I�� Row�͍�����s
    Cells(Row, main.SocEmployeeNameColumn).Select
    
    'ValidatePattern���Ăяo���āA�Z���̒l�̃t�H�[�}�b�g���`�F�b�N
    If Not ValidatePattern(NamePattern, ActiveCell.Value) Then
        
        '�G���[�J�E���^�[�̑��₷
        ErrorCount = ErrorCount + 1
        
        '�I������Ă���Z���ɁA�G���[�̃X�^�C����K�p����
        Call main.ApplyErrorStyle(ActiveCell)
    End If
    
    '�G���[�̐���0���ǂ�����Ԃ�
    ValidateEmployeeName = ErrorCount = 0

End Function


'�J�����Ԃ̃t�H�[�}�b�g���`�F�b�N���鏈��
'
'���� Row ���������s����s
'
'�߂�l True�F��v False�F�s��v
'
Private Function ValidateWorkingHoursFormat(Row As Long) As Boolean
    
    '�G���[�̐���ێ�����
    Dim ErrorCount As Integer: ErrorCount = 0
    
    '�q�摤�̎��Ԃ̃Z����I�� Row�͍�����s
    Cells(Row, main.CtsWorkingHoursColumn).Select
    
    'ValidatePattern���Ăяo���āA�Z���̒l�̃t�H�[�}�b�g���`�F�b�N
    If Not ValidatePattern(CtsWorkingHoursPattern, ActiveCell.Value) Then
        
        '�G���[�J�E���^�[�̑��₷
        ErrorCount = ErrorCount + 1
        
        '�I������Ă���Z���ɁA�G���[�̃X�^�C����K�p����
        Call main.ApplyErrorStyle(ActiveCell)
        
    End If
    
    'Socia���̎��Ԃ̃Z����I�� Row�͍�����s
    Cells(Row, main.SocWorkingHoursColumn).Select
    
    'ValidatePattern���Ăяo���āA�Z���̒l�̃t�H�[�}�b�g���`�F�b�N
    If Not ValidatePattern(SocWorkingHoursPattern, ActiveCell.Text) Then
        
        '�G���[�J�E���^�[�̑��₷
        ErrorCount = ErrorCount + 1
        
        '�I������Ă���Z���ɁA�G���[�̃X�^�C����K�p����
        Call main.ApplyErrorStyle(ActiveCell)
        
    End If
    
    '�G���[�̐���0���ǂ�����Ԃ�
    ValidateWorkingHoursFormat = ErrorCount = 0
    
End Function


'�q��^�C���V�[�g��Socia�̃f�[�^�̈�v���`�F�b�N(�Ј��ԍ��Ǝ����̂�)
'
'���� Row ���������s����s
'
'�߂�l True�F��v False�F�s��v
'
Private Function CompareDatas(Row As Long) As Boolean
    
    '�f�[�^�̈�v��Ԃ�ێ�
    Dim IsDataMatch As Boolean: IsDataMatch = True
    
    '�Ј��ԍ��Ǝ����̈�v���`�F�b�N
    If (Cells(Row, main.CtsEmployeeNumColumn) <> Cells(Row, main.SocEmployeeNumColumn)) Or _
        (Cells(Row, main.CtsEmployeeNameColumn) <> Cells(Row, main.SocEmployeeNameColumn)) Then
        '�ǂ��炩��ł���v���Ȃ��ꍇ
        
        '�s��v��ێ�
        IsDataMatch = False
        
        '�s�ɃG���[�̃X�^�C����K�p
        Call main.ApplyErrorStyle(main.GetDataRowRange(Row))
        
    End If
    
    '�f�[�^�̈�v��Ԃ�Ԃ��@True�F��v�@False�F�s��v
    CompareDatas = IsDataMatch
    
End Function


'�w�肳�ꂽ���K�\�����g���ăf�[�^�̑Ó������`�F�b�N
'
'�����@ValidatePattern�F���K�\���̃p�^�[��(String�^)
'�����@Data�F�`�F�b�N������f�[�^(String�^)
'
'�߂�l�@True�F��v�@False�F�s��v
'
Function ValidatePattern(ByVal ValidationPattern As String, ByVal Data As String) As Boolean
    
    '���K�\�����g�p�ł���悤�ɐݒ�
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    '���K�\���I�u�W�F�N�g�̐ݒ�
    With regex
        .Global = True '������S�̂Ō�������悤�ɐݒ�
        .IgnoreCase = False  '�啶���Ə���������ʂ���悤�ɐݒ�
        .Pattern = ValidationPattern '���K�\���̃p�^�[����ݒ�
    End With
    
    '���K�\�����}�b�`���O
    ValidatePattern = regex.Test(Data)
    
End Function

'�ϊ���̘J�����Ԃ���v���Ă��邱�Ƃ��`�F�b�N���鏈��
'
'���� Row: ��r���������s����s��\��
'
'�߂�l�@True�F��v�@False:�s��v
Function ValidateWorkingHours(Row As Long) As Boolean

    '�q��^�C���V�[�g�̎��ԂƕύX���̎��Ԃ���v���Ă��邩�ǂ�����Ԃ�
    ValidateWorkingHours = Cells(Row, main.CtsWorkingHoursColumn) = Cells(Row, main.ConvertedColumn)
    
End Function
