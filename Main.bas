Attribute VB_Name = "Main"
Option Explicit
'�f�[�^���n�܂�s
Public Const FirstDataRow = 3

'�q��^�C���V�[�g - �Ј��ԍ�
Public Const CtsEmployeeNumColumn = 1

'�q��^�C���V�[�g - ����
Public Const CtsEmployeeNameColumn = 2

'�q��^�C���V�[�g - ����
Public Const CtsWorkingHoursColumn = 3

'Socia - �Ј��ԍ�
Public Const SocEmployeeNumColumn = 4

'Socia - ����
Public Const SocEmployeeNameColumn = 5

'Socia - ����
Public Const SocWorkingHoursColumn = 6

'�ϊ���̗�ԍ�
Public Const ConvertedColumn = 7

'�`�F�b�N�̗�ԍ�
Public Const CheckColumn = 8


'�^�C���V�[�g�`�F�b�N�{�^���̃N���b�N����
Sub BtnTimeSheetClick()
    
    '��ʍX�V�𖳌��ɂ���
    Application.ScreenUpdating = False
  
    'ClearFormats���Ăяo���āA�V�[�g�̃t�H�[�}�b�g��������Ԃɖ߂�
    Call ClearFormats
    
    '�f�[�^�̃`�F�b�N���s��
    If Not Validation.ValidateData Then
        '�s���Ȓl������ꍇ���b�Z�[�W�{�b�N�X��\������
        MsgBox _
            "���͓��e�Ɍ�肪��܂��B" & vbNewLine & "�n�C���C�g���ꂽ�Z���E�s�̓��͓��e�����m�F���������B", _
            vbOKOnly + vbCritical, _
            "�o���f�[�V�����G���["
            
        '�X�g�b�v�I�y���[�V�����֍s��
        GoTo StopOp
    End If
    
    '���Ԃ̃`�F�b�N�������Ăяo��
    Call CheckWorkingHours
    
'�X�g�b�v�I�y���[�V�������x��
'�������~�߂�Ƃ��ɂ�����艺�����s
StopOp:

    '��ʍX�V��L���ɂ���
    Application.ScreenUpdating = True
    
    '�V�[�g�̏オ�\�������悤��A1�̃Z����I������
    Cells(1, 1).Select
    
    End Sub
    

'�f�[�^�폜�{�^���̃N���b�N
Sub BtnDeleteClick()
    
    'G3����H�̍ŏI�s�܂ł̃f�[�^������
    Range(Cells(FirstDataRow, ConvertedColumn), _
        Cells(GetLastRowInColumn(CheckColumn), CheckColumn)).ClearContents
        
    '�O���t�֌W�̃N���A�������Ă�
    Call ClearGraph
End Sub


'�G���[�X�^�C����K�p����
'
'���� Cell �X�^�C����K�p����Z���������͔͈�(Range)
'
Sub ApplyErrorStyle(Cell As Range)

    With Cell
        '�w�i�F��Ԃ��ۂ��F�ɕύX
        .Interior.color = RGB(255, 71, 71)
        
        '�����̐F�𔒂ɕύX
        .Font.color = RGB(255, 255, 255)
        
        '�����𑾕����ɕύX
        .Font.Bold = True
    End With

End Sub


'�t�H�[�}�b�g�̏��������s��
Sub ClearFormats()

    'A3����H��V�[�g�̍ŏI�s�܂ł̃t�H�[�}�b�g������
    Range("A3", Cells(Rows.Count, CheckColumn)).ClearFormats
       
    '�q�摤���ԗ�(C��)�ƕϊ����(G��)�̏����ݒ���s��
    Application.Union( _
        Columns(CtsWorkingHoursColumn), _
        Columns(ConvertedColumn)).NumberFormat = "0.00"
        
    'Socia�����ԗ�(F��̏����ݒ���s��
    Columns(SocWorkingHoursColumn).NumberFormat = "[h]:mm"
    
    '�`�F�b�N��(H��)�𒆉������ɐݒ�
    Columns(CheckColumn).HorizontalAlignment = xlCenter
    
    '�q�摤���ԗ�(C��)�̃f�[�^���n�܂�s(3)���瓯��ƃV�[�g�̍ŏI�s�܂ł�
    'Socia�����ԗ�(F��)�̃f�[�^���n�܂�s(3)���瓯��ƃV�[�g�̍ŏI�s�܂�
    '�܂ł͈̔͂ňȉ����������s
    With Application.Union( _
        Range( _
            Cells(FirstDataRow, CtsWorkingHoursColumn), _
            Cells(GetLastRowInSheet, CtsWorkingHoursColumn)), _
        Range( _
            Cells(FirstDataRow, SocWorkingHoursColumn), _
            Cells(GetLastRowInSheet, SocWorkingHoursColumn)))
        
        '�����̐F�𔒐F�ɐݒ�
        .Font.color = RGB(255, 255, 255)
        
        '�����𑾕����ɐݒ�
        .Font.Bold = True
        
        '�w�i�F����ۂ��F�ɕύX
        .Interior.color = RGB(155, 194, 230)
        
    End With
    
End Sub


'�J�����Ԃ̈�v���`�F�b�N���鏈��
Private Sub CheckWorkingHours()

    '�Z���ɕ\������l��ێ�
    Dim Msg As String
    
    '��v����ێ�
    Dim CorrectCount As Long: CorrectCount = 0
    
    '�s��v����ێ�
    Dim WrongCount As Long: WrongCount = 0
    
    'Socia�����ԗ�ƃf�[�^���n�܂�s�̃Z����I��(F3)
    Cells(FirstDataRow, SocWorkingHoursColumn).Select
    
    '�I������Ă���Z������łȂ��Ԉȉ��̏��������s
    Do While ActiveCell <> ""
        
        '�Z���ɕ\������l����ɂ���
        Msg = ""
    
        '�I������Ă���Z���Ɠ����s�ɂ���ϊ����(G��)�ɕϊ������J�����Ԃ�ݒ�
        Cells(ActiveCell.Row, ConvertedColumn).Value = ConvertHoursToDecimal(ActiveCell.Value)
        
        '�ϊ���̘J�����Ԃ��q�摤�ƈ�v���Ă��邱�Ƃ��`�F�b�N
        If Validation.ValidateWorkingHours(ActiveCell.Row) Then
            
            '�Z���ɕ\������l��ݒ�
            Msg = "�Z"
            
            '��v���𑝂₷
            CorrectCount = CorrectCount + 1
        Else
            
            '�Z���ɕ\������l��ݒ�
            Msg = "�~"
            
            '��v���𑝂₷
            WrongCount = WrongCount + 1
        End If
        
        '�`�F�b�N��ɒl�𔽉f
        Cells(ActiveCell.Row, CheckColumn).Value = Msg
        
        '�I������Ă���Z���̈���̃Z����I��
        ActiveCell.Offset(1, 0).Select
    Loop
    
    '�O���t�p�̃f�[�^�𔽉f���鏈�����Ăяo��
    Call PopulateGraphData(CorrectCount, WrongCount)
    
End Sub


'�O���t�Ƀf�[�^�����鏈�����s��
'
'���� CorrectCount ���Ԃɕs�����Ȃ�������
'���� WrongCount ���Ԃɕs������������
'
Private Sub PopulateGraphData(CorrectCount, WrongCount)
    
    '���ꂼ��̃Z���ɒl������
    Cells(3, "N").Value = CorrectCount
    Cells(4, "N").Value = WrongCount
    
    'TODO�@�O���t�̍쐬����
End Sub


'�O���t�̃N���A����
Private Sub ClearGraph()

    '���ꂼ��̃Z���̒l��0�ɐݒ�
    Cells(3, "N").Value = 0
    Cells(4, "N").Value = 0
    
    'TODO�@�O���t�̃f�[�^���N���A
End Sub


'��Ŏg�p����Ă���Ō�̍s���擾���鏈��
'
'���� ColumnIndex ��̔ԍ�
'
'�߂�l �w�肳�ꂽ��Ŏg�p����Ă���ŏI�s�������̓f�[�^���n�܂�s
'
Function GetLastRowInColumn(ColumnIndex) As Long
    
    '�ŏI�s��ێ�
    Dim LastRow As Long
    
    '��̍ŏI�s���擾
    LastRow = Cells(Rows.Count, ColumnIndex).End(xlUp).Row
    
    '�ŏI�s���f�[�^�̊J�n�s�����������ꍇ�̓f�[�^���n�܂�s��Ԃ�
    '����ȊO�͎擾�����l��Ԃ�
    GetLastRowInColumn = IIf(LastRow < FirstDataRow, FirstDataRow, LastRow)
End Function


'�V�[�g�Ŏg�p����Ă���Ō�̍s���擾���鏈��
'
'�߂�l �V�[�g�Ŏg�p����Ă���ŏI�s
'
Function GetLastRowInSheet() As Long
    '�V�[�g�Ŏg�p����Ă���ŏI�s���擾���ĕԂ�
    GetLastRowInSheet = ActiveSheet.UsedRange.Rows.Count
End Function



'�f�[�^�̍s�͈̔͂��擾���� �q��^�C���V�[�g��Socia�̕����̂�
'
'���� Row �͈͂��擾�������s
'
'�߂�l �w�肳�ꂽ�s�̋q�摤�Ј��ԍ��񂩂�Socia�����ԗ�܂ł͈̔�
'
Function GetDataRowRange(ByVal Row As Long) As Range
    '�͈͂��w�肳�ꂽ�s�̋q�摤�Ј��ԍ�����Socia�����Ԃ܂łɎw�肵�ĕԂ�
    Set GetDataRowRange = Range(Cells(Row, CtsEmployeeNumColumn), Cells(Row, SocWorkingHoursColumn))

End Function


'���Ԃ������`���ɕς�k�����鏈��
'
'���� Data �ϊ����鎞��
'
'�߂�l �����œn���ꂽ���Ԃ���������������
'
Private Function ConvertHoursToDecimal(Data As Variant) As Double
    
    '���Ԃ������ɕϊ����ĕԂ�
    ConvertHoursToDecimal = Int(Data) * 24 + Hour(Data) + Round(Minute(Data) / 60, 2)
    
End Function



