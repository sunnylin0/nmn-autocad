Attribute VB_Name = "preferences"
Option Explicit


Type AMusicType
    DEFINE_FontFileName  As String          '�w�]�r�����W�r
    DEFINE_TEXT_SIZE As Double          '�w�]�r�����j�p

    
    A_TEXT_WIDTH   As Double   '�w�]�@�Ӧr���e��  �r�Ϊ� �ʤ���

    LINE_PASE As Double '''�o�O�q���I �n�� �r�Ϊ� �ʤ���
    LINE_JUMP   As Double                   '�C�@�ө�u�M��u������ �r�Ϊ� �ʤ���
    LINE_THICKNESS   As Double                   '�o�O��u���p�� �r�Ϊ� �ʤ���
    LINE_LEN   As Double                   '�o�O��u������ �r�Ϊ� �ʤ���

    DROP_UP   As Double                   '���K�ת���m  �r�Ϊ� �ʤ���
    DROP_DOWN   As Double                   '�C�K�ת���m  �r�Ϊ� �ʤ���
    DROP_CIRCLE_RADIUS_SIZE   As Double                   '�ꪺ ���| SIZE �r�Ϊ� �ʤ���
    DROP_INTERVAL  As Double                   '�C�K�ת����Z �r�Ϊ� �ʤ���
    DROP_ADD_LINE_INTERVAL  As Double                   '�U�I �h�@���u�I�n�U�h�h��  �r�Ϊ� �ʤ���

    BAR_WITCH As Double    '�p�`�u���ʲ�

    'IsGiFill As Boolean                 '�b���X���ϧΤ����W���Ͷ�R�ϰ�
    IsGiFill As Boolean                 '�O�_��R�ϰ�
    IsDebugDialog As Boolean                 '�O�_�s�X DebugDialog

    IsDrawCursory As Boolean                 '�O�_ ø�۩w���I
    mHowData As Integer        '���h�֦���
    
    iTONE As Integer         ' * ��
    iFinge As Integer    '�o�O���k��     _+)(*&
    iScale  As Integer     '�o�O�L�C����   .,:
    iNote  As Integer      '�o���D��       1234567.|l
    iTempo  As Integer    '�o�O��l       -=368acefz
    iTowFinge  As Integer     '�o�O���k��ĤG��  _+)(*&
    iSlur  As Integer         ' �s���Ŧ�    (3456)
    iAdd As Integer           '�o�O�X���� [2467]
    
    L_BUF_MD  As Integer   '�o�O��Buf_md�n�h�֦�
    L_WHAT  As Integer    '�o�O���k�ݭn��b���@��
    
    
    
    f�@�� As Integer
    f�G�� As Integer
    f�T�� As Integer
    f�|�� As Integer
    f�ũ� As Integer
    f�|�� As Integer
    
    '2015/1/10 �]�w�϶�����ơA�n�I����m���@��
    
    bkMidHight As Double    '���󤤶��U�I������
    bkMidUpHight As Double    '���󤤶��W�I������
    
    bkInster_Line_H As Double    '������I���u������
    barInsterNumberSize As Double   '�p�`���r�j�p
    
    
    wBar As Double          '�p�`�u���e��
    wMete As Double         '�縹���e��
    wNote As Double         '���Ū��e��
    wOther As Double
    wNote2note As Double    '�C���C�窺���Z

    extraScale As Double    '�e�󤸯� ���
End Type

Public amt As AMusicType
Public rTime As runTime

'MusicTextBlock  �e�϶��θ��

Public Type MusicFing
    sTONE As String        ' �ɭ��O�� ��
    sFinge As String   '�o�O���k��     _+)(*&
    sScale As String   '�o�O���C����   .,:
    sNote As String   '�o���D��       1234567.|l
    sTempo As String   '�o�O��l       -=368acefz
    sTowFinge As String   '�o�O���k��ĤG��  _+)(*&
    sSlur As String
    
    iScale As Integer     '�o�O���C����   .,:
    iNote As Integer     '�o���D��       1234567.|l
    iTempo As Integer    '�o�O��l       -=368acefz
    mfPosition As point    '�o�O�ݳo�Ĥ@���k�r�W�U���h��
    mfTowPosition As point '�o�O�ݳo�ĤG���k�r�W�U���h��
End Type


'�������n�I��m
Public Type GripPoints
    gptMid As point       '�b������
    gptLeftDown As point     '�b���U
    gptMidUp As point     '�b���W
    gptLeft As point      '�b����
    gptRight As point      '�b�k��
    atPt As point           '�ƪ`�I
End Type

    
Public Type Glode
    check1 As Boolean
    fontName As String
    FONTSIZE As Double
    Many As Integer     '�X�n���]�w
    bar As Integer
    mete As Integer     '3/4 �縹�����l --> 3
    mete2 As Integer    '3/4 �縹������ --> 4
    
    
    barsperstaff  As Integer '�]�w�C��X�p�`
    durationIndex As Double ''�ְO �{�b�w�gŪ���C�p�`�����Ū����`�X
    currLine As Integer     '�{�b�O�ĴX��
    currMeasure As Integer     '�{�b�O �o�檺�ĴX�p�`
    
    pagewidth As Double
    LeftSpace As Double
    RightSpace As Double
    BarToNoteSpace As Double     '�p�`�u�쭵�Ū��ť�
    TrackToTrack As Double
    LineToLine As Double
    MIN_X As Double
    Beat_MIN_X As Double
    IsBarAlign As Boolean       '�p�`�O�_���
    
    lastRightPoint As New point '�p���̫�@�Ӥ����� �k���I
End Type


    ''''''''''''''''''''''''''''''''''
Sub AMT_LOAD()
    Dim DEFINE_TEXT_SIZE  As Double
    amt.DEFINE_FontFileName = "�ө���"              '�w�]�r�����W�r
    amt.DEFINE_TEXT_SIZE = 1000              '�w�]�r�����j�p
    DEFINE_TEXT_SIZE = 1000
    
    amt.A_TEXT_WIDTH = 726# / DEFINE_TEXT_SIZE               '�w�]�@�Ӧr���e��  �r�Ϊ� �ʤ���
    
    amt.LINE_PASE = 592.2 / DEFINE_TEXT_SIZE            '�o�O�q���I �n�� �r�Ϊ� �ʤ���
    amt.LINE_JUMP = 155 / DEFINE_TEXT_SIZE            '�C�@�ө�u�M��u������ �r�Ϊ� �ʤ���
    amt.LINE_THICKNESS = 65 / DEFINE_TEXT_SIZE            '�o�O��u���p�� �r�Ϊ� �ʤ���
    amt.LINE_LEN = 723.5 / DEFINE_TEXT_SIZE            '�o�O��u������ �r�Ϊ� �ʤ���
    
    amt.DROP_UP = 1950# / DEFINE_TEXT_SIZE             '���K�ת���m  �r�Ϊ� �ʤ���
    amt.DROP_DOWN = 500# / DEFINE_TEXT_SIZE             '�C�K�ת���m  �r�Ϊ� �ʤ���
    amt.DROP_CIRCLE_RADIUS_SIZE = 200# / DEFINE_TEXT_SIZE / 2           '�ꪺ ���| SIZE �r�Ϊ� �ʤ���
    amt.DROP_INTERVAL = 290# / DEFINE_TEXT_SIZE             '�C�K�ת����Z �r�Ϊ� �ʤ���
    amt.DROP_ADD_LINE_INTERVAL = 170# / DEFINE_TEXT_SIZE               '�U�I �h�@���u�I�n�U�h�h��  �r�Ϊ� �ʤ���
    
    amt.BAR_WITCH = 340 / DEFINE_TEXT_SIZE                '�p�`�u���ʲ�

    '''bool  IsGiFill    =true ''�b���X���ϧΤ����W���Ͷ�R�ϰ�
    amt.IsGiFill = False '�O�_��R�ϰ�
    amt.IsDebugDialog = False      '�O�_�s�X DebugDialog

    amt.IsDrawCursory = True    '�O�_ ø�۩w���I

    amt.mHowData = 9    '��9�檺���

    amt.iTONE = 1        ' * ��
    amt.iFinge = 2    '�o�O���k��     _+)(*&
    amt.iScale = 3    '�o�O���C����   .,:
    amt.iNote = 4     '�o���D��       1234567.|l
    amt.iTempo = 5    '�o�O��l       -=368acefz
    amt.iTowFinge = 6    '�o�O���k��ĤG��  _+)(*&
    amt.iSlur = 7        ' �s���Ŧ�    (3456)
    amt.iAdd = 8        ' �X���Ŧ�    [3456]
    
    amt.L_BUF_MD = 9  '�o�O��Buf_md�n�h�֦�
    amt.L_WHAT = 8    '�o�O���k�ݭn��b���@��

    '�G�J��
    amt.f�@�� = 1
    amt.f�G�� = 2
    amt.f�T�� = 4
    amt.f�|�� = 8
    amt.f�ũ� = 16
    amt.f�|�� = 32
    
    
    
    '2015/1/10 �]�w�϶�����ơA�n�I����m���@��
    amt.bkMidHight = 690# / DEFINE_TEXT_SIZE      '���󤤶��U�I������
    amt.bkMidUpHight = 1690# / DEFINE_TEXT_SIZE     '���󤤶��W�I������
    amt.bkInster_Line_H = 97.8 / DEFINE_TEXT_SIZE    '������I���u������
    amt.barInsterNumberSize = 0.6
    
    
    amt.wBar = 1    '�p�`�u���e��
    amt.wMete = 1      '�縹���e��
    amt.wNote = 0.75        '���Ū��e��
    amt.wOther = 1
    amt.wNote2note = 0.5 '�C���C�窺���Z
    
    amt.extraScale = 0.6 '
    
End Sub
    
    
