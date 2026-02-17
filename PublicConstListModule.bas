Attribute VB_Name = "PublicConstListModule"
'===============================================================================
' ���W���[����: PublicConstListModule
' ����: �V�X�e���S�̂Ŏg�p����萔�E�񋓌^�̒�`
' �C������:
'   - �X�y���~�X�C���iPulic��Public, LASTANAME��LASTNAME, ALLOCATESCOTE��ALLOCATESCORE�j
'   - Integer��Long�Ή�
'   - �V�X�e������l�̒萔�ǉ�
'===============================================================================
Option Explicit

'===============================================================================
' �V�X�e������l
'===============================================================================
Public Const MAX_CHILDREN As Long = 40          ' ���������
Public Const MAX_TESTS As Long = 1000           ' �e�X�g�����
Public Const MAX_PERSPECTIVES As Long = 5       ' �]���ϓ_�����

'===============================================================================
' �V�[�g�ی�p�X���[�h�i�둀��h�~�p�A�铽�ړI�ł͂Ȃ��j
'===============================================================================
Public Const SHEET_PROTECT_PASSWORD As String = "KenichiroY"

'===============================================================================
' �F�萔
'===============================================================================
Public Const COLOR_NORMAL As Long = 15466475    'RGB(235, 255, 235)
Public Const COLOR_BLUE As Long = 16770740      'RGB(180, 230, 255)
Public Const COLOR_LIGHTBLUE As Long = 16777160 'RGB(200, 255, 255)
Public Const COLOR_GREEN As Long = 13172680     'RGB(200, 255, 200)
Public Const COLOR_ERROR As Long = 255          'RGB(255, 0, 0) - �G���[�\���p
Public Const COLOR_RETEST_HEADER As Long = 7882751   'RGB(255, 200, 120) - �ǎ����w�b�_�[�p�I�����W
Public Const COLOR_RETEST_CELL As Long = 11854079    'RGB(255, 230, 180) - �ǎ������_�Z���p���I�����W

'===============================================================================
' ���_�ϊ�����
'===============================================================================
Public Enum eConversionType
    convNone = 0        ' �ϊ��Ȃ�
    convSqrt = 1        ' ������
    convLog2 = 2        ' ��2�̑ΐ�
End Enum

'===============================================================================
' ����V�[�g�萔
'===============================================================================
Public Const RNG_NAMELIST_CHILDCOUNT As String = "F8"
Public Const NAMELIST_HEADER_ROW As Long = 10
Public Const NAMELIST_DATA_START_ROW As Long = 11
Public Const NAMELIST_COL_END_DATE As Long = 6   ' F��F�ݐЏI����

'===============================================================================
' �e�X�g���̓V�[�g�萔
'===============================================================================
Public Const RNG_INPUT_SUBJECT As String = "D4"
Public Const RNG_INPUT_CATEGORY As String = "F4"
Public Const RNG_INPUT_DATE As String = "J4"
Public Const RNG_INPUT_TEST_NAME As String = "D6"
Public Const RNG_INPUT_TEST_REMARK As String = "J8"

Public Enum eRowInput
    rowPerspective = 8
    rowDetail = 10
    rowAllocateScore = 12
    rowClippingSup = 14
    rowClippingInf = 16
    rowConvScore = 18
    rowAdjScoreSup = 20
    rowAdjScoreInf = 22
    rowAdjAllocateScore = 24
    rowWeight = 26
    rowChildStart = 31
End Enum

Public Enum eColInput
    colDataStart = 4
    colDataEnd = 8          ' �ő�5��i4-8�j
End Enum

'===============================================================================
' �f�[�^�V�[�g�萔
'===============================================================================
Public Enum eRowData
    rowKey = 4
    rowTestDate = 5
    rowSubject = 6
    rowCategory = 7
    rowTestName = 8
    rowPerspective = 9
    rowDetail = 10
    rowAllocationScore = 11
    rowClippingSup = 12
    rowClippingInf = 13
    rowConvScore = 14
    rowAdjScoreSup = 15
    rowAdjScoreInf = 16
    rowAdjAllocateScore = 17
    rowWeight = 18
    rowAverage = 19
    rowMedian = 20
    rowStdDev = 21
    rowCV = 22
    rowChildStart = 23
End Enum

Public Enum eColData
    colCode = 1
    colLastName = 2
    colFirstName = 3
    colDataStart = 4
End Enum

'===============================================================================
' ���j���[�V�[�g�萔
'===============================================================================
Public Enum eRowMenu
    rowStart = 11
End Enum

Public Enum eColMenu
    colCode = 2
    colLastName = 3
    colFirstName = 4
    colSubject = 5
    colPerspective = 6
    colTestName = 7
    colDetail = 8
    colScore = 9
    colAllocateScore = 10
    colToRow = 11
    colToCol = 12
End Enum

'===============================================================================
' Subject�V�[�g�萔
'===============================================================================
Public Const RNG_SUBJECT_SUBJECT As String = "B2"
Public Const RNG_SUBJECT_ISADJUST As String = "B4"
Public Const RNG_SUBJECT_ADJSCORE_DISP As String = "B5"
Public Const RNG_SUBJECT_STATS_DISP As String = "B6"        ' ���v�s�\�����
Public Const RNG_SUBJECT_WEIGHT_NORMALIZED As String = "B7"  ' �d�ݐ��K�����

' �d�ݐ��K���̊�z�_
Public Const NORMALIZE_BASE_SCORE As Long = 100

Public Enum eColShiftSubject
    colNoWeightSum = 2          ' �d�݂Ȃ����v
    colNoWeightAllocate = 3     ' �d�݂Ȃ��z�_
    colIncludeWeightSum = 4     ' �d�ݕt�����v
    colIncludeWeightAllocate = 5 ' �d�ݕt���z�_
    colNoWeightRatio = 7        ' �d�݂Ȃ��B����
    colIncludeWeightRatio = 8   ' ���d�B����
    colABCBorder = 10           ' ABC臒l
End Enum

Public Enum eRowSubject
    rowKey = 4
    rowTestDate = 5
    rowSubject = 6
    rowCategory = 7
    rowTestName = 8
    rowPerspective = 9
    rowDetail = 10
    rowAllocationScore = 11
    rowClippingSup = 12
    rowClippingInf = 13
    rowConvScore = 14
    rowAdjScoreSup = 15
    rowAdjScoreInf = 16
    rowAdjAllocateScore = 17
    rowWeight = 18
    rowAverage = 19
    rowMedian = 20
    rowStdDev = 21
    rowCV = 22
    rowChildStart = 23
End Enum

'===============================================================================
' Result�V�[�g�萔
'===============================================================================
Public Const RESULT_SUBJECT_ROW As Long = 8         ' ���Ȗ��s
Public Const RESULT_PERSPECTIVE_ROW As Long = 9     ' �ϓ_�s
Public Const RESULT_LABEL_ROW As Long = 10          ' ���x���s�i�B����/ABC�j
Public Const RESULT_DATA_START_ROW As Long = 11     ' �����f�[�^�J�n�s
Public Const RESULT_DATA_START_COL As Long = 4      ' �f�[�^�J�n��iD��j

'===============================================================================
' Setting�V�[�g�萔
'===============================================================================
Public Const SETTING_SUBJECT_START_ROW As Long = 3
Public Const SETTING_SUBJECT_COL As Long = 2
Public Const SETTING_KEY_CHAR_COL As Long = 1
Public Const SETTING_KEY_COUNT_COL As Long = 3
Public Const SETTING_PERSPECTIVE_COL As Long = 4
Public Const SETTING_CATEGORY_COL As Long = 6
Public Const SETTING_AB_THRESHOLD_COL As Long = 8
Public Const SETTING_BC_THRESHOLD_COL As Long = 9

'===============================================================================
' �G���[���b�Z�[�W�萔
'===============================================================================
Public Const ERR_MSG_REQUIRED_FIELD As String = "�K�{���ڂ����͂���Ă��܂���B" & vbCrLf & _
    "���ȁA�J�e�S���A���{���A�e�X�g�������ׂē��͂��Ă��������B"
Public Const ERR_MSG_NO_SCORE As String = "�_����1�����͂���Ă��܂���B" & vbCrLf & _
    "���Ȃ��Ƃ�1�l���̓_������͂��Ă��������B"
Public Const ERR_MSG_MISSING_PERSPECTIVE As String = "�Y���̕]���ϓ_�����͂���Ă��܂���B"
Public Const ERR_MSG_MISSING_ALLOCATE As String = "�Y���̔z�_�����͂���Ă��܂���B"
Public Const ERR_MSG_SCORE_EXCEEDS As String = "���_���z�_�𒴂��Ă��܂��B" & vbCrLf & _
    "�s: {ROW}, ��: {COL}" & vbCrLf & "���_: {SCORE}, �z�_: {ALLOCATE}"
Public Const ERR_MSG_NEGATIVE_SCORE As String = "���_�ɕ��̒l�͓��͂ł��܂���B" & vbCrLf & _
    "�s: {ROW}, ��: {COL}"
Public Const ERR_MSG_ZERO_ALLOCATE As String = "�z�_��0�͐ݒ�ł��܂���B�i�[�����Z�G���[�̌����ɂȂ�܂��j"
Public Const ERR_MSG_MAX_TESTS As String = "�e�X�g��������i{MAX}���j�ɒB���Ă��܂��B" & vbCrLf & _
    "�V�����t�@�C�����쐬���Ă��������B"
Public Const ERR_MSG_UNEXPECTED As String = "�\�����Ȃ��G���[���������܂����B" & vbCrLf & _
    "�G���[�ԍ�: {NUM}" & vbCrLf & "�G���[���e: {DESC}" & vbCrLf & vbCrLf & _
    "���̉�ʂ̃X�N���[���V���b�g���Ǘ��҂ɂ����肭�������B"

Public Const MSG_POSTING_SUCCESS As String = "�e�X�g�f�[�^�̓o�^���������܂����B"
Public Const MSG_CONFIRM_DELETE As String = "�I�������f�[�^���폜���Ă�낵���ł����H" & vbCrLf & _
    "���̑���͎������܂���B"

'===============================================================================
' �ǎ��֘A�萔
'===============================================================================
Public Const RETEST_MARKER As String = "N"              ' �ǎ����}�[�J�[�i�f�[�^�V�[�g���_�Z���ɓ���l�j
Public Const RETEST_FILE_SUFFIX As String = "_�ǎ�"      ' �ǎ��t�@�C�����T�t�B�b�N�X
Public Const RETEST_FILE_EXT As String = ".xlsm"         ' �ǎ��t�@�C���g���q

' �e���v���[�g�V�[�g���i�{�̃t�@�C����VeryHidden�V�[�g�j
Public Const RT_MENU_TEMPLATE_NAME As String = "RT_MENU"       ' MENU�e���v���[�g�V�[�g��
Public Const RT_TEMPLATE_NAME As String = "RT_TEMPLATE"        ' �e�X�g�e���v���[�g�V�[�g��

' �ǎ��ݒ�̍s�ʒu�i�e�X�g���̓V�[�g�j
' �s28: �ǎ��L���i�񂲂�: D28�`H28�A���͋K���� "����" ��I���j
Public Const ROW_INPUT_RETEST As Long = 28               ' �ǎ��L���s
Public Const RETEST_ENABLED_VALUE As String = "����"      ' �ǎ��L���̔���l

' �ǎ��V�[�g�̃Z���ʒu�i�ǎ��t�@�C�����̊e�e�X�g�V�[�g�j
' �e�X�g���A-B��i��}���̉e�����󂯂Ȃ��j
Public Const RNG_RT_PARENT_KEY As String = "B3"          ' �ǎ����L�[
Public Const RNG_RT_SUBJECT As String = "B4"             ' ����
Public Const RNG_RT_TEST_NAME As String = "B5"           ' �e�X�g��
Public Const RNG_RT_PERSPECTIVE As String = "B6"         ' �ϓ_
Public Const RNG_RT_DETAIL As String = "B7"              ' �ڍ�
' �Z�o�ݒ�D-E��i��}��:F��ȍ~�Ȃ̂ŉe�����󂯂Ȃ��j
Public Const RNG_RT_ALLOCATE As String = "E3"            ' �z�_
Public Const RNG_RT_PASS_SCORE As String = "E4"          ' ���i�_�i�󗓉j
Public Const RNG_RT_METHOD As String = "E5"              ' �Z�o���@
Public Const RNG_RT_PARAM As String = "E6"               ' �����䃿�l�i�Z�o���@�������_�̏ꍇ�̂݁j
Public Const RNG_RT_STATUS As String = "E7"              ' ��ԁi�ǎ��� / ���� / ���f�ς݁j

' �ǎ��V�[�g�̃f�[�^�̈�
Public Const RT_HEADER_ROW As Long = 10                  ' �w�b�_�[�s
Public Const RT_DATA_START_ROW As Long = 11              ' �����f�[�^�J�n�s
Public Const RT_COL_CODE As Long = 1                     ' A��F�R�[�h
Public Const RT_COL_LASTNAME As Long = 2                 ' B��F��
Public Const RT_COL_FIRSTNAME As Long = 3                ' C��F��
Public Const RT_COL_ORIGINAL As Long = 4                 ' D��F�{��
Public Const RT_COL_RETEST_START As Long = 5             ' E��`�F�ǎ�1, �ǎ�2, ...
Public Const RT_COL_FINAL_OFFSET As Long = 1             ' �ŏI��͍Ō�̒ǎ���+1�E

' �Z�o���@�̑I����
Public Const RT_METHOD_PASS_SCORE As String = "���i�_"
Public Const RT_METHOD_MAX As String = "�ő�l"
Public Const RT_METHOD_AVERAGE As String = "���ϒl"
Public Const RT_METHOD_MEDIAN As String = "�����l"
Public Const RT_METHOD_INTERPOLATION As String = "�����_"
Public Const RT_METHOD_ORIGINAL_ONLY As String = "�{���̂�"

' �ǎ��t�@�C��MENU�V�[�g�̃Z���ʒu
Public Const RT_MENU_DATA_START_ROW As Long = 4
Public Const RT_MENU_COL_KEY As Long = 1                 ' A��F�L�[
Public Const RT_MENU_COL_SUBJECT As Long = 2             ' B��F����
Public Const RT_MENU_COL_TESTNAME As Long = 3            ' C��F�e�X�g��
Public Const RT_MENU_COL_PERSPECTIVE As Long = 4         ' D��F�ϓ_
Public Const RT_MENU_COL_STATUS As Long = 5              ' E��F���
Public Const RT_MENU_COL_REMAINING As Long = 6           ' F��F�c��l��
Public Const RT_MENU_COL_SHEETNAME As Long = 7           ' G��F�V�[�g��


