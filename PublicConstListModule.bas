'===============================================================================
' モジュール名: PublicConstListModule
' 説明: システム全体で使用する定数・列挙型の定義
' 修正履歴:
'   - スペルミス修正（Pulic→Public, LASTANAME→LASTNAME, ALLOCATESCOTE→ALLOCATESCORE）
'   - Integer→Long対応
'   - システム上限値の定数追加
'===============================================================================
Option Explicit

'===============================================================================
' システム上限値
'===============================================================================
Public Const MAX_CHILDREN As Long = 40          ' 児童数上限
Public Const MAX_TESTS As Long = 1000           ' テスト数上限
Public Const MAX_PERSPECTIVES As Long = 5       ' 評価観点数上限

'===============================================================================
' 色定数
'===============================================================================
Public Const COLOR_NORMAL As Long = 15466475    'RGB(235, 255, 235)
Public Const COLOR_BLUE As Long = 16770740      'RGB(180, 230, 255)
Public Const COLOR_LIGHTBLUE As Long = 16777160 'RGB(200, 255, 255)
Public Const COLOR_GREEN As Long = 13172680     'RGB(200, 255, 200)
Public Const COLOR_ERROR As Long = 255          'RGB(255, 0, 0) - エラー表示用
Public Const COLOR_RETEST_HEADER As Long = 7882751   'RGB(255, 200, 120) - 追試中ヘッダー用オレンジ
Public Const COLOR_RETEST_CELL As Long = 11854079    'RGB(255, 230, 180) - 追試中得点セル用薄オレンジ

'===============================================================================
' 得点変換方式
'===============================================================================
Public Enum eConversionType
    convNone = 0        ' 変換なし
    convSqrt = 1        ' 平方根
    convLog2 = 2        ' 底2の対数
End Enum

'===============================================================================
' 名簿シート定数
'===============================================================================
Public Const RNG_NAMELIST_CHILDCOUNT As String = "E8"
Public Const NAMELIST_HEADER_ROW As Long = 10
Public Const NAMELIST_DATA_START_ROW As Long = 11

'===============================================================================
' テスト入力シート定数
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
    colDataEnd = 8          ' 最大5列（4-8）
End Enum

'===============================================================================
' データシート定数
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
    colLastName = 2         ' 修正: LASTANAME → colLastName
    colFirstName = 3
    colDataStart = 4
End Enum

'===============================================================================
' メニューシート定数
'===============================================================================
Public Enum eRowMenu
    rowStart = 11
End Enum

Public Enum eColMenu
    colCode = 2
    colLastName = 3         ' 修正: LASTNAME
    colFirstName = 4
    colSubject = 5
    colPerspective = 6
    colTestName = 7
    colDetail = 8
    colScore = 9
    colAllocateScore = 10   ' 修正: ALLOCATESCOTE → colAllocateScore
    colToRow = 11
    colToCol = 12
End Enum

'===============================================================================
' Subjectシート定数
'===============================================================================
Public Const RNG_SUBJECT_SUBJECT As String = "B2"
Public Const RNG_SUBJECT_ISADJUST As String = "B4"
Public Const RNG_SUBJECT_ADJSCORE_DISP As String = "B5"
Public Const RNG_SUBJECT_WEIGHT_NORMALIZED As String = "B6"  ' 重み正規化状態

' 重み正規化の基準配点
Public Const NORMALIZE_BASE_SCORE As Long = 100

Public Enum eColShiftSubject
    colNoWeightSum = 2          ' 重み無し合計
    colNoWeightAllocate = 3     ' 重み無し配点
    colIncludeWeightSum = 4     ' 重みあり合計
    colIncludeWeightAllocate = 5 ' 重みあり配点
    colNoWeightRatio = 7        ' 重み無し達成率
    colIncludeWeightRatio = 8   ' 加重達成率
    colABCBorder = 10           ' ABC閾値
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
' Resultシート定数
'===============================================================================
Public Const RESULT_SUBJECT_ROW As Long = 8         ' 教科名行
Public Const RESULT_PERSPECTIVE_ROW As Long = 9     ' 観点行
Public Const RESULT_LABEL_ROW As Long = 10          ' ラベル行（達成率/ABC）
Public Const RESULT_DATA_START_ROW As Long = 11     ' 児童データ開始行
Public Const RESULT_DATA_START_COL As Long = 4      ' データ開始列（D列）

'===============================================================================
' Settingシート定数
'===============================================================================
Public Const SETTING_SUBJECT_START_ROW As Long = 3
Public Const SETTING_SUBJECT_COL As Long = 2
Public Const SETTING_KEY_CHAR_COL As Long = 1
Public Const SETTING_KEY_COUNT_COL As Long = 3
Public Const SETTING_PERSPECTIVE_COL As Long = 4
Public Const SETTING_CATEGORY_COL As Long = 7
Public Const SETTING_AB_THRESHOLD_COL As Long = 8
Public Const SETTING_BC_THRESHOLD_COL As Long = 9

'===============================================================================
' エラーメッセージ定数
'===============================================================================
Public Const ERR_MSG_REQUIRED_FIELD As String = "必須項目が入力されていません。" & vbCrLf & _
    "教科、カテゴリ、実施日、テスト名をすべて入力してください。"
Public Const ERR_MSG_NO_SCORE As String = "点数が1件も入力されていません。" & vbCrLf & _
    "少なくとも1人分の点数を入力してください。"
Public Const ERR_MSG_MISSING_PERSPECTIVE As String = "列目の評価観点が入力されていません。"
Public Const ERR_MSG_MISSING_ALLOCATE As String = "列目の配点が入力されていません。"
Public Const ERR_MSG_SCORE_EXCEEDS As String = "得点が配点を超えています。" & vbCrLf & _
    "行: {ROW}, 列: {COL}" & vbCrLf & "得点: {SCORE}, 配点: {ALLOCATE}"
Public Const ERR_MSG_NEGATIVE_SCORE As String = "得点に負の値は入力できません。" & vbCrLf & _
    "行: {ROW}, 列: {COL}"
Public Const ERR_MSG_ZERO_ALLOCATE As String = "配点に0は設定できません。（ゼロ除算エラーの原因になります）"
Public Const ERR_MSG_MAX_TESTS As String = "テスト数が上限（{MAX}件）に達しています。" & vbCrLf & _
    "新しいファイルを作成してください。"
Public Const ERR_MSG_UNEXPECTED As String = "予期しないエラーが発生しました。" & vbCrLf & _
    "エラー番号: {NUM}" & vbCrLf & "エラー内容: {DESC}" & vbCrLf & vbCrLf & _
    "この画面のスクリーンショットを管理者にお見せください。"

Public Const MSG_POSTING_SUCCESS As String = "テストデータの登録が完了しました。"
Public Const MSG_CONFIRM_DELETE As String = "選択したデータを削除してもよろしいですか？" & vbCrLf & _
    "この操作は取り消せません。"

'===============================================================================
' 追試関連定数
'===============================================================================
Public Const RETEST_MARKER As String = "N"              ' 追試中マーカー（データシート得点セルに入る値）
Public Const RETEST_FILE_SUFFIX As String = "_追試"      ' 追試ファイル名サフィックス
Public Const RETEST_FILE_EXT As String = ".xlsm"         ' 追試ファイル拡張子

' テンプレートシート名（本体ファイルのVeryHiddenシート）
Public Const RT_MENU_TEMPLATE_NAME As String = "RT_MENU"       ' MENUテンプレートシート名
Public Const RT_TEMPLATE_NAME As String = "RT_TEMPLATE"        ' テストテンプレートシート名

' 追試設定の行位置（テスト入力シート）
' 行28: 追試有無（列ごと: D28～H28、入力規則で "あり" を選択）
Public Const ROW_INPUT_RETEST As Long = 28               ' 追試有無行
Public Const RETEST_ENABLED_VALUE As String = "あり"      ' 追試有効の判定値

' 追試シートのセル位置（追試ファイル内の各テストシート）
' テスト情報はA-B列（列挿入の影響を受けない）
Public Const RNG_RT_PARENT_KEY As String = "B3"          ' 追試元キー
Public Const RNG_RT_SUBJECT As String = "B4"             ' 教科
Public Const RNG_RT_TEST_NAME As String = "B5"           ' テスト名
Public Const RNG_RT_PERSPECTIVE As String = "B6"         ' 観点
Public Const RNG_RT_DETAIL As String = "B7"              ' 詳細
' 算出設定はD-E列（列挿入はF列以降なので影響を受けない）
Public Const RNG_RT_ALLOCATE As String = "E3"            ' 配点
Public Const RNG_RT_PASS_SCORE As String = "E4"          ' 合格点（空欄可）
Public Const RNG_RT_METHOD As String = "E5"              ' 算出方法
Public Const RNG_RT_PARAM As String = "E6"               ' 内分比α値（算出方法が内分点の場合のみ）
Public Const RNG_RT_STATUS As String = "E7"              ' 状態（追試中 / 完了 / 反映済み）

' 追試シートのデータ領域
Public Const RT_HEADER_ROW As Long = 10                  ' ヘッダー行
Public Const RT_DATA_START_ROW As Long = 11              ' 児童データ開始行
Public Const RT_COL_CODE As Long = 1                     ' A列：コード
Public Const RT_COL_LASTNAME As Long = 2                 ' B列：姓
Public Const RT_COL_FIRSTNAME As Long = 3                ' C列：名
Public Const RT_COL_ORIGINAL As Long = 4                 ' D列：本試
Public Const RT_COL_RETEST_START As Long = 5             ' E列～：追試1, 追試2, ...
Public Const RT_COL_FINAL_OFFSET As Long = 1             ' 最終列は最後の追試列の1つ右

' 算出方法の選択肢
Public Const RT_METHOD_PASS_SCORE As String = "合格点"
Public Const RT_METHOD_MAX As String = "最大値"
Public Const RT_METHOD_AVERAGE As String = "平均値"
Public Const RT_METHOD_MEDIAN As String = "中央値"
Public Const RT_METHOD_INTERPOLATION As String = "内分点"
Public Const RT_METHOD_ORIGINAL_ONLY As String = "本試のみ"

' 追試ファイルMENUシートのセル位置
Public Const RT_MENU_DATA_START_ROW As Long = 4
Public Const RT_MENU_COL_KEY As Long = 1                 ' A列：キー
Public Const RT_MENU_COL_SUBJECT As Long = 2             ' B列：教科
Public Const RT_MENU_COL_TESTNAME As Long = 3            ' C列：テスト名
Public Const RT_MENU_COL_PERSPECTIVE As Long = 4         ' D列：観点
Public Const RT_MENU_COL_STATUS As Long = 5              ' E列：状態
Public Const RT_MENU_COL_REMAINING As Long = 6           ' F列：残り人数
Public Const RT_MENU_COL_SHEETNAME As Long = 7           ' G列：シート名
