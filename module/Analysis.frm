VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Analysis
   Caption         =   "比較分析"
   ClientHeight    =   8655.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11475
   OleObjectBlob   =   "Analysis.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' フォーム名: Analysis（比較分析）
' 機能: テスト比較分析フォーム
'       教科を選択し、テストを2グループ（各最大3つ）に振り分けて
'       IndividualAnalysisシートにデータを転記する
' コントロール:
'   Com_Subject  - 教科選択コンボボックス
'   UnitList     - テスト一覧リストボックス（3列: テスト名, 観点, 列番号）
'   Det1         - グループ1リストボックス（3列: テスト名, 観点, 列番号）
'   Det2         - グループ2リストボックス（3列: テスト名, 観点, 列番号）
'   btn_det1     - グループ1に追加ボタン
'   btn_det2     - グループ2に追加ボタン
'   Deldet1      - グループ1から削除ボタン
'   Deldet2      - グループ2から削除ボタン
'   CommandButton1 - 実行ボタン
'===============================================================================
Option Explicit

' IndividualAnalysisシートのレイアウト定数
Private Const IA_HEADER_START_ROW As Long = 9     ' ヘッダー開始行
Private Const IA_KEY_ROW As Long = 9              ' キー行
Private Const IA_DATE_ROW As Long = 10            ' 日付行
Private Const IA_SUBJECT_ROW As Long = 11         ' 教科行
Private Const IA_TESTNAME_ROW As Long = 12        ' テスト名行
Private Const IA_PERSPECTIVE_ROW As Long = 13     ' 観点行
Private Const IA_ALLOCATE_ROW As Long = 14        ' 配点行
Private Const IA_CHILD_START_ROW As Long = 15     ' 児童データ開始行

Private Const IA_DET1_COL_START As Long = 3       ' グループ1開始列（C列）
Private Const IA_DET2_COL_START As Long = 8       ' グループ2開始列（H列）

Private Const MAX_SELECTION As Long = 3           ' 各グループの最大選択数

'===============================================================================
' フォーム初期化
' 処理: Settingシートから教科一覧を読み込み、コンボボックスに設定
'===============================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim subjectCount As Long

    Com_Subject.Style = fmStyleDropDownList

    ' Settingシートの教科数を取得（B3から下方向にカウント）
    subjectCount = WorksheetFunction.CountA( _
        sh_setting.Range(sh_setting.Cells(SETTING_SUBJECT_START_ROW, SETTING_SUBJECT_COL), _
                         sh_setting.Cells(SETTING_SUBJECT_START_ROW + 20, SETTING_SUBJECT_COL)))

    For i = 1 To subjectCount
        Com_Subject.AddItem sh_setting.Cells(SETTING_SUBJECT_START_ROW + i - 1, SETTING_SUBJECT_COL).value
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "フォームの初期化中にエラーが発生しました。" & vbCrLf & _
           "エラー: " & Err.Description, vbExclamation, "比較分析"
End Sub

'===============================================================================
' 教科選択変更イベント
' 処理: 選択された教科に一致するテストをデータシートから検索し、
'       UnitListに表示（テスト名, 観点, 列番号）
'===============================================================================
Private Sub Com_Subject_Change()
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim lastCol As Long

    UnitList.Clear

    With Sh_data
        lastCol = .Cells(eRowData.rowKey, Columns.Count).End(xlToLeft).Column

        For i = eColData.colDataStart To lastCol
            If Com_Subject.value = .Cells(eRowData.rowSubject, i).value Then
                UnitList.AddItem .Cells(eRowData.rowTestName, i).value
                UnitList.List(UnitList.ListCount - 1, 1) = .Cells(eRowData.rowPerspective, i).value
                UnitList.List(UnitList.ListCount - 1, 2) = i
            End If
        Next i
    End With

    Exit Sub

ErrorHandler:
    MsgBox "テスト一覧の取得中にエラーが発生しました。" & vbCrLf & _
           "エラー: " & Err.Description, vbExclamation, "比較分析"
End Sub

'===============================================================================
' グループ1追加ボタン
'===============================================================================
Private Sub btn_det1_Click()
    If UnitList.ListIndex = -1 Then Exit Sub
    If Det1.ListCount >= MAX_SELECTION Then
        MsgBox "選択できる数は" & MAX_SELECTION & "までです。", vbInformation, "比較分析"
        Exit Sub
    End If
    Det1.AddItem UnitList.List(UnitList.ListIndex, 0)
    Det1.List(Det1.ListCount - 1, 1) = UnitList.List(UnitList.ListIndex, 1)
    Det1.List(Det1.ListCount - 1, 2) = UnitList.List(UnitList.ListIndex, 2)
End Sub

'===============================================================================
' グループ2追加ボタン
'===============================================================================
Private Sub btn_det2_Click()
    If UnitList.ListIndex = -1 Then Exit Sub
    If Det2.ListCount >= MAX_SELECTION Then
        MsgBox "選択できる数は" & MAX_SELECTION & "までです。", vbInformation, "比較分析"
        Exit Sub
    End If
    Det2.AddItem UnitList.List(UnitList.ListIndex, 0)
    Det2.List(Det2.ListCount - 1, 1) = UnitList.List(UnitList.ListIndex, 1)
    Det2.List(Det2.ListCount - 1, 2) = UnitList.List(UnitList.ListIndex, 2)
End Sub

'===============================================================================
' グループ1削除ボタン
'===============================================================================
Private Sub Deldet1_Click()
    If Det1.ListIndex >= 0 Then Det1.RemoveItem Det1.ListIndex
End Sub

'===============================================================================
' グループ2削除ボタン
'===============================================================================
Private Sub Deldet2_Click()
    If Det2.ListIndex >= 0 Then Det2.RemoveItem Det2.ListIndex
End Sub

'===============================================================================
' 実行ボタン
' 処理: 選択されたテストデータをIndividualAnalysisシートに転記
'
' IndividualAnalysisシートのレイアウト:
'   行9:  キー          | Det1: C-E列 | Det2: H-J列
'   行10: 日付          |             |
'   行11: 教科          |             |
'   行12: テスト名      |             |
'   行13: 観点          |             |
'   行14: 配点          |             |
'   行15～: 児童得点     |             |
'===============================================================================
Private Sub CommandButton1_Click()
    On Error GoTo ErrorHandler

    ' 児童数を名簿シートから取得
    Dim childCount As Long
    childCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).value
    If childCount <= 0 Then
        MsgBox "名簿に児童データがありません。", vbExclamation, "比較分析"
        Exit Sub
    End If

    ' 選択チェック
    If Det1.ListCount = 0 And Det2.ListCount = 0 Then
        MsgBox "テストを1つ以上選択してください。", vbExclamation, "比較分析"
        Exit Sub
    End If

    Dim i As Long
    Dim j As Long
    Dim dataCol As Long
    Dim destCol As Long
    Dim lastChildRow As Long
    lastChildRow = IA_CHILD_START_ROW + childCount - 1

    With sh_individual
        ' 前回データをクリア
        ' グループ1: C9～E列（最終児童行まで）
        .Range(.Cells(IA_HEADER_START_ROW, IA_DET1_COL_START), _
               .Cells(lastChildRow, IA_DET1_COL_START + MAX_SELECTION - 1)).ClearContents
        ' グループ2: H9～J列（最終児童行まで）
        .Range(.Cells(IA_HEADER_START_ROW, IA_DET2_COL_START), _
               .Cells(lastChildRow, IA_DET2_COL_START + MAX_SELECTION - 1)).ClearContents

        ' === グループ1のデータ転記 ===
        For i = 1 To Det1.ListCount
            dataCol = CLng(Det1.List(i - 1, 2))
            destCol = IA_DET1_COL_START + i - 1

            ' ヘッダー情報
            .Cells(IA_KEY_ROW, destCol).value = Sh_data.Cells(eRowData.rowKey, dataCol).value
            .Cells(IA_DATE_ROW, destCol).value = Sh_data.Cells(eRowData.rowTestDate, dataCol).value
            .Cells(IA_SUBJECT_ROW, destCol).value = Sh_data.Cells(eRowData.rowSubject, dataCol).value
            .Cells(IA_TESTNAME_ROW, destCol).value = Sh_data.Cells(eRowData.rowTestName, dataCol).value
            .Cells(IA_PERSPECTIVE_ROW, destCol).value = Sh_data.Cells(eRowData.rowPerspective, dataCol).value
            .Cells(IA_ALLOCATE_ROW, destCol).value = Sh_data.Cells(eRowData.rowAllocationScore, dataCol).value

            ' 児童得点
            For j = 1 To childCount
                .Cells(IA_CHILD_START_ROW + j - 1, destCol).value = _
                    Sh_data.Cells(eRowData.rowChildStart + j - 1, dataCol).value
            Next j
        Next i

        ' === グループ2のデータ転記 ===
        For i = 1 To Det2.ListCount
            dataCol = CLng(Det2.List(i - 1, 2))
            destCol = IA_DET2_COL_START + i - 1

            ' ヘッダー情報
            .Cells(IA_KEY_ROW, destCol).value = Sh_data.Cells(eRowData.rowKey, dataCol).value
            .Cells(IA_DATE_ROW, destCol).value = Sh_data.Cells(eRowData.rowTestDate, dataCol).value
            .Cells(IA_SUBJECT_ROW, destCol).value = Sh_data.Cells(eRowData.rowSubject, dataCol).value
            .Cells(IA_TESTNAME_ROW, destCol).value = Sh_data.Cells(eRowData.rowTestName, dataCol).value
            .Cells(IA_PERSPECTIVE_ROW, destCol).value = Sh_data.Cells(eRowData.rowPerspective, dataCol).value
            .Cells(IA_ALLOCATE_ROW, destCol).value = Sh_data.Cells(eRowData.rowAllocationScore, dataCol).value

            ' 児童得点
            For j = 1 To childCount
                .Cells(IA_CHILD_START_ROW + j - 1, destCol).value = _
                    Sh_data.Cells(eRowData.rowChildStart + j - 1, dataCol).value
            Next j
        Next i

        ' 分析シートをアクティブに
        .Activate
    End With

    Exit Sub

ErrorHandler:
    MsgBox "データ転記中にエラーが発生しました。" & vbCrLf & _
           "エラー: " & Err.Description, vbExclamation, "比較分析"
End Sub
