Attribute VB_Name = "InitialSetupModule"
'===============================================================================
' モジュール名: InitialSetupModule
' 用途: 初期セットアップの実行ロジック
'===============================================================================
Option Explicit

'===============================================================================
' セットアップフォームの表示（エントリーポイント）
'===============================================================================
Public Sub ShowSetupForm()
    ' テストデータが既にある場合はブロック
    If Sh_data.Cells(eRowData.rowKey, eColData.colDataStart).value <> "" Then
        Call ErrorHandlerModule.ShowValidationError( _
            "テストデータが既に登録されているため、初期設定ウィザードは使用できません。" & vbCrLf & _
            "設定を変更するには、Settingシートを直接編集してください。" & vbCrLf & _
            "完全に初期化するには、Settingシートの「完全初期化」ボタンを使用してください。")
        Exit Sub
    End If

    frmSetup.Show vbModeless
End Sub

'===============================================================================
' セットアップ実行
'===============================================================================
Public Sub ExecuteSetup(ByRef subjects() As String, ByRef keyChars() As String, _
                        ByVal subjectCount As Long)
    On Error GoTo ErrorHandler

    Call ErrorHandlerModule.BeginProcess

    ' ----- 1. Settingシートの教科列をクリア（観点・閾値・カテゴリは保持） -----
    With sh_setting
        .Range(.Cells(SETTING_SUBJECT_START_ROW, SETTING_KEY_CHAR_COL), _
               .Cells(20, SETTING_KEY_COUNT_COL)).ClearContents
    End With

    ' ----- 2. 教科を書込み -----
    Dim i As Long
    With sh_setting
        For i = 1 To subjectCount
            .Cells(SETTING_SUBJECT_START_ROW + i - 1, SETTING_KEY_CHAR_COL).value = keyChars(i)
            .Cells(SETTING_SUBJECT_START_ROW + i - 1, SETTING_SUBJECT_COL).value = subjects(i)
            .Cells(SETTING_SUBJECT_START_ROW + i - 1, SETTING_KEY_COUNT_COL).value = 0
        Next i
    End With

    ' ----- 3. 観点を書込み（既存値がなければデフォルト設定） -----
    If Trim(sh_setting.Cells(SETTING_SUBJECT_START_ROW, SETTING_PERSPECTIVE_COL).value & "") = "" Then
        With sh_setting
            .Cells(SETTING_SUBJECT_START_ROW, SETTING_PERSPECTIVE_COL).value = "知識・技能"
            .Cells(SETTING_SUBJECT_START_ROW + 1, SETTING_PERSPECTIVE_COL).value = "思考・判断・表現"
            .Cells(SETTING_SUBJECT_START_ROW + 2, SETTING_PERSPECTIVE_COL).value = "主体的に学習に取り組む態度"
        End With
    End If

    ' ----- 4. ABC閾値を書込み（既存値がなければデフォルト設定） -----
    If Trim(sh_setting.Cells(SETTING_SUBJECT_START_ROW, SETTING_AB_THRESHOLD_COL).value & "") = "" Then
        With sh_setting
            .Cells(SETTING_SUBJECT_START_ROW, SETTING_AB_THRESHOLD_COL).value = 80
            .Cells(SETTING_SUBJECT_START_ROW, SETTING_BC_THRESHOLD_COL).value = 50

            .Cells(SETTING_SUBJECT_START_ROW + 1, SETTING_AB_THRESHOLD_COL).value = 75
            .Cells(SETTING_SUBJECT_START_ROW + 1, SETTING_BC_THRESHOLD_COL).value = 45

            .Cells(SETTING_SUBJECT_START_ROW + 2, SETTING_AB_THRESHOLD_COL).value = 70
            .Cells(SETTING_SUBJECT_START_ROW + 2, SETTING_BC_THRESHOLD_COL).value = 40
        End With
    End If

    ' ----- 5. カテゴリを書込み（既存値がなければデフォルト設定） -----
    If Trim(sh_setting.Cells(SETTING_SUBJECT_START_ROW, SETTING_CATEGORY_COL).value & "") = "" Then
        With sh_setting
            .Cells(SETTING_SUBJECT_START_ROW, SETTING_CATEGORY_COL).value = "単元テスト"
            .Cells(SETTING_SUBJECT_START_ROW + 1, SETTING_CATEGORY_COL).value = "まとめテスト"
            .Cells(SETTING_SUBJECT_START_ROW + 2, SETTING_CATEGORY_COL).value = "スキルテスト"
        End With
    End If

    ' ----- 6. Resultシートをクリアして列見出し再生成 -----
    With sh_result
        Dim lastCol As Long
        lastCol = .Cells(RESULT_SUBJECT_ROW, Columns.count).End(xlToLeft).Column
        If lastCol >= RESULT_DATA_START_COL Then
            Dim clearLastRow As Long
            clearLastRow = RESULT_DATA_START_ROW + MAX_CHILDREN + 5
            .Range(.Cells(RESULT_SUBJECT_ROW, RESULT_DATA_START_COL), _
                   .Cells(clearLastRow, lastCol)).Clear
        End If
    End With
    Call ResultModule.GenerateResultHeaders

    ' ----- 7. Resultシートのデザイン適用 -----
    Call FormatResultModule.FormatResultSheet

    ' ----- 8. Subject観点チェックボックス初期化 -----
    Call InitializeSubjectCheckboxesFromSetup

    ' ----- 9. sh_inputのドロップダウン設定 -----
    Call SetupInputValidation

    ' ----- 10. Settingシートの教科列をロック -----
    Call SetupSettingSheetProtection

    ' ----- 完了メッセージ（観点はSettingシートの実際の値を表示） -----
    Dim msgPerspectives As String
    Dim pIdx As Long
    Dim pName As String
    Dim pCount As Long
    pCount = 0
    For pIdx = SETTING_SUBJECT_START_ROW To SETTING_SUBJECT_START_ROW + MAX_PERSPECTIVES - 1
        pName = Trim(sh_setting.Cells(pIdx, SETTING_PERSPECTIVE_COL).value & "")
        If pName = "" Then Exit For
        If pCount > 0 Then msgPerspectives = msgPerspectives & " / "
        msgPerspectives = msgPerspectives & pName
        pCount = pCount + 1
    Next pIdx

    MsgBox "初期設定が完了しました。" & vbCrLf & vbCrLf & _
           "教科数: " & subjectCount & vbCrLf & _
           "評価観点: " & pCount & "（" & msgPerspectives & "）" & vbCrLf & _
           "ABC閾値: Settingシートで変更可能" & vbCrLf & _
           "カテゴリ: Settingシートで追加可能", _
           vbInformation, "初期設定完了"

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("InitialSetupModule", "ExecuteSetup")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' Subject観点チェックボックスの初期化
'===============================================================================
Private Sub InitializeSubjectCheckboxesFromSetup()
    On Error Resume Next

    Dim i As Long

    With sh_subject
        For i = 1 To MAX_PERSPECTIVES
            If Trim(sh_setting.Cells(i + 2, SETTING_PERSPECTIVE_COL).value & "") = "" Then
                .CheckBoxes("perspective" & i).Visible = False
            Else
                .CheckBoxes("perspective" & i).Caption = _
                    sh_setting.Cells(i + 2, SETTING_PERSPECTIVE_COL).value
                .CheckBoxes("perspective" & i).Visible = True
            End If
        Next i
    End With

    On Error GoTo 0
End Sub

'===============================================================================
' sh_inputにデータ入力規則（ドロップダウン）を設定
'===============================================================================
Private Sub SetupInputValidation()
    On Error Resume Next

    ' シート保護を一時的に解除
    sh_input.Unprotect Password:=SHEET_PROTECT_PASSWORD

    ' 教科のドロップダウン（D4）
    With sh_input.Range(RNG_INPUT_SUBJECT).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=Setting!$B$3:$B$20"
        .InCellDropdown = True
    End With

    ' カテゴリのドロップダウン（F4）
    With sh_input.Range("F4").Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=Setting!$F$3:$F$13"
        .InCellDropdown = True
    End With

    ' シート保護を復帰
    sh_input.Protect Password:=SHEET_PROTECT_PASSWORD, UserInterfaceOnly:=True

    ' Subjectシートの教科ドロップダウン（B2）
    sh_subject.Unprotect Password:=SHEET_PROTECT_PASSWORD
    With sh_subject.Range(RNG_SUBJECT_SUBJECT).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=Setting!$B$3:$B$20"
        .InCellDropdown = True
    End With
    sh_subject.Protect Password:=SHEET_PROTECT_PASSWORD, UserInterfaceOnly:=True

    On Error GoTo 0
End Sub

'===============================================================================
' Settingシートの教科列（A-C列）をセルレベルでロック
' D列以降（観点・カテゴリ・閾値）はユーザーが自由に編集可能
'===============================================================================
Public Sub SetupSettingSheetProtection()
    On Error Resume Next

    With sh_setting
        ' 一旦保護を解除
        .Unprotect Password:=SHEET_PROTECT_PASSWORD

        ' 全セルをロック解除
        .Cells.Locked = False

        ' 教科列（A3:C20）のみロック
        .Range(.Cells(SETTING_SUBJECT_START_ROW, SETTING_KEY_CHAR_COL), _
               .Cells(20, SETTING_KEY_COUNT_COL)).Locked = True

        ' 保護を再設定（UserInterfaceOnlyで VBAからの操作は許可）
        .Protect Password:=SHEET_PROTECT_PASSWORD, _
            DrawingObjects:=True, Contents:=True, Scenarios:=False, _
            UserInterfaceOnly:=True
    End With

    On Error GoTo 0
End Sub
