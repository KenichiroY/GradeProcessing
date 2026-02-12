'===============================================================================
' モジュール名: sh_input (テスト入力シート)
' 説明: テスト入力シートのイベント処理とUI制御
' 修正内容:
'   - 未完成関数（check_INPUT）を削除
'   - エラーハンドリング追加
'===============================================================================
Option Explicit

'===============================================================================
' 調整後配点の行表示制御
' 説明: クリッピング・得点変換の設定に応じて行の表示/非表示を切り替え
'===============================================================================
Public Sub AdjustedAllocateScoreVisibility()
    On Error Resume Next
    
    ' 得点調整オプションが1つでも有効な場合は調整後配点行を表示
    If sh_input.Cells(18, 14).Value = 0 Then
        sh_input.Rows("24:25").Hidden = True
    Else
        sh_input.Rows("24:25").Hidden = False
    End If
    
    On Error GoTo 0
End Sub

'===============================================================================
' クリッピングチェックボックスの変更イベント
'===============================================================================
Public Sub Cb_clipping()
    On Error GoTo ErrorHandler
    
    Dim cb As CheckBox
    Set cb = sh_input.CheckBoxes("Cb_clipping")
    
    If cb.Value = xlOn Then
        sh_input.Rows("14:17").Hidden = False
    Else
        sh_input.Rows("14:17").Hidden = True
    End If
    
    Call AdjustedAllocateScoreVisibility
    Exit Sub
    
ErrorHandler:
    Debug.Print "Cb_clipping Error: " & Err.Description
End Sub

'===============================================================================
' 得点変換チェックボックスの変更イベント
'===============================================================================
Public Sub Cb_convertScore()
    On Error GoTo ErrorHandler
    
    Dim cb As CheckBox
    Set cb = sh_input.CheckBoxes("Cb_convertScore")
    
    If cb.Value = xlOn Then
        sh_input.Rows("18:19").Hidden = False
    Else
        sh_input.Rows("18:19").Hidden = True
    End If
    
    Call AdjustedAllocateScoreVisibility
    Exit Sub
    
ErrorHandler:
    Debug.Print "Cb_convertScore Error: " & Err.Description
End Sub

'===============================================================================
' 得点範囲調整チェックボックスの変更イベント
'===============================================================================
Public Sub Cb_adjustScore()
    On Error GoTo ErrorHandler
    
    Dim cb As CheckBox
    Set cb = sh_input.CheckBoxes("Cb_adjustScore")
    
    If cb.Value = xlOn Then
        sh_input.Rows("20:23").Hidden = False
    Else
        sh_input.Rows("20:23").Hidden = True
    End If
    
    Call AdjustedAllocateScoreVisibility
    Exit Sub
    
ErrorHandler:
    Debug.Print "Cb_adjustScore Error: " & Err.Description
End Sub

'===============================================================================
' 入力フォームのクリアボタン
'===============================================================================
Public Sub ClearInputForm()
    On Error GoTo ErrorHandler
    
    If ErrorHandlerModule.ShowConfirmation( _
        "入力中のデータをすべてクリアしますか？", "クリア確認") Then
        Call PostingModule.ResetInputForm
        Call ErrorHandlerModule.ShowInfo("入力フォームをクリアしました。")
    End If
    
    Exit Sub
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("sh_input", "ClearInputForm")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 在籍/全員切替ボタン
' 説明: ボタンのキャプションが「在籍」のとき、名簿F列の在籍終了日が
'       実施日(J4)より前の児童の行を非表示にする。
'       キャプションが「全員」のときは全児童行を再表示する。
' ボタン名: Btn_enrollment（手動配置、マクロ割り当て: sh_input.ToggleEnrollmentFilter）
'===============================================================================
Public Sub ToggleEnrollmentFilter()
    On Error GoTo ErrorHandler

    Dim btn As Button
    Set btn = sh_input.Buttons("Btn_enrollment")

    Dim childCount As Long
    childCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).Value
    If childCount <= 0 Then Exit Sub

    If btn.Caption = "全員" Then
        ' --- 現在は全員表示中 → 在籍のみに切替 ---

        ' 実施日を取得
        Dim testDate As Variant
        testDate = sh_input.Range(RNG_INPUT_DATE).Value

        If IsEmpty(testDate) Or Not IsDate(testDate) Then
            MsgBox "実施日（J4）が入力されていません。" & vbCrLf & _
                   "実施日を入力してから再度お試しください。", _
                   vbExclamation, "在籍フィルター"
            Exit Sub
        End If

        Dim dtTest As Date
        dtTest = CDate(testDate)

        Dim i As Long
        Dim endDateVal As Variant
        Dim inputRow As Long

        For i = 1 To childCount
            inputRow = eRowInput.rowChildStart + i - 1
            endDateVal = sh_namelist.Cells(NAMELIST_DATA_START_ROW + i - 1, _
                                           NAMELIST_COL_END_DATE).Value

            ' 在籍終了日が入力済み かつ 実施日より前 → 非表示
            If Not IsEmpty(endDateVal) And IsDate(endDateVal) Then
                If CDate(endDateVal) < dtTest Then
                    sh_input.Rows(inputRow).Hidden = True
                End If
            End If
        Next i

        btn.Caption = "在籍"

    Else
        ' --- 現在は在籍のみ表示中 → 全員表示に切替 ---
        Dim startRow As Long
        Dim endRow As Long
        startRow = eRowInput.rowChildStart
        endRow = eRowInput.rowChildStart + childCount - 1

        sh_input.Rows(startRow & ":" & endRow).Hidden = False

        btn.Caption = "全員"
    End If

    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("sh_input", "ToggleEnrollmentFilter")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub
