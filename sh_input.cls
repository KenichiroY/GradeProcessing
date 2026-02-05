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
