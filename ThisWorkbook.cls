'===============================================================================
' モジュール名: ThisWorkbook
' 説明: ワークブックレベルのイベント処理
' 修正内容:
'   - コメントアウトされていたコードを整理
'   - エラーハンドリング追加
'   - 初期化処理の整理
'===============================================================================
Option Explicit

'===============================================================================
' ワークブックを開いた時の処理
'===============================================================================
Private Sub Workbook_Open()
    On Error GoTo ErrorHandler

    ' 入力フォームの日付を更新
    Call InitializeInputForm

    ' Subjectシートのチェックボックスを初期化
    Call InitializeSubjectCheckboxes

    ' Resultシートの列見出しを初期化
    Call InitializeResultHeaders

    ' データシートの得点セルを保護
    Call InitializeDataSheetProtection

    Exit Sub
    
ErrorHandler:
    ' 起動時エラーはユーザーに通知
    MsgBox "初期化中にエラーが発生しました。" & vbCrLf & _
           "一部の機能が正常に動作しない可能性があります。" & vbCrLf & vbCrLf & _
           "エラー: " & Err.Description, vbExclamation, "成績処理システム"
End Sub

'===============================================================================
' 入力フォームの初期化
'===============================================================================
Private Sub InitializeInputForm()
    On Error Resume Next
    
    ' 実施日に今日の日付を設定
    sh_input.Range(RNG_INPUT_DATE).Value = Date
    
    On Error GoTo 0
End Sub

'===============================================================================
' Subjectシートのチェックボックス初期化
'===============================================================================
Private Sub InitializeSubjectCheckboxes()
    On Error Resume Next
    
    Dim i As Long
    
    With sh_subject
        For i = 1 To MAX_PERSPECTIVES
            If Trim(sh_setting.Cells(i + 2, SETTING_PERSPECTIVE_COL).Value & "") = "" Then
                ' 観点が設定されていない場合は非表示
                .OLEObjects("perspective" & i).Visible = False
            Else
                ' 観点名を設定して表示
                .OLEObjects("perspective" & i).Object.Caption = _
                    sh_setting.Cells(i + 2, SETTING_PERSPECTIVE_COL).Value
                .OLEObjects("perspective" & i).Visible = True
            End If
        Next i
    End With
    
    On Error GoTo 0
End Sub

'===============================================================================
' Resultシートの列見出し初期化
'===============================================================================
Private Sub InitializeResultHeaders()
    On Error Resume Next

    ' Resultシートの列見出しを自動生成
    Call ResultModule.GenerateResultHeaders

    On Error GoTo 0
End Sub

'===============================================================================
' データシートの得点セル保護
'===============================================================================
Private Sub InitializeDataSheetProtection()
    On Error Resume Next

    ' 得点セルのみを保護（ダブルクリックでフォームから修正可能）
    Call DataManagementModule.ProtectScoreCells

    On Error GoTo 0
End Sub

'===============================================================================
' ワークブックを閉じる前の処理
'===============================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' 必要に応じて保存確認などを追加
    
    On Error GoTo 0
End Sub

'===============================================================================
' ワークブック保存前の処理
'===============================================================================
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    
    ' 保存前にデータの整合性チェックなどを行う場合はここに追加
    
    On Error GoTo 0
End Sub
