'===============================================================================
' モジュール名: sh_subject (Subjectシート)
' 説明: Subjectシートのイベント処理とUI制御
' 修正内容:
'   - 変数宣言をLong型に統一
'   - エラーハンドリング追加
'   - デバッグコード削除
'===============================================================================
Option Explicit

'===============================================================================
' チェックボックスの初期化
'===============================================================================
Public Sub Bysubject_checkbox_initialize()
    On Error Resume Next
    
    Dim i As Long
    
    With sh_subject
        For i = 1 To MAX_PERSPECTIVES
            If Trim(sh_setting.Cells(i + 2, SETTING_PERSPECTIVE_COL).Value & "") = "" Then
                .OLEObjects("perspective" & i).Visible = False
            Else
                .OLEObjects("perspective" & i).Object.Caption = _
                    sh_setting.Cells(i + 2, SETTING_PERSPECTIVE_COL).Value
                .OLEObjects("perspective" & i).Visible = True
            End If
        Next i
    End With
    
    On Error GoTo 0
End Sub

'===============================================================================
' 全チェックボックスを表示
'===============================================================================
Public Sub AllCheckbox_Visible()
    On Error Resume Next
    
    Dim i As Long
    
    For i = 1 To MAX_PERSPECTIVES
        sh_subject.OLEObjects("perspective" & i).Visible = True
    Next i
    
    On Error GoTo 0
End Sub

'===============================================================================
' データ消去ボタンのクリックイベント
'===============================================================================
Private Sub Delete_Sh_Subject_Click()
    On Error GoTo ErrorHandler

    Dim lastCol As Long
    Dim lastRow As Long

    lastCol = sh_subject.Cells(eRowSubject.rowKey, Columns.Count).End(xlToLeft).Column

    ' データがない場合
    If lastCol < eColData.colDataStart Then
        Call ErrorHandlerModule.ShowInfo("消去するデータがありません。")
        Exit Sub
    End If

    ' 確認ダイアログ
    If Not ErrorHandlerModule.ShowConfirmation( _
        "Subjectシートのデータをすべて消去しますか？" & vbCrLf & _
        "（データシートのデータは残ります）", "消去確認") Then
        Exit Sub
    End If

    ' データをクリア（D列以降、児童データ最終行まで動的に計算）
    ' 行範囲: 4行目（rowKey）から児童データ最終行（rowChildStart + MAX_CHILDREN - 1）まで
    lastRow = eRowSubject.rowChildStart + MAX_CHILDREN - 1
    sh_subject.Range(sh_subject.Cells(eRowSubject.rowKey, eColData.colDataStart), _
                     sh_subject.Cells(lastRow, Columns.Count)).ClearContents

    ' 重み正規化状態もリセット
    Call SubjectModule.ResetWeightNormalizedStatus

    Call ErrorHandlerModule.ShowSuccess("データを消去しました。")
    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("sh_subject", "Delete_Sh_Subject_Click")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 更新ボタンのクリックイベント
'===============================================================================
Private Sub Update_Click()
    Call SubjectModule.CollectSubjectData
End Sub

'===============================================================================
' 評価ボタンのクリックイベント
'===============================================================================
Private Sub Ope_result_Click()
    On Error GoTo ErrorHandler
    
    ' 教科・観点の整合性チェック
    If Not SubjectModule.CheckSubjectPerspectiveConsistency() Then
        If Not ErrorHandlerModule.ShowConfirmation( _
            "教科または観点の異なる項目が混在しています。" & vbCrLf & _
            "このまま評価計算を行ってもよろしいですか？" & vbCrLf & vbCrLf & _
            "※この計算は何度でもやり直すことができます。", "確認") Then
            Exit Sub
        End If
    End If
    
    Call SubjectModule.CalculateABCEvaluation
    Exit Sub
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("sh_subject", "Ope_result_Click")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' セルのダブルクリックイベント（ABC閾値の採用/最終決定）
' 7行目: ●（候補）/ ★（採用）の表示
' 8行目: 最終決定をダブルクリックでResult転記
'===============================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim childCount As Long
    Dim lastCol As Long
    Dim finalDecisionCol As Long
    Dim ratioCol As Long

    childCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).Value

    With sh_subject
        ' 8行目の最終列を取得
        lastCol = .Cells(8, Columns.Count).End(xlToLeft).Column

        ' ===========================================
        ' ●（候補）をダブルクリック → ★（採用）に変更
        ' ===========================================
        If Target.Row = 7 And Target.Value = "●" Then
            Cancel = True

            ' 既存の採用（★と採用テキスト）を候補（●と候補テキスト）に戻す
            Call ResetAdoptedToCandidate

            ' クリックした列を採用に変更（7行目:★、8行目:採用）
            Target.Value = "★"
            .Cells(8, Target.Column).Value = "採用"

            ' 最終決定列を探す
            finalDecisionCol = FindFinalDecisionColumn()

            If finalDecisionCol > 0 Then
                ' 閾値をコピー
                .Cells(9, finalDecisionCol).Value = .Cells(9, Target.Column).Value
                .Cells(10, finalDecisionCol).Value = .Cells(10, Target.Column).Value

                ' A/B/C計をコピー
                .Cells(18, Target.Column).Copy .Cells(18, finalDecisionCol)
                .Cells(19, Target.Column).Copy .Cells(19, finalDecisionCol)
                .Cells(20, Target.Column).Copy .Cells(20, finalDecisionCol)

                ' 各児童のABC評価をコピー
                For i = 1 To childCount
                    .Cells(i + eRowSubject.rowChildStart - 1, finalDecisionCol).Value = _
                        .Cells(i + eRowSubject.rowChildStart - 1, Target.Column).Value
                Next i
            End If

            Exit Sub
        End If

        ' ===========================================
        ' 最終決定をダブルクリック → Result転記 + スナップショット保存
        ' ===========================================
        If Target.Row = 8 And Target.Value = "最終決定" Then
            Cancel = True

            ' 採用されているか確認
            If Not HasAdoptedCandidate() Then
                Call ErrorHandlerModule.ShowValidationError( _
                    "閾値が採用されていません。" & vbCrLf & _
                    "先に●をダブルクリックして閾値を採用してください。")
                Exit Sub
            End If

            ' 達成率列と最終決定ABC列を取得
            ratioCol = FindRatioColumn()
            finalDecisionCol = Target.Column

            If ratioCol > 0 Then
                Call ResultModule.FinalizeEvaluation(ratioCol, finalDecisionCol)
            End If

            Exit Sub
        End If

        ' ===========================================
        ' 「候補」テキスト（8行目）をダブルクリック → 採用に変更
        ' ===========================================
        If Target.Row = 8 And Target.Value = "候補" Then
            Cancel = True

            ' 既存の採用（★と採用テキスト）を候補（●と候補テキスト）に戻す
            Call ResetAdoptedToCandidate

            ' クリックした列を採用に変更（7行目:★、8行目:採用）
            .Cells(7, Target.Column).Value = "★"
            Target.Value = "採用"

            ' 最終決定列を探す
            finalDecisionCol = FindFinalDecisionColumn()

            If finalDecisionCol > 0 Then
                ' 閾値をコピー
                .Cells(9, finalDecisionCol).Value = .Cells(9, Target.Column).Value
                .Cells(10, finalDecisionCol).Value = .Cells(10, Target.Column).Value

                ' A/B/C計をコピー
                .Cells(18, Target.Column).Copy .Cells(18, finalDecisionCol)
                .Cells(19, Target.Column).Copy .Cells(19, finalDecisionCol)
                .Cells(20, Target.Column).Copy .Cells(20, finalDecisionCol)

                ' 各児童のABC評価をコピー
                For i = 1 To childCount
                    .Cells(i + eRowSubject.rowChildStart - 1, finalDecisionCol).Value = _
                        .Cells(i + eRowSubject.rowChildStart - 1, Target.Column).Value
                Next i
            End If
        End If
    End With

    Exit Sub

ErrorHandler:
    Debug.Print "Worksheet_BeforeDoubleClick Error: " & Err.Description
End Sub

'===============================================================================
' 採用（★/採用）を候補（●/候補）に戻す
' 7行目と8行目を連動してリセット
'===============================================================================
Private Sub ResetAdoptedToCandidate()
    Dim col As Long
    Dim lastCol As Long

    With sh_subject
        lastCol = .Cells(8, Columns.Count).End(xlToLeft).Column

        For col = eColData.colDataStart To lastCol
            ' 7行目: ★ → ●
            If .Cells(7, col).Value = "★" Then
                .Cells(7, col).Value = "●"
            End If
            ' 8行目: 採用 → 候補
            If .Cells(8, col).Value = "採用" Then
                .Cells(8, col).Value = "候補"
            End If
        Next col
    End With
End Sub

'===============================================================================
' 採用されている候補があるか確認
'===============================================================================
Private Function HasAdoptedCandidate() As Boolean
    Dim col As Long
    Dim lastCol As Long

    HasAdoptedCandidate = False

    With sh_subject
        lastCol = .Cells(7, Columns.Count).End(xlToLeft).Column

        For col = eColData.colDataStart To lastCol
            If .Cells(7, col).Value = "★" Then
                HasAdoptedCandidate = True
                Exit Function
            End If
            If .Cells(8, col).Value = "採用" Then
                HasAdoptedCandidate = True
                Exit Function
            End If
        Next col
    End With
End Function

'===============================================================================
' 最終決定列を検索
'===============================================================================
Private Function FindFinalDecisionColumn() As Long
    Dim col As Long
    Dim lastCol As Long

    FindFinalDecisionColumn = 0

    With sh_subject
        lastCol = .Cells(8, Columns.Count).End(xlToLeft).Column

        For col = lastCol To eColData.colDataStart Step -1
            If .Cells(8, col).Value = "最終決定" Then
                FindFinalDecisionColumn = col
                Exit Function
            End If
        Next col
    End With
End Function

'===============================================================================
' 加重達成率列を検索
'===============================================================================
Private Function FindRatioColumn() As Long
    Dim col As Long
    Dim lastCol As Long

    FindRatioColumn = 0

    With sh_subject
        lastCol = .Cells(8, Columns.Count).End(xlToLeft).Column

        For col = eColData.colDataStart To lastCol
            If .Cells(8, col).Value = "加重達成率" Then
                FindRatioColumn = col
                Exit Function
            End If
        Next col
    End With
End Function

'===============================================================================
' 得点調整行の表示切替ボタン
'===============================================================================
Public Sub Btm_adjustscore_hide_reveal()
    On Error Resume Next
    
    With sh_subject
        If .Range(RNG_SUBJECT_ADJSCORE_DISP).Value = "表示" Then
            .Rows("12:16").Hidden = True
            .Range(RNG_SUBJECT_ADJSCORE_DISP).Value = "非表示"
        Else
            .Rows("12:16").Hidden = False
            .Range(RNG_SUBJECT_ADJSCORE_DISP).Value = "表示"
        End If
    End With
    
    On Error GoTo 0
End Sub

'===============================================================================
' 得点調整の有効/無効切替ボタン
'===============================================================================
Public Sub Btn_IS_adj_score()
    On Error GoTo ErrorHandler
    
    Dim lastCol As Long
    
    With sh_subject
        lastCol = .Cells(eRowSubject.rowKey, Columns.Count).End(xlToLeft).Column
        
        If lastCol >= eColData.colDataStart Then
            Call ErrorHandlerModule.ShowValidationError( _
                "得点調整の切り替えは、データを消去してから行ってください。" & vbCrLf & _
                "「消去」ボタンを押してから切り替えてください。")
            Exit Sub
        End If
        
        If .Range(RNG_SUBJECT_ISADJUST).Value = "有効" Then
            .Range(RNG_SUBJECT_ISADJUST).Value = "無効"
            Call ErrorHandlerModule.ShowInfo("得点調整を無効にしました。")
        Else
            .Range(RNG_SUBJECT_ISADJUST).Value = "有効"
            Call ErrorHandlerModule.ShowInfo("得点調整を有効にしました。")
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("sh_subject", "Btn_IS_adj_score")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 重み正規化ボタンのクリックイベント
' 説明: 配点を基準（100点）に正規化し、すべてのテストを同じ重要度として扱う
'===============================================================================
Public Sub Btn_NormalizeWeight_Click()
    Call SubjectModule.NormalizeWeightByAllocateScore
End Sub
