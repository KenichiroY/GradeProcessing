Attribute VB_Name = "SubjectModule"
'===============================================================================
' モジュール名: SubjectModule
' 説明: 教科別のデータ集計・ABC評価計算機能を提供
' 修正内容:
'   - 変数宣言をLong型に統一
'   - エラーハンドリング追加
'   - パフォーマンス改善
'   - デバッグコード削除
'===============================================================================
Option Explicit

'===============================================================================
' 教科別データ収集
' 説明: データシートから指定教科・観点のテストデータを収集し、Subjectシートに転記
'===============================================================================
Public Sub CollectSubjectData()
    On Error GoTo ErrorHandler
    
    Dim i As Long, j As Long
    Dim childCount As Long
    Dim targetSubject As String
    Dim scoreList() As Variant
    Dim SubjectInfo() As Variant
    Dim perspective() As Variant
    Dim lastColDatabase As Long
    Dim lastColSubject As Long
    Dim foundFlag As Boolean
    Dim perspectiveCount As Long
    
    ' 処理開始
    Call ErrorHandlerModule.BeginProcess
    
    ' 児童数取得
    childCount = GetChildCount()
    If childCount = 0 Then
        Call ErrorHandlerModule.ShowValidationError("名簿に児童が登録されていません。")
        GoTo CleanExit
    End If
    
    ' 目的の教科を取得
    targetSubject = sh_subject.Range(RNG_SUBJECT_SUBJECT).value
    If Trim(targetSubject) = "" Then
        Call ErrorHandlerModule.ShowValidationError("教科を選択してください。")
        GoTo CleanExit
    End If
    
    ' 最終列の取得
    lastColDatabase = Sh_data.Cells(eRowData.rowKey, Columns.count).End(xlToLeft).Column
    lastColSubject = sh_subject.Cells(eRowSubject.rowKey, Columns.count).End(xlToLeft).Column
    
    ' データがない場合
    If lastColDatabase < eColData.colDataStart Then
        Call ErrorHandlerModule.ShowInfo("登録されているテストデータがありません。")
        GoTo CleanExit
    End If
    
    ' 観点チェックボックスから選択された観点を取得
    perspectiveCount = GetSelectedPerspectives(perspective)
    If perspectiveCount = 0 Then
        Call ErrorHandlerModule.ShowValidationError("評価観点を1つ以上選択してください。")
        GoTo CleanExit
    End If
    
    ' 配列の初期化
    ReDim SubjectInfo(14, 0)
    ReDim scoreList(childCount, 0)
    
    ' データシートから該当データを抽出
    With Sh_data
        For i = eColData.colDataStart To lastColDatabase
            ' 教科が一致するかチェック
            If CStr(.Cells(eRowData.rowSubject, i).value) = targetSubject Then
                ' 追試中チェック（"N" が1つでもあれば除外）
                If hasRetestMarker(i) Then
                    GoTo NextColumn
                End If

                ' 観点が選択されているかチェック
                If IsPerspectiveSelected(CStr(.Cells(eRowData.rowPerspective, i).value), perspective) Then
                    ' 配列のサイズを拡張
                    ReDim Preserve SubjectInfo(14, UBound(SubjectInfo, 2) + 1)
                    ReDim Preserve scoreList(childCount, UBound(scoreList, 2) + 1)
                    
                    ' テスト情報を格納
                    For j = 0 To 14
                        SubjectInfo(j, UBound(SubjectInfo, 2)) = .Cells(j + eRowData.rowKey, i).value
                    Next j
                    
                    ' 得点データを格納
                    For j = 1 To childCount
                        scoreList(j, UBound(scoreList, 2)) = .Cells(j + eRowData.rowChildStart - 1, i).value
                    Next j
                End If
            End If
NextColumn:
        Next i
    End With

    ' 抽出データがない場合
    If UBound(scoreList, 2) = 0 Then
        Call ErrorHandlerModule.ShowInfo("選択した条件に一致するテストデータがありません。")
        GoTo CleanExit
    End If
    
    ' Subjectシートへの書き込み
    Call WriteToSubjectSheet(SubjectInfo, scoreList, childCount, lastColSubject)
    
    
CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub
    
ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("SubjectModule", "CollectSubjectData")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' Subjectシートへのデータ書き込み
'===============================================================================
Private Sub WriteToSubjectSheet(ByRef SubjectInfo() As Variant, ByRef scoreList() As Variant, _
                                ByVal childCount As Long, ByVal lastColSubject As Long)
    On Error GoTo ErrorHandler
    
    Dim i As Long, j As Long
    Dim foundFlag As Boolean
    Dim colLetter As String
    Dim isAdjustEnabled As Boolean
    
    isAdjustEnabled = (sh_subject.Range(RNG_SUBJECT_ISADJUST).value = "有効")
    
    With sh_subject
        For i = 1 To UBound(scoreList, 2)
            ' 既に登録済みかチェック
            foundFlag = False
            For j = eColData.colDataStart To lastColSubject
                If .Cells(eRowSubject.rowKey, j).value = SubjectInfo(0, i) Then
                    foundFlag = True
                    Exit For
                End If
            Next j
            
            If Not foundFlag Then
                ' 新規データの場合、書き込み
                lastColSubject = .Cells(eRowSubject.rowKey, Columns.count).End(xlToLeft).Column + 1
                If lastColSubject < eColData.colDataStart Then
                    lastColSubject = eColData.colDataStart
                End If
                
                colLetter = PostingModule.ColumnIndexToLetter(lastColSubject)
                
                ' テスト情報の転記
                For j = 0 To 14
                    .Cells(j + eRowSubject.rowKey, lastColSubject).value = SubjectInfo(j, i)
                Next j
                
                ' 調整後配点の数式
                .Cells(eRowSubject.rowAdjAllocateScore, lastColSubject).formula = _
                    "=調整後配点計算(" & colLetter & eRowSubject.rowAllocationScore & "," & _
                    colLetter & eRowSubject.rowClippingSup & "," & _
                    colLetter & eRowSubject.rowConvScore & "," & _
                    colLetter & eRowSubject.rowAdjScoreSup & ")"
                
                ' 統計値の数式
                .Cells(eRowSubject.rowAverage, lastColSubject).formula = _
                    "=IFERROR(AVERAGE(" & colLetter & eRowSubject.rowChildStart & ":" & _
                    colLetter & (eRowSubject.rowChildStart + childCount - 1) & "),"""")"
                .Cells(eRowSubject.rowMedian, lastColSubject).formula = _
                    "=IFERROR(MEDIAN(" & colLetter & eRowSubject.rowChildStart & ":" & _
                    colLetter & (eRowSubject.rowChildStart + childCount - 1) & "),"""")"
                .Cells(eRowSubject.rowStdDev, lastColSubject).formula = _
                    "=IFERROR(STDEV.P(" & colLetter & eRowSubject.rowChildStart & ":" & _
                    colLetter & (eRowSubject.rowChildStart + childCount - 1) & "),"""")"
                .Cells(eRowSubject.rowCV, lastColSubject).formula = _
                    "=IFERROR(" & colLetter & eRowSubject.rowStdDev & "/" & colLetter & eRowSubject.rowAverage & ","""")"
                
                ' 得点データの転記
                For j = 1 To childCount
                    If Trim(scoreList(j, i) & "") = "" Then
                        ' 空欄はそのまま
                    ElseIf scoreList(j, i) = "-" Then
                        .Cells(j + eRowSubject.rowChildStart - 1, lastColSubject).value = "-"
                    Else
                        If isAdjustEnabled Then
                            ' 得点調整が有効な場合
                            .Cells(j + eRowSubject.rowChildStart - 1, lastColSubject).value = _
                                調整後得点計算_shsubject(CDbl(scoreList(j, i)), lastColSubject)
                        Else
                            .Cells(j + eRowSubject.rowChildStart - 1, lastColSubject).value = scoreList(j, i)
                        End If
                    End If
                Next j
            End If
        Next i
    End With
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "SubjectModule.WriteToSubjectSheet", Err.Description
End Sub

'===============================================================================
' ABC評価の計算
'===============================================================================
Public Sub CalculateABCEvaluation()
    On Error GoTo ErrorHandler
    
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim abcThresholdCount As Long
    Dim childCount As Long
    Dim colLetter As String
    Dim isAdjustEnabled As Boolean
    
    ' 処理開始
    Call ErrorHandlerModule.BeginProcess
    
    ' データ取得
    lastCol = sh_subject.Cells(eRowSubject.rowKey, Columns.count).End(xlToLeft).Column
    childCount = GetChildCount()
    
    ' データがない場合
    If lastCol < eColData.colDataStart Then
        Call ErrorHandlerModule.ShowInfo("評価対象のデータがありません。" & vbCrLf & _
            "先に「更新」ボタンでデータを収集してください。")
        GoTo CleanExit
    End If
    
    ' ABC閾値の数を取得
    abcThresholdCount = Application.WorksheetFunction.CountA(sh_setting.Range("H3:H20"))
    If abcThresholdCount = 0 Then
        Call ErrorHandlerModule.ShowValidationError("ABC評価の閾値がSettingシートに設定されていません。")
        GoTo CleanExit
    End If
    
    isAdjustEnabled = (sh_subject.Range(RNG_SUBJECT_ISADJUST).value = "有効")
    
    ' ヘッダーの作成
    Call CreateEvaluationHeaders(lastCol, abcThresholdCount)
    
    ' 各児童の評価計算
    Call CalculateChildEvaluations(lastCol, childCount, abcThresholdCount, isAdjustEnabled)
    
    sh_subject.Rows("19:22").Hidden = False
    sh_subject.Range(RNG_SUBJECT_STATS_DISP).value = "表示"

    
CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub
    
ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("SubjectModule", "CalculateABCEvaluation")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 評価ヘッダーの作成
'===============================================================================
Private Sub CreateEvaluationHeaders(ByVal lastCol As Long, ByVal abcThresholdCount As Long)
    Dim j As Long
    Dim colLetter As String
    
    With sh_subject
        ' 集計列のヘッダー
        .Cells(8, lastCol + eColShiftSubject.colNoWeightSum) = "重み無し合計"
        .Cells(8, lastCol + eColShiftSubject.colNoWeightAllocate) = "重み無し配点"
        .Cells(8, lastCol + eColShiftSubject.colIncludeWeightSum) = "加重合計"
        .Cells(8, lastCol + eColShiftSubject.colIncludeWeightAllocate) = "加重配点"
        .Cells(8, lastCol + eColShiftSubject.colNoWeightRatio) = "重み無し達成率"
        .Cells(8, lastCol + eColShiftSubject.colIncludeWeightRatio) = "加重達成率"
        .Cells(8, lastCol + eColShiftSubject.colABCBorder) = "ABC閾値"
        .Cells(9, lastCol + eColShiftSubject.colABCBorder) = "A/B"
        .Cells(10, lastCol + eColShiftSubject.colABCBorder) = "B/C"
        
        ' 素配点・調整後配点の集計式
        colLetter = PostingModule.ColumnIndexToLetter(lastCol)
        .Cells(eRowSubject.rowAllocationScore, lastCol + eColShiftSubject.colNoWeightSum).formula = _
            "=SUM(D" & eRowSubject.rowAllocationScore & ":" & colLetter & eRowSubject.rowAllocationScore & ")"
        .Cells(eRowSubject.rowAdjAllocateScore, lastCol + eColShiftSubject.colNoWeightSum).formula = _
            "=SUM(D" & eRowSubject.rowAdjAllocateScore & ":" & colLetter & eRowSubject.rowAdjAllocateScore & ")"
        
        ' ABC閾値候補
        For j = 1 To abcThresholdCount
            .Cells(7, lastCol + eColShiftSubject.colABCBorder + j) = "●"  ' 候補マーク
            .Cells(8, lastCol + eColShiftSubject.colABCBorder + j) = "候補"
            .Cells(9, lastCol + eColShiftSubject.colABCBorder + j) = sh_setting.Cells(j + 2, SETTING_AB_THRESHOLD_COL).value
            .Cells(10, lastCol + eColShiftSubject.colABCBorder + j) = sh_setting.Cells(j + 2, SETTING_BC_THRESHOLD_COL).value
        Next j
        
        ' A/B/C計ラベル
        .Cells(18, lastCol + eColShiftSubject.colABCBorder) = "A計"
        .Cells(19, lastCol + eColShiftSubject.colABCBorder) = "B計"
        .Cells(20, lastCol + eColShiftSubject.colABCBorder) = "C計"
        
        ' 最終決定ラベル
        .Cells(8, lastCol + eColShiftSubject.colABCBorder + abcThresholdCount + 2) = "最終決定"
    End With
End Sub

'===============================================================================
' 各児童の評価計算
'===============================================================================
Private Sub CalculateChildEvaluations(ByVal lastCol As Long, ByVal childCount As Long, _
                                       ByVal abcThresholdCount As Long, ByVal isAdjustEnabled As Boolean)
    Dim i As Long, j As Long
    Dim rowNum As Long
    Dim colLetter As String
    Dim colLetterSum As String, colLetterAllocate As String, colLetterRatio As String
    Dim colLetterBorder As String
    Dim colLetterWSum As String, colLetterWAllocate As String
    Dim allocateRow As Long
    Dim rowABThreshold As Long
    Dim rowBCThreshold As Long

    ' ABC閾値の行番号（Subjectシート上の固定位置）
    rowABThreshold = 9   ' A/B閾値行
    rowBCThreshold = 10  ' B/C閾値行

    colLetter = PostingModule.ColumnIndexToLetter(lastCol)
    colLetterSum = PostingModule.ColumnIndexToLetter(lastCol + eColShiftSubject.colNoWeightSum)
    colLetterAllocate = PostingModule.ColumnIndexToLetter(lastCol + eColShiftSubject.colNoWeightAllocate)
    colLetterRatio = PostingModule.ColumnIndexToLetter(lastCol + eColShiftSubject.colIncludeWeightRatio)
    colLetterWSum = PostingModule.ColumnIndexToLetter(lastCol + eColShiftSubject.colIncludeWeightSum)
    colLetterWAllocate = PostingModule.ColumnIndexToLetter(lastCol + eColShiftSubject.colIncludeWeightAllocate)

    ' 配点行（調整有効/無効で変わる）
    If isAdjustEnabled Then
        allocateRow = eRowSubject.rowAdjAllocateScore
    Else
        allocateRow = eRowSubject.rowAllocationScore
    End If

    With sh_subject
        For i = 1 To childCount
            rowNum = i + eRowSubject.rowChildStart - 1

            ' 重み無し合計
            .Cells(rowNum, lastCol + eColShiftSubject.colNoWeightSum).formula = _
                "=SUM(D" & rowNum & ":" & colLetter & rowNum & ")"

            ' 加重合計
            .Cells(rowNum, lastCol + eColShiftSubject.colIncludeWeightSum).formula = _
                "=SUMPRODUCT(D" & rowNum & ":" & colLetter & rowNum & ",D" & eRowSubject.rowWeight & ":" & colLetter & eRowSubject.rowWeight & ")"

            ' 重み無し配点（欠席"-"を除外）
            .Cells(rowNum, lastCol + eColShiftSubject.colNoWeightAllocate).formula = _
                "=SUMPRODUCT((D" & rowNum & ":" & colLetter & rowNum & "<>""-"")*D" & allocateRow & ":" & colLetter & allocateRow & ")"

            ' 加重配点（欠席"-"を除外）
            .Cells(rowNum, lastCol + eColShiftSubject.colIncludeWeightAllocate).formula = _
                "=SUMPRODUCT((D" & rowNum & ":" & colLetter & rowNum & "<>""-"")*D" & allocateRow & ":" & colLetter & allocateRow & "*D" & eRowSubject.rowWeight & ":" & colLetter & eRowSubject.rowWeight & ")"

            ' 重み無し達成率
            .Cells(rowNum, lastCol + eColShiftSubject.colNoWeightRatio).formula = _
                "=ROUND(100*" & colLetterSum & rowNum & "/" & colLetterAllocate & rowNum & ",1)"

            ' 加重達成率
            .Cells(rowNum, lastCol + eColShiftSubject.colIncludeWeightRatio).formula = _
                "=ROUND(100*" & colLetterWSum & rowNum & "/" & colLetterWAllocate & rowNum & ",1)"

            ' ABC判定（各閾値パターン）
            For j = 1 To abcThresholdCount
                colLetterBorder = PostingModule.ColumnIndexToLetter(lastCol + eColShiftSubject.colABCBorder + j)
                .Cells(rowNum, lastCol + eColShiftSubject.colABCBorder + j).formula = _
                    "=IF(" & colLetterRatio & rowNum & ">=" & colLetterBorder & rowABThreshold & ",""A"",IF(" & _
                    colLetterRatio & rowNum & ">=" & colLetterBorder & rowBCThreshold & ",""B"",""C""))"
            Next j
        Next i

        ' A/B/C計の数式
        For j = 1 To abcThresholdCount
            colLetterBorder = PostingModule.ColumnIndexToLetter(lastCol + eColShiftSubject.colABCBorder + j)
            .Cells(18, lastCol + eColShiftSubject.colABCBorder + j).formula = _
                "=COUNTIF(" & colLetterBorder & eRowSubject.rowChildStart & ":" & _
                colLetterBorder & (eRowSubject.rowChildStart + childCount - 1) & ",""A"")"
            .Cells(19, lastCol + eColShiftSubject.colABCBorder + j).formula = _
                "=COUNTIF(" & colLetterBorder & eRowSubject.rowChildStart & ":" & _
                colLetterBorder & (eRowSubject.rowChildStart + childCount - 1) & ",""B"")"
            .Cells(20, lastCol + eColShiftSubject.colABCBorder + j).formula = _
                "=COUNTIF(" & colLetterBorder & eRowSubject.rowChildStart & ":" & _
                colLetterBorder & (eRowSubject.rowChildStart + childCount - 1) & ",""C"")"
        Next j
    End With
End Sub

'===============================================================================
' 選択された観点を取得
'===============================================================================
Private Function GetSelectedPerspectives(ByRef perspective() As Variant) As Long
    Dim i As Long
    Dim count As Long
    
    count = Application.WorksheetFunction.CountA(sh_setting.Range("D3:D16"))
    ReDim perspective(count)
    
    GetSelectedPerspectives = 0
    
    For i = 1 To count
        On Error Resume Next
        If sh_subject.CheckBoxes("perspective" & i).value = xlOn Then
            perspective(i) = sh_subject.CheckBoxes("perspective" & i).Caption
            GetSelectedPerspectives = GetSelectedPerspectives + 1
        Else
            perspective(i) = ""
        End If
        On Error GoTo 0
    Next i
End Function

'===============================================================================
' 観点が選択されているかチェック
'===============================================================================
Private Function IsPerspectiveSelected(ByVal perspectiveName As String, ByRef perspective() As Variant) As Boolean
    Dim i As Long
    
    IsPerspectiveSelected = False
    
    For i = 1 To UBound(perspective)
        If perspectiveName = perspective(i) Then
            IsPerspectiveSelected = True
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' 児童数を取得
'===============================================================================
Private Function GetChildCount() As Long
    On Error Resume Next
    GetChildCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).value
    If Err.Number <> 0 Or GetChildCount < 0 Then
        GetChildCount = 0
    End If
    On Error GoTo 0
End Function

'===============================================================================
' 教科・観点の整合性チェック
'===============================================================================
Public Function CheckSubjectPerspectiveConsistency() As Boolean
    Dim i As Long
    Dim lastCol As Long
    
    lastCol = sh_subject.Cells(eRowSubject.rowKey, Columns.count).End(xlToLeft).Column
    
    If lastCol < eColData.colDataStart Then
        CheckSubjectPerspectiveConsistency = True
        Exit Function
    End If
    
    With sh_subject
        For i = eColData.colDataStart To lastCol - 1
            If .Cells(eRowSubject.rowSubject, i).value <> .Cells(eRowSubject.rowSubject, i + 1).value Or _
               .Cells(eRowSubject.rowPerspective, i).value <> .Cells(eRowSubject.rowPerspective, i + 1).value Then
                CheckSubjectPerspectiveConsistency = False
                Exit Function
            End If
        Next i
    End With
    
    CheckSubjectPerspectiveConsistency = True
End Function

'===============================================================================
' 重みで配点を正規化
' 説明: 各テストの配点を基準配点（100点）に換算し、重みを調整する
'       これにより、配点の異なるテストを同じ重要度として扱える
'===============================================================================
Public Sub NormalizeWeightByAllocateScore()
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim lastCol As Long
    Dim allocateScore As Double
    Dim currentWeight As Double
    Dim newWeight As Double
    Dim isAdjustEnabled As Boolean
    Dim allocateRow As Long
    Dim testCount As Long

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' データがあるか確認
    lastCol = sh_subject.Cells(eRowSubject.rowKey, Columns.count).End(xlToLeft).Column

    If lastCol < eColData.colDataStart Then
        Call ErrorHandlerModule.ShowInfo("正規化するデータがありません。" & vbCrLf & _
            "先に「追加」ボタンでデータを収集してください。")
        GoTo CleanExit
    End If

    ' 既に正規化済みかチェック
    If sh_subject.Range(RNG_SUBJECT_WEIGHT_NORMALIZED).value = "実施済" Then
        Call ErrorHandlerModule.ShowInfo("既に重み正規化が実施されています。" & vbCrLf & _
            "再度正規化する場合は、「消去」→「追加」でデータを再取得してください。")
        GoTo CleanExit
    End If

    ' 確認ダイアログ
    If Not ErrorHandlerModule.ShowConfirmation( _
        "重みを正規化しますか？" & vbCrLf & vbCrLf & _
        "この操作は現在の重み設定を上書きします。" & vbCrLf & _
        "手動で設定した重みがある場合は失われます。" & vbCrLf & vbCrLf & _
        "※元に戻すには「消去」→「追加」で" & vbCrLf & _
        "　データを再取得してください。", _
        "重み正規化の確認") Then
        GoTo CleanExit
    End If

    ' 得点調整の有効/無効を確認（配点行を決定）
    isAdjustEnabled = (sh_subject.Range(RNG_SUBJECT_ISADJUST).value = "有効")
    If isAdjustEnabled Then
        allocateRow = eRowSubject.rowAdjAllocateScore
    Else
        allocateRow = eRowSubject.rowAllocationScore
    End If

    testCount = lastCol - eColData.colDataStart + 1

    ' 各テストの重みを正規化
    With sh_subject
        For i = eColData.colDataStart To lastCol
            ' 配点を取得
            allocateScore = 0
            If IsNumeric(.Cells(allocateRow, i).value) Then
                allocateScore = CDbl(.Cells(allocateRow, i).value)
            End If

            ' 配点が0以下の場合はスキップ（エラー防止）
            If allocateScore <= 0 Then
                GoTo NextTest
            End If

            ' 現在の重みを取得（空欄は1として扱う）
            currentWeight = 1
            If IsNumeric(.Cells(eRowSubject.rowWeight, i).value) Then
                If .Cells(eRowSubject.rowWeight, i).value <> "" Then
                    currentWeight = CDbl(.Cells(eRowSubject.rowWeight, i).value)
                End If
            End If

            ' 正規化した重みを計算: (基準配点 / 配点) × 現在の重み
            newWeight = (NORMALIZE_BASE_SCORE / allocateScore) * currentWeight

            ' 小数点以下2桁に丸める
            newWeight = Round(newWeight, 2)

            ' 重みを更新
            .Cells(eRowSubject.rowWeight, i).value = newWeight
NextTest:
        Next i

        ' 正規化済みフラグを設定
        .Range(RNG_SUBJECT_WEIGHT_NORMALIZED).value = "実施済"
    End With

    Call ErrorHandlerModule.ShowSuccess( _
        "重みの正規化が完了しました。" & vbCrLf & vbCrLf & _
        "対象テスト数: " & testCount & " 件" & vbCrLf & _
        "基準配点: " & NORMALIZE_BASE_SCORE & " 点")

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("SubjectModule", "NormalizeWeightByAllocateScore")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 重み正規化状態をリセット
' 説明: データ消去時に正規化状態もリセットする
'===============================================================================
Public Sub ResetWeightNormalizedStatus()
    On Error Resume Next
    sh_subject.Range(RNG_SUBJECT_WEIGHT_NORMALIZED).value = ""
    On Error GoTo 0
End Sub

'===============================================================================
' 指定列に追試中マーカー "N" があるか確認
' 引数：colIndex - データシートの列番号
' 戻り値：True = "N" が1つ以上ある（追試未完了）
'===============================================================================
Private Function hasRetestMarker(ByVal colIndex As Long) As Boolean
    Dim j As Long
    Dim childCount As Long

    hasRetestMarker = False
    childCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).value

    With Sh_data
        For j = eRowData.rowChildStart To eRowData.rowChildStart + childCount - 1
            If CStr(.Cells(j, colIndex).value) = RETEST_MARKER Then
                hasRetestMarker = True
                Exit Function
            End If
        Next j
    End With
End Function




