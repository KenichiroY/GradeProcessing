'===============================================================================
' モジュール名: PostingModule
' 説明: テスト結果の登録（転記）機能を提供
' 修正内容:
'   - 変数宣言をLong型に統一
'   - エラーハンドリング追加
'   - 入力検証の統合
'   - パフォーマンス改善（ScreenUpdating, Calculation）
'===============================================================================
Option Explicit

'===============================================================================
' テストデータ登録メイン処理
'===============================================================================
Public Sub Posting()
    On Error GoTo ErrorHandler
    
    Dim i As Long, j As Long
    Dim hasData As Boolean
    Dim lastRow As Long
    Dim lastRowData As Long
    Dim lastColData As Long
    Dim numTest As Long
    Dim validationResult As ValidationResult
    Dim subjectName As String
    
    ' 処理開始
    Call ErrorHandlerModule.BeginProcess
    
    ' ========================================
    ' 基本情報の取得
    ' ========================================
    lastRow = sh_input.Cells(Rows.Count, 2).End(xlUp).Row
    lastRowData = Sh_data.Cells(Rows.Count, 2).End(xlUp).Row
    lastColData = Sh_data.Cells(eRowData.rowKey, Columns.Count).End(xlToLeft).Column
    
    ' データがない場合は初期列を設定
    If lastColData < eColData.colDataStart Then
        lastColData = eColData.colDataStart - 1
    End If
    
    ' ========================================
    ' 検証1: テスト数上限チェック
    ' ========================================
    validationResult = ValidationModule.ValidateTestCountLimit()
    If Not validationResult.IsValid Then
        Call ErrorHandlerModule.ShowValidationError(validationResult.ErrorMessage)
        GoTo CleanExit
    End If
    
    ' ========================================
    ' 検証2: 必須項目チェック
    ' ========================================
    validationResult = ValidationModule.ValidateRequiredFields()
    If Not validationResult.IsValid Then
        Call ErrorHandlerModule.ShowValidationError(validationResult.ErrorMessage)
        GoTo CleanExit
    End If
    
    ' ========================================
    ' 検証3: 教科名の存在チェック
    ' ========================================
    subjectName = sh_input.Range(RNG_INPUT_SUBJECT).Value
    validationResult = ValidationModule.ValidateSubjectExists(subjectName)
    If Not validationResult.IsValid Then
        Call ErrorHandlerModule.ShowValidationError(validationResult.ErrorMessage)
        GoTo CleanExit
    End If
    
    ' ========================================
    ' 検証4: 得点入力の有無チェック
    ' ========================================
    If Not ValidationModule.HasAnyScoreInput(lastRow) Then
        Call ErrorHandlerModule.ShowValidationError(ERR_MSG_NO_SCORE)
        GoTo CleanExit
    End If
    
    ' ========================================
    ' 検証5: 各列のデータ検証＆登録テスト数カウント
    ' ========================================
    numTest = 0
    
    For i = eColInput.colDataStart To eColInput.colDataEnd
        ' この列にデータがあるかチェック
        hasData = False
        For j = eRowInput.rowChildStart To lastRow
            If Trim(sh_input.Cells(j, i).Value & "") <> "" Then
                hasData = True
                Exit For
            End If
        Next j
        
        If Not hasData Then
            Exit For  ' データがない列以降は処理しない
        End If
        
        ' 得点データの検証
        validationResult = ValidationModule.ValidateScoreData(i, lastRow)
        If Not validationResult.IsValid Then
            Call ErrorHandlerModule.ShowValidationError(validationResult.ErrorMessage)
            ' エラー箇所をハイライト
            If validationResult.ErrorRow > 0 And validationResult.ErrorCol > 0 Then
                sh_input.Cells(validationResult.ErrorRow, validationResult.ErrorCol).Select
            End If
            GoTo CleanExit
        End If
        
        ' クリッピング設定の検証
        validationResult = ValidationModule.ValidateClippingSettings(i)
        If Not validationResult.IsValid Then
            Call ErrorHandlerModule.ShowValidationError(validationResult.ErrorMessage)
            GoTo CleanExit
        End If
        
        ' 重み設定の検証
        validationResult = ValidationModule.ValidateWeight(i)
        If Not validationResult.IsValid Then
            Call ErrorHandlerModule.ShowValidationError(validationResult.ErrorMessage)
            GoTo CleanExit
        End If
        
        numTest = numTest + 1
    Next i
    
    ' テスト数が0の場合（通常はここに来ない）
    If numTest = 0 Then
        Call ErrorHandlerModule.ShowValidationError(ERR_MSG_NO_SCORE)
        GoTo CleanExit
    End If
    
    ' ========================================
    ' データ転記処理
    ' ========================================
    Call TransferData(numTest, lastRow, lastRowData, lastColData)
    
    ' ========================================
    ' 入力フォームのクリア
    ' ========================================
    Call ResetInputForm

    ' ========================================
    ' 得点セルの保護を再設定
    ' ========================================
    Call DataManagementModule.ProtectScoreCells

    ' 成功メッセージ
    Call ErrorHandlerModule.ShowSuccess(MSG_POSTING_SUCCESS)

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub
    
ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("PostingModule", "Posting")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' データ転記の実処理
'===============================================================================
Private Sub TransferData(ByVal numTest As Long, ByVal lastRow As Long, _
                         ByVal lastRowData As Long, ByVal lastColData As Long)
    On Error GoTo ErrorHandler
    
    Dim i As Long, j As Long
    Dim targetCol As Long
    Dim colLetter As String
    Dim weightValue As Variant
    
    With sh_input
        For i = 1 To numTest
            targetCol = lastColData + i
            colLetter = ColumnIndexToLetter(targetCol)
            
            ' ヘッダー情報の転記
            Sh_data.Cells(eRowData.rowKey, targetCol) = AssignKey(.Range(RNG_INPUT_SUBJECT).Value)
            Sh_data.Cells(eRowData.rowTestDate, targetCol) = .Range(RNG_INPUT_DATE).Value
            Sh_data.Cells(eRowData.rowSubject, targetCol) = .Range(RNG_INPUT_SUBJECT).Value
            Sh_data.Cells(eRowData.rowCategory, targetCol) = .Range(RNG_INPUT_CATEGORY).Value
            Sh_data.Cells(eRowData.rowTestName, targetCol) = .Range(RNG_INPUT_TEST_NAME).Value
            Sh_data.Cells(eRowData.rowPerspective, targetCol) = .Cells(eRowInput.rowPerspective, eColInput.colDataStart + i - 1).Value
            Sh_data.Cells(eRowData.rowDetail, targetCol) = .Cells(eRowInput.rowDetail, eColInput.colDataStart + i - 1).Value
            Sh_data.Cells(eRowData.rowAllocationScore, targetCol) = .Cells(eRowInput.rowAllocateScore, eColInput.colDataStart + i - 1).Value
            Sh_data.Cells(eRowData.rowClippingSup, targetCol) = .Cells(eRowInput.rowClippingSup, eColInput.colDataStart + i - 1).Value
            Sh_data.Cells(eRowData.rowClippingInf, targetCol) = .Cells(eRowInput.rowClippingInf, eColInput.colDataStart + i - 1).Value
            Sh_data.Cells(eRowData.rowConvScore, targetCol) = .Cells(eRowInput.rowConvScore, eColInput.colDataStart + i - 1).Value
            Sh_data.Cells(eRowData.rowAdjScoreSup, targetCol) = .Cells(eRowInput.rowAdjScoreSup, eColInput.colDataStart + i - 1).Value
            Sh_data.Cells(eRowData.rowAdjScoreInf, targetCol) = .Cells(eRowInput.rowAdjScoreInf, eColInput.colDataStart + i - 1).Value

            ' 調整後配点の数式を設定
            Sh_data.Cells(eRowData.rowAdjAllocateScore, targetCol).Formula = _
                "=調整後配点計算(" & colLetter & eRowData.rowAllocationScore & "," & _
                colLetter & eRowData.rowClippingSup & "," & _
                colLetter & eRowData.rowConvScore & "," & _
                colLetter & eRowData.rowAdjScoreSup & ")"

            ' 重み（空欄の場合は1）
            weightValue = .Cells(eRowInput.rowWeight, eColInput.colDataStart + i - 1).Value
            If Trim(weightValue & "") = "" Then
                Sh_data.Cells(eRowData.rowWeight, targetCol) = 1
            Else
                Sh_data.Cells(eRowData.rowWeight, targetCol) = weightValue
            End If
            
            ' 統計計算式の設定
            Sh_data.Cells(eRowData.rowAverage, targetCol).Formula = _
                "=AVERAGE(" & colLetter & eRowData.rowChildStart & ":" & colLetter & lastRowData & ")"
            Sh_data.Cells(eRowData.rowMedian, targetCol).Formula = _
                "=MEDIAN(" & colLetter & eRowData.rowChildStart & ":" & colLetter & lastRowData & ")"
            Sh_data.Cells(eRowData.rowStdDev, targetCol).Formula = _
                "=STDEV.P(" & colLetter & eRowData.rowChildStart & ":" & colLetter & lastRowData & ")"
            Sh_data.Cells(eRowData.rowCV, targetCol).Formula = _
                "=" & colLetter & eRowData.rowStdDev & "/" & colLetter & eRowData.rowAverage
            
            ' 児童の得点データ転記
            For j = 1 To lastRow - eRowInput.rowChildStart + 1
                Sh_data.Cells(eRowData.rowChildStart + j - 1, targetCol) = _
                    .Cells(eRowInput.rowChildStart + j - 1, eColInput.colDataStart + i - 1).Value
            Next j
        Next i
    End With
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "PostingModule.TransferData", Err.Description
End Sub

'===============================================================================
' 入力フォームのリセット
'===============================================================================
Public Sub ResetInputForm()
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = sh_input.Cells(Rows.Count, 2).End(xlUp).Row
    
    With sh_input
        .Range(RNG_INPUT_SUBJECT).ClearContents
        .Range(RNG_INPUT_CATEGORY).ClearContents
        .Range(RNG_INPUT_TEST_NAME).MergeArea.ClearContents
        .Range(RNG_INPUT_TEST_REMARK).ClearContents
        
        ' 評価観点～調整範囲下限をクリア
        .Range(.Cells(eRowInput.rowPerspective, eColInput.colDataStart), _
               .Cells(eRowInput.rowAdjScoreInf, eColInput.colDataEnd)).ClearContents
        
        ' 重みをクリア
        .Range(.Cells(eRowInput.rowWeight, eColInput.colDataStart), _
               .Cells(eRowInput.rowWeight, eColInput.colDataEnd)).ClearContents
        
        ' 児童の得点データをクリア
        .Range(.Cells(eRowInput.rowChildStart, eColInput.colDataStart), _
               .Cells(lastRow, eColInput.colDataEnd)).ClearContents
    End With
    
    Exit Sub
    
ErrorHandler:
    ' リセット処理のエラーは無視（メイン処理は成功しているため）
    Debug.Print "ResetInputForm Error: " & Err.Description
End Sub

'===============================================================================
' テストキーの自動採番
' 引数: subjectName - 教科名
' 戻り値: 新しいテストキー（例: J001, S002）
'===============================================================================
Private Function AssignKey(ByVal subjectName As String) As String
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim keyChar As String
    Dim keyCount As Long
    
    i = SETTING_SUBJECT_START_ROW
    
    With sh_setting
        Do While Trim(.Cells(i, SETTING_SUBJECT_COL).Value & "") <> ""
            If .Cells(i, SETTING_SUBJECT_COL).Value = subjectName Then
                keyChar = .Cells(i, SETTING_KEY_CHAR_COL).Value
                keyCount = Val(.Cells(i, SETTING_KEY_COUNT_COL).Value) + 1
                
                ' カウンタを更新
                .Cells(i, SETTING_KEY_COUNT_COL).Value = keyCount
                
                ' キーを生成して返す
                AssignKey = keyChar & Format(keyCount, "000")
                Exit Function
            End If
            i = i + 1
        Loop
    End With
    
    ' 見つからない場合（通常は検証で弾かれるので来ない）
    AssignKey = "ERR001"
    Exit Function
    
ErrorHandler:
    AssignKey = "ERR" & Format(Err.Number, "000")
End Function

'===============================================================================
' 列番号を列文字に変換
' 引数: colIndex - 列番号（1始まり）
' 戻り値: 列文字（A, B, ..., Z, AA, AB, ...）
'===============================================================================
Public Function ColumnIndexToLetter(ByVal colIndex As Long) As String
    Dim colAddress As String
    colAddress = Columns(colIndex).Address
    ColumnIndexToLetter = Split(colAddress, "$")(2)
End Function
