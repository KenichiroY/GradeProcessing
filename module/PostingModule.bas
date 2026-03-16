Attribute VB_Name = "PostingModule"
'===============================================================================
' モジュール名: PostingModule
' 用途: テスト結果の登録（転記）機能群
' メンテナンス:
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
    Dim validationResult As validationResult
    Dim subjectName As String
    Dim hasAnyRetest As Boolean
    Dim retestFlags() As Boolean   ' 列ごとの追試フラグ

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' ========================================
    ' 基本情報の取得
    ' ========================================
    lastRow = sh_input.Cells(Rows.count, 2).End(xlUp).Row
    lastRowData = Sh_data.Cells(Rows.count, 2).End(xlUp).Row
    lastColData = Sh_data.Cells(eRowData.rowKey, Columns.count).End(xlToLeft).Column

    ' データがない場合は初期値設定
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
    subjectName = sh_input.Range(RNG_INPUT_SUBJECT).value
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
    ' 検証5: 各列のデータ検証＋登録テスト数カウント
    ' ========================================
    numTest = 0

    For i = eColInput.colDataStart To eColInput.colDataEnd
        ' この列にデータがあるかチェック
        hasData = False
        For j = eRowInput.rowChildStart To lastRow
            If Trim(sh_input.Cells(j, i).value & "") <> "" Then
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
    ' 追試設定の判定（列ごとのセル値から）
    ' ========================================
    ReDim retestFlags(1 To numTest)
    hasAnyRetest = False
    For i = 1 To numTest
        retestFlags(i) = IsRetestEnabledForColumn(i)
        If retestFlags(i) Then
            hasAnyRetest = True
        End If
    Next i

    ' ========================================
    ' データ転記処理
    ' ========================================
    Call TransferData(numTest, lastRow, lastRowData, lastColData)

    ' ========================================
    ' 追試処理（追試ONの列のみ）
    ' ========================================
    If hasAnyRetest Then
        ' 追試ONの列のデータシート得点を "N" で上書き
        Call MarkAsRetestPending(numTest, lastRowData, lastColData, retestFlags)

        ' 追試ファイルへの転記（追試ONの列のみ）
        Call RetestModule.CreateRetestSheet(numTest, lastRow, lastColData, retestFlags)
    End If

    ' ========================================
    ' 入力フォームのクリア
    ' ========================================
    Call ResetInputForm

    ' ========================================
    ' 得点セルの保護を再設定
    ' ========================================
    Call DataManagementModule.ProtectScoreCells

    ' 完了メッセージ
    If Not hasAnyRetest Then
        Call ErrorHandlerModule.ShowSuccess(MSG_POSTING_SUCCESS)
    End If

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
            Sh_data.Cells(eRowData.rowKey, targetCol) = AssignKey(.Range(RNG_INPUT_SUBJECT).value)
            Sh_data.Cells(eRowData.rowTestDate, targetCol) = .Range(RNG_INPUT_DATE).value
            Sh_data.Cells(eRowData.rowSubject, targetCol) = .Range(RNG_INPUT_SUBJECT).value
            Sh_data.Cells(eRowData.rowCategory, targetCol) = .Range(RNG_INPUT_CATEGORY).value
            Sh_data.Cells(eRowData.rowTestName, targetCol) = .Range(RNG_INPUT_TEST_NAME).value
            Sh_data.Cells(eRowData.rowPerspective, targetCol) = .Cells(eRowInput.rowPerspective, eColInput.colDataStart + i - 1).value
            Sh_data.Cells(eRowData.rowDetail, targetCol) = .Cells(eRowInput.rowDetail, eColInput.colDataStart + i - 1).value
            Sh_data.Cells(eRowData.rowAllocationScore, targetCol) = .Cells(eRowInput.rowAllocateScore, eColInput.colDataStart + i - 1).value
            Sh_data.Cells(eRowData.rowClippingSup, targetCol) = .Cells(eRowInput.rowClippingSup, eColInput.colDataStart + i - 1).value
            Sh_data.Cells(eRowData.rowClippingInf, targetCol) = .Cells(eRowInput.rowClippingInf, eColInput.colDataStart + i - 1).value
            Sh_data.Cells(eRowData.rowConvScore, targetCol) = .Cells(eRowInput.rowConvScore, eColInput.colDataStart + i - 1).value
            Sh_data.Cells(eRowData.rowAdjScoreSup, targetCol) = .Cells(eRowInput.rowAdjScoreSup, eColInput.colDataStart + i - 1).value
            Sh_data.Cells(eRowData.rowAdjScoreInf, targetCol) = .Cells(eRowInput.rowAdjScoreInf, eColInput.colDataStart + i - 1).value

            ' 調整後配点の数式を設定
            Sh_data.Cells(eRowData.rowAdjAllocateScore, targetCol).formula = _
                "=調整後配点計算(" & colLetter & eRowData.rowAllocationScore & "," & _
                colLetter & eRowData.rowClippingSup & "," & _
                colLetter & eRowData.rowConvScore & "," & _
                colLetter & eRowData.rowAdjScoreSup & ")"

            ' 重み（空欄の場合は1）
            weightValue = .Cells(eRowInput.rowWeight, eColInput.colDataStart + i - 1).value
            If Trim(weightValue & "") = "" Then
                Sh_data.Cells(eRowData.rowWeight, targetCol) = 1
            Else
                Sh_data.Cells(eRowData.rowWeight, targetCol) = weightValue
            End If

            ' 統計計算式の設定（"N"追試中マーカーと"-"免除を除外）
            Dim rng As String
            rng = colLetter & eRowData.rowChildStart & ":" & colLetter & lastRowData

            Sh_data.Cells(eRowData.rowAverage, targetCol).formula = _
                "=IFERROR(AVERAGEIFS(" & rng & "," & rng & ",""<>N""," & rng & ",""<>-""),"""")"
            Sh_data.Cells(eRowData.rowMedian, targetCol).FormulaArray = _
                "=IFERROR(MEDIAN(IF((" & rng & "<>""N"")*(" & rng & "<>""-"")*(" & rng & "<>"""")," & rng & ")),"""")"
            Sh_data.Cells(eRowData.rowStdDev, targetCol).FormulaArray = _
                "=IFERROR(STDEV.P(IF((" & rng & "<>""N"")*(" & rng & "<>""-"")*(" & rng & "<>"""")," & rng & ")),"""")"
            Sh_data.Cells(eRowData.rowCV, targetCol).formula = _
                "=IFERROR(" & colLetter & eRowData.rowStdDev & "/" & colLetter & eRowData.rowAverage & ","""")"

            ' 児童の得点データ転記
            For j = 1 To lastRow - eRowInput.rowChildStart + 1
                Sh_data.Cells(eRowData.rowChildStart + j - 1, targetCol) = _
                    .Cells(eRowInput.rowChildStart + j - 1, eColInput.colDataStart + i - 1).value
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
    lastRow = sh_input.Cells(Rows.count, 2).End(xlUp).Row

    With sh_input
        .Range(RNG_INPUT_SUBJECT).ClearContents
        .Range(RNG_INPUT_CATEGORY).ClearContents
        .Range(RNG_INPUT_TEST_NAME).MergeArea.ClearContents
        .Range(RNG_INPUT_TEST_REMARK).ClearContents

        ' 評価観点～調整範囲をクリア
        .Range(.Cells(eRowInput.rowPerspective, eColInput.colDataStart), _
               .Cells(eRowInput.rowAdjScoreInf, eColInput.colDataEnd)).ClearContents

        ' 重みをクリア
        .Range(.Cells(eRowInput.rowWeight, eColInput.colDataStart), _
               .Cells(eRowInput.rowWeight, eColInput.colDataEnd)).ClearContents

        ' 児童の得点データをクリア
        .Range(.Cells(eRowInput.rowChildStart, eColInput.colDataStart), _
               .Cells(lastRow, eColInput.colDataEnd)).ClearContents
    End With

    ' 追試関連のフィールドクリア（行28の追試有無セル）
    sh_input.Range(sh_input.Cells(ROW_INPUT_RETEST, eColInput.colDataStart), _
                   sh_input.Cells(ROW_INPUT_RETEST, eColInput.colDataEnd)).ClearContents

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
        Do While Trim(.Cells(i, SETTING_SUBJECT_COL).value & "") <> ""
            If .Cells(i, SETTING_SUBJECT_COL).value = subjectName Then
                keyChar = .Cells(i, SETTING_KEY_CHAR_COL).value
                keyCount = val(.Cells(i, SETTING_KEY_COUNT_COL).value) + 1

                ' カウンタを更新
                .Cells(i, SETTING_KEY_COUNT_COL).value = keyCount

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

'===============================================================================
' 指定列番号（1始まり）の追試が有効かどうか判定
' 引数: testIndex - テストのインデックス（1～5）
' 説明: 行28の該当列セルが "あり" なら追試有効と判定
'===============================================================================
Public Function IsRetestEnabledForColumn(ByVal testIndex As Long) As Boolean
    Dim cellValue As String
    cellValue = Trim(sh_input.Cells(ROW_INPUT_RETEST, eColInput.colDataStart + testIndex - 1).value & "")
    IsRetestEnabledForColumn = (cellValue = RETEST_ENABLED_VALUE)
End Function

'===============================================================================
' データシートの得点セルを追試中マーカー "N" で上書き
' 説明：追試付きのテスト登録時、追試ONの列の全児童の得点を "N" に置き換え
'       TransferData で一旦実データを書き込んだ後に呼び出す
' 引数:
'   numTest - 今回登録したテスト数（列数）
'   lastRowData - データシートの児童データ最終行
'   lastColData - TransferData呼び出し前のデータシート最終列
'   retestFlags() - 列ごとの追試フラグ（True=追試あり）
'===============================================================================
Private Sub MarkAsRetestPending(ByVal numTest As Long, ByVal lastRowData As Long, _
                                 ByVal lastColData As Long, ByRef retestFlags() As Boolean)
    Dim i As Long, j As Long
    Dim targetCol As Long

    With Sh_data
        For i = 1 To numTest
            If retestFlags(i) Then
                targetCol = lastColData + i
                For j = eRowData.rowChildStart To lastRowData
                    ' 空欄や "-"（免除）はそのまま残す
                    If Trim(.Cells(j, targetCol).value & "") <> "" And _
                       .Cells(j, targetCol).value <> "-" Then
                        .Cells(j, targetCol).value = RETEST_MARKER
                    End If
                Next j
                ' 追試中列にオレンジのフォーマットを適用
                Call UIFormatModule.ApplyRetestColumnFormat(targetCol)
            End If
        Next i
    End With
End Sub
