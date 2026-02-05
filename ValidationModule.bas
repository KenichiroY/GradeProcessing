'===============================================================================
' モジュール名: ValidationModule
' 説明: 入力データの検証機能を提供
' 目的: 不正なデータの登録を防止し、エラーを未然に防ぐ
' 修正: 全てのExit Function前に戻り値を設定するよう修正
'===============================================================================
Option Explicit

'===============================================================================
' 検証結果を格納する構造体
'===============================================================================
Public Type ValidationResult
    IsValid As Boolean
    ErrorMessage As String
    ErrorRow As Long
    ErrorCol As Long
End Type

'===============================================================================
' テスト入力シートの必須項目をチェック
'===============================================================================
Public Function ValidateRequiredFields() As ValidationResult
    Dim result As ValidationResult
    result.IsValid = True
    
    With sh_input
        ' 教科チェック
        If Trim(.Range(RNG_INPUT_SUBJECT).Value & "") = "" Then
            result.IsValid = False
            result.ErrorMessage = "「教科」が選択されていません。"
            ValidateRequiredFields = result
            Exit Function
        End If
        
        ' カテゴリチェック
        If Trim(.Range(RNG_INPUT_CATEGORY).Value & "") = "" Then
            result.IsValid = False
            result.ErrorMessage = "「カテゴリ」が選択されていません。"
            ValidateRequiredFields = result
            Exit Function
        End If
        
        ' 実施日チェック
        If Trim(.Range(RNG_INPUT_DATE).Value & "") = "" Then
            result.IsValid = False
            result.ErrorMessage = "「実施日」が入力されていません。"
            ValidateRequiredFields = result
            Exit Function
        End If
        
        ' 日付形式チェック
        If Not IsDate(.Range(RNG_INPUT_DATE).Value) Then
            result.IsValid = False
            result.ErrorMessage = "「実施日」の形式が正しくありません。" & vbCrLf & _
                                  "例：2024/04/01 の形式で入力してください。"
            ValidateRequiredFields = result
            Exit Function
        End If
        
        ' テスト名チェック
        If Trim(.Range(RNG_INPUT_TEST_NAME).Value & "") = "" Then
            result.IsValid = False
            result.ErrorMessage = "「テスト名」が入力されていません。"
            ValidateRequiredFields = result
            Exit Function
        End If
    End With
    
    ValidateRequiredFields = result
End Function

'===============================================================================
' 得点データの検証
' 引数:
'   colIndex - チェックする列番号（eColInput.colDataStart基準）
'   lastRow - データの最終行
'===============================================================================
Public Function ValidateScoreData(ByVal colIndex As Long, ByVal lastRow As Long) As ValidationResult
    Dim result As ValidationResult
    Dim i As Long
    Dim scoreValue As Variant
    Dim allocateScore As Variant
    
    result.IsValid = True
    
    With sh_input
        ' 配点チェック
        allocateScore = .Cells(eRowInput.rowAllocateScore, colIndex).Value
        
        ' 配点が空でないかチェック
        If Trim(allocateScore & "") = "" Then
            result.IsValid = False
            result.ErrorMessage = (colIndex - eColInput.colDataStart + 1) & ERR_MSG_MISSING_ALLOCATE
            result.ErrorCol = colIndex
            ValidateScoreData = result
            Exit Function
        End If
        
        ' 配点が数値かチェック
        If Not IsNumeric(allocateScore) Then
            result.IsValid = False
            result.ErrorMessage = "配点には数値を入力してください。" & vbCrLf & _
                                  "列: " & (colIndex - eColInput.colDataStart + 1)
            result.ErrorCol = colIndex
            ValidateScoreData = result
            Exit Function
        End If
        
        ' 配点が0でないかチェック
        If CDbl(allocateScore) = 0 Then
            result.IsValid = False
            result.ErrorMessage = ERR_MSG_ZERO_ALLOCATE
            result.ErrorCol = colIndex
            ValidateScoreData = result
            Exit Function
        End If
        
        ' 配点が負でないかチェック
        If CDbl(allocateScore) < 0 Then
            result.IsValid = False
            result.ErrorMessage = "配点に負の値は設定できません。" & vbCrLf & _
                                  "列: " & (colIndex - eColInput.colDataStart + 1)
            result.ErrorCol = colIndex
            ValidateScoreData = result
            Exit Function
        End If
        
        ' 評価観点チェック
        If Trim(.Cells(eRowInput.rowPerspective, colIndex).Value & "") = "" Then
            result.IsValid = False
            result.ErrorMessage = (colIndex - eColInput.colDataStart + 1) & ERR_MSG_MISSING_PERSPECTIVE
            result.ErrorCol = colIndex
            ValidateScoreData = result
            Exit Function
        End If
        
        ' 各児童の得点をチェック
        For i = eRowInput.rowChildStart To lastRow
            scoreValue = .Cells(i, colIndex).Value
            
            ' 空欄または "-"（欠席）はスキップ
            If Trim(scoreValue & "") = "" Or scoreValue = "-" Then
                GoTo NextRow
            End If
            
            ' 数値チェック
            If Not IsNumeric(scoreValue) Then
                result.IsValid = False
                result.ErrorMessage = "得点には数値を入力してください。" & vbCrLf & _
                                      "行: " & (i - eRowInput.rowChildStart + 1) & "人目" & vbCrLf & _
                                      "列: " & (colIndex - eColInput.colDataStart + 1) & vbCrLf & _
                                      "入力値: " & scoreValue
                result.ErrorRow = i
                result.ErrorCol = colIndex
                ValidateScoreData = result
                Exit Function
            End If
            
            ' 負の値チェック
            If CDbl(scoreValue) < 0 Then
                result.IsValid = False
                result.ErrorMessage = "得点に負の値は入力できません。" & vbCrLf & _
                                      "行: " & (i - eRowInput.rowChildStart + 1) & "人目" & vbCrLf & _
                                      "列: " & (colIndex - eColInput.colDataStart + 1) & vbCrLf & _
                                      "入力値: " & scoreValue
                result.ErrorRow = i
                result.ErrorCol = colIndex
                ValidateScoreData = result
                Exit Function
            End If
            
            ' 得点が配点を超えていないかチェック
            If CDbl(scoreValue) > CDbl(allocateScore) Then
                result.IsValid = False
                result.ErrorMessage = "得点が配点を超えています。" & vbCrLf & _
                                      "行: " & (i - eRowInput.rowChildStart + 1) & "人目" & vbCrLf & _
                                      "列: " & (colIndex - eColInput.colDataStart + 1) & vbCrLf & _
                                      "得点: " & scoreValue & " / 配点: " & allocateScore
                result.ErrorRow = i
                result.ErrorCol = colIndex
                ValidateScoreData = result
                Exit Function
            End If
NextRow:
        Next i
    End With
    
    ValidateScoreData = result
End Function

'===============================================================================
' クリッピング設定の検証
'===============================================================================
Public Function ValidateClippingSettings(ByVal colIndex As Long) As ValidationResult
    Dim result As ValidationResult
    Dim clipSup As Variant, clipInf As Variant
    Dim allocateScore As Variant
    
    result.IsValid = True
    
    With sh_input
        allocateScore = .Cells(eRowInput.rowAllocateScore, colIndex).Value
        clipSup = .Cells(eRowInput.rowClippingSup, colIndex).Value
        clipInf = .Cells(eRowInput.rowClippingInf, colIndex).Value
        
        ' 上限値のチェック
        If Trim(clipSup & "") <> "" Then
            If Not IsNumeric(clipSup) Then
                result.IsValid = False
                result.ErrorMessage = "クリッピング上限には数値を入力してください。"
                ValidateClippingSettings = result
                Exit Function
            End If
            If CDbl(clipSup) > CDbl(allocateScore) Then
                result.IsValid = False
                result.ErrorMessage = "クリッピング上限が配点を超えています。" & vbCrLf & _
                                      "上限: " & clipSup & " / 配点: " & allocateScore
                ValidateClippingSettings = result
                Exit Function
            End If
        End If
        
        ' 下限値のチェック
        If Trim(clipInf & "") <> "" Then
            If Not IsNumeric(clipInf) Then
                result.IsValid = False
                result.ErrorMessage = "クリッピング下限には数値を入力してください。"
                ValidateClippingSettings = result
                Exit Function
            End If
            If CDbl(clipInf) < 0 Then
                result.IsValid = False
                result.ErrorMessage = "クリッピング下限に負の値は設定できません。"
                ValidateClippingSettings = result
                Exit Function
            End If
        End If
        
        ' 上限 > 下限のチェック
        If Trim(clipSup & "") <> "" And Trim(clipInf & "") <> "" Then
            If CDbl(clipSup) < CDbl(clipInf) Then
                result.IsValid = False
                result.ErrorMessage = "クリッピング上限が下限より小さくなっています。" & vbCrLf & _
                                      "上限: " & clipSup & " / 下限: " & clipInf
                ValidateClippingSettings = result
                Exit Function
            End If
        End If
    End With
    
    ValidateClippingSettings = result
End Function

'===============================================================================
' 重み設定の検証
'===============================================================================
Public Function ValidateWeight(ByVal colIndex As Long) As ValidationResult
    Dim result As ValidationResult
    Dim weightValue As Variant
    
    result.IsValid = True
    
    With sh_input
        weightValue = .Cells(eRowInput.rowWeight, colIndex).Value
        
        ' 空欄は1として扱うのでOK
        If Trim(weightValue & "") = "" Then
            ValidateWeight = result
            Exit Function
        End If
        
        ' 数値チェック
        If Not IsNumeric(weightValue) Then
            result.IsValid = False
            result.ErrorMessage = "重みには数値を入力してください。"
            ValidateWeight = result
            Exit Function
        End If
        
        ' 負の値チェック
        If CDbl(weightValue) < 0 Then
            result.IsValid = False
            result.ErrorMessage = "重みに負の値は設定できません。"
            ValidateWeight = result
            Exit Function
        End If
    End With
    
    ValidateWeight = result
End Function

'===============================================================================
' テスト数上限チェック
'===============================================================================
Public Function ValidateTestCountLimit() As ValidationResult
    Dim result As ValidationResult
    Dim currentTestCount As Long
    
    result.IsValid = True
    
    With Sh_data
        currentTestCount = .Cells(eRowData.rowKey, Columns.Count).End(xlToLeft).Column - eColData.colDataStart + 1
        
        If currentTestCount >= MAX_TESTS Then
            result.IsValid = False
            result.ErrorMessage = Replace(ERR_MSG_MAX_TESTS, "{MAX}", CStr(MAX_TESTS))
            ValidateTestCountLimit = result
            Exit Function
        End If
    End With
    
    ValidateTestCountLimit = result
End Function

'===============================================================================
' 教科名が設定に存在するかチェック
'===============================================================================
Public Function ValidateSubjectExists(ByVal subjectName As String) As ValidationResult
    Dim result As ValidationResult
    Dim i As Long
    
    result.IsValid = False
    
    With sh_setting
        For i = SETTING_SUBJECT_START_ROW To SETTING_SUBJECT_START_ROW + 10
            If .Cells(i, SETTING_SUBJECT_COL).Value = subjectName Then
                result.IsValid = True
                ValidateSubjectExists = result
                Exit Function
            End If
            If Trim(.Cells(i, SETTING_SUBJECT_COL).Value & "") = "" Then
                Exit For
            End If
        Next i
    End With
    
    If Not result.IsValid Then
        result.ErrorMessage = "教科「" & subjectName & "」は設定シートに登録されていません。" & vbCrLf & _
                              "Settingシートで教科を登録してください。"
    End If
    
    ValidateSubjectExists = result
End Function

'===============================================================================
' 得点入力があるかチェック（最低1件）
'===============================================================================
Public Function HasAnyScoreInput(ByVal lastRow As Long) As Boolean
    Dim i As Long, j As Long
    
    HasAnyScoreInput = False
    
    With sh_input
        For i = eColInput.colDataStart To eColInput.colDataEnd
            For j = eRowInput.rowChildStart To lastRow
                If Trim(.Cells(j, i).Value & "") <> "" Then
                    HasAnyScoreInput = True
                    Exit Function
                End If
            Next j
        Next i
    End With
End Function
