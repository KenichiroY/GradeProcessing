'===============================================================================
' モジュール名: ScoreCalculationModule
' 説明: 得点調整・変換の計算機能を提供
' 修正内容:
'   - 関数名を英語に統一（日本語関数名は互換性のため残す）
'   - 型を明示的に指定
'   - エラーハンドリング追加
'   - デバッグコード削除
'===============================================================================
Option Explicit

'===============================================================================
' 調整後配点を計算（ワークシート関数として使用可能）
' 引数:
'   allocateScore - 元の配点
'   clippingSup - クリッピング上限（空欄可）
'   conversionType - 変換方式（"平方根"/"対数"/空欄）
'   adjustedScore - 調整後配点（明示的に指定された場合）
' 戻り値: 調整後の配点
'===============================================================================
Public Function CalculateAdjustedAllocateScore(ByVal allocateScore As Variant, _
                                               ByVal clippingSup As Variant, _
                                               ByVal conversionType As String, _
                                               ByVal adjustedScore As Variant) As Double
    On Error GoTo ErrorHandler
    
    Dim result As Double
    
    ' 明示的に調整後配点が指定されている場合はそれを使用
    If Trim(adjustedScore & "") <> "" And adjustedScore <> -1 Then
        CalculateAdjustedAllocateScore = CDbl(adjustedScore)
        Exit Function
    End If
    
    ' 基本配点
    result = CDbl(allocateScore)
    
    ' クリッピング上限が指定されている場合
    If Trim(clippingSup & "") <> "" Then
        result = CDbl(clippingSup)
    End If
    
    ' 得点変換
    Select Case conversionType
        Case "平方根"
            If result > 0 Then
                result = Sqr(result)
            End If
        Case "対数"
            If result > 0 Then
                result = Log(result) / Log(2)   ' 底2の対数
            End If
    End Select
    
    CalculateAdjustedAllocateScore = Round(result, 2)
    Exit Function
    
ErrorHandler:
    CalculateAdjustedAllocateScore = 0
End Function

'===============================================================================
' 調整後得点を計算
' 引数:
'   score - 元の得点
'   allocateScore - 配点
'   clippingSup - クリッピング上限（オプション、デフォルト=-1で配点を使用）
'   clippingInf - クリッピング下限（オプション、デフォルト=-1）
'   conversionType - 変換方式（オプション、デフォルト="ID"）
'   adjustedSup - 調整後範囲上限（オプション、デフォルト=-1）
'   adjustedInf - 調整後範囲下限（オプション、デフォルト=0）
' 戻り値: 調整後の得点
'===============================================================================
Public Function CalculateAdjustedScore(ByVal score As Double, _
                                       ByVal allocateScore As Double, _
                                       Optional ByVal clippingSup As Double = -1, _
                                       Optional ByVal clippingInf As Double = -1, _
                                       Optional ByVal conversionType As String = "ID", _
                                       Optional ByVal adjustedSup As Double = -1, _
                                       Optional ByVal adjustedInf As Double = 0) As Double
    On Error GoTo ErrorHandler
    
    Dim clippedScore As Double
    Dim convertedScore As Double
    Dim convertedAllocate As Double
    Dim finalAllocate As Double
    
    ' クリッピング上限のデフォルト設定
    If clippingSup = -1 Then
        clippingSup = allocateScore
    End If
    
    ' クリッピング処理
    clippedScore = Application.WorksheetFunction.Max( _
                   Application.WorksheetFunction.Min(clippingSup, score), _
                   IIf(clippingInf = -1, 0, clippingInf))
    
    ' 得点変換
    convertedScore = ConvertScore(clippedScore, conversionType)
    
    ' 調整後配点の計算
    convertedAllocate = CalculateAdjustedAllocateScore(allocateScore, clippingSup, conversionType, "")
    finalAllocate = CalculateAdjustedAllocateScore(allocateScore, clippingSup, conversionType, adjustedSup)
    
    ' 範囲調整
    If convertedAllocate <> 0 Then
        CalculateAdjustedScore = (convertedScore / convertedAllocate) * (finalAllocate - adjustedInf) + adjustedInf
    Else
        CalculateAdjustedScore = 0
    End If
    
    Exit Function
    
ErrorHandler:
    CalculateAdjustedScore = 0
End Function

'===============================================================================
' Subjectシート用の調整後得点計算
' 引数:
'   score - 元の得点
'   colIndex - Subjectシートの列番号
' 戻り値: 調整後の得点
'===============================================================================
Public Function CalculateAdjustedScoreForSubject(ByVal score As Double, _
                                                  ByVal colIndex As Long) As Double
    On Error GoTo ErrorHandler
    
    Dim allocateScore As Double
    Dim clippingSup As Double
    Dim clippingInf As Double
    Dim conversionType As String
    Dim adjustedSup As Double
    Dim adjustedInf As Double
    Dim clippedConvertedAllocate As Double
    Dim clippedScore As Double
    Dim convertedScore As Double
    
    With sh_subject
        ' 配点
        allocateScore = CDbl(.Cells(eRowSubject.rowAllocationScore, colIndex).Value)
        
        ' クリッピング上限
        If Trim(.Cells(eRowSubject.rowClippingSup, colIndex).Value & "") = "" Then
            clippingSup = allocateScore
        Else
            clippingSup = CDbl(.Cells(eRowSubject.rowClippingSup, colIndex).Value)
        End If
        
        ' クリッピング下限
        If Trim(.Cells(eRowSubject.rowClippingInf, colIndex).Value & "") = "" Then
            clippingInf = 0
        Else
            clippingInf = CDbl(.Cells(eRowSubject.rowClippingInf, colIndex).Value)
        End If
        
        ' 変換方式
        If Trim(.Cells(eRowSubject.rowConvScore, colIndex).Value & "") = "" Then
            conversionType = "ID"
        Else
            conversionType = CStr(.Cells(eRowSubject.rowConvScore, colIndex).Value)
        End If
        
        ' 調整後範囲上限
        If Trim(.Cells(eRowSubject.rowAdjScoreSup, colIndex).Value & "") = "" Then
            adjustedSup = -1
        Else
            adjustedSup = CDbl(.Cells(eRowSubject.rowAdjScoreSup, colIndex).Value)
        End If
        
        ' 調整後範囲下限
        If Trim(.Cells(eRowSubject.rowAdjScoreInf, colIndex).Value & "") = "" Then
            adjustedInf = 0
        Else
            adjustedInf = CDbl(.Cells(eRowSubject.rowAdjScoreInf, colIndex).Value)
        End If
    End With
    
    ' クリッピング処理
    clippedScore = Application.WorksheetFunction.Max( _
                   Application.WorksheetFunction.Min(clippingSup, score), clippingInf)
    
    ' 得点変換
    convertedScore = ConvertScore(clippedScore, conversionType)
    
    ' クリップ・変換後の配点
    clippedConvertedAllocate = ConvertScore( _
        Application.WorksheetFunction.Max( _
            Application.WorksheetFunction.Min(clippingSup, allocateScore), clippingInf), _
        conversionType)
    
    ' 範囲調整
    If clippedConvertedAllocate <> 0 Then
        CalculateAdjustedScoreForSubject = (convertedScore / clippedConvertedAllocate) * _
            (CalculateAdjustedAllocateScore(allocateScore, clippingSup, conversionType, adjustedSup) - adjustedInf) + adjustedInf
    Else
        CalculateAdjustedScoreForSubject = 0
    End If
    
    Exit Function
    
ErrorHandler:
    CalculateAdjustedScoreForSubject = 0
End Function

'===============================================================================
' 得点変換（内部関数）
' 引数:
'   score - 変換前の得点
'   conversionType - 変換方式
' 戻り値: 変換後の得点
'===============================================================================
Private Function ConvertScore(ByVal score As Double, ByVal conversionType As String) As Double
    On Error GoTo ErrorHandler
    
    Select Case conversionType
        Case "平方根"
            If score >= 0 Then
                ConvertScore = Sqr(score)
            Else
                ConvertScore = 0
            End If
        Case "対数"
            If score > 0 Then
                ConvertScore = Log(score) / Log(2)
            Else
                ConvertScore = 0
            End If
        Case Else   ' "ID" または空欄
            ConvertScore = score
    End Select
    
    Exit Function
    
ErrorHandler:
    ConvertScore = score
End Function

'===============================================================================
' 注意: 日本語関数名（調整後配点計算、調整後得点計算、得点変換 等）は
'       Module1.bas に定義されています。
'       ワークシート数式から呼び出されるため、Module1 の関数を使用してください。
'       このモジュールの英語関数名は VBA コード内から呼び出す用途です。
'===============================================================================
