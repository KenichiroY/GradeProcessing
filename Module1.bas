'===============================================================================
' モジュール名: Module1
' 説明: 得点調整・変換の計算機能を提供
' 注意: このモジュール名は変更しないでください。
'       ワークシートの数式がこのモジュールの関数を参照しています。
'===============================================================================
Option Explicit

'===============================================================================
' 調整後配点を計算（ワークシート関数として使用）
' 引数:
'   alsc - 元の配点（allocation score）
'   clip - クリッピング上限
'   conv - 変換方式（"平方根"/"対数"/空欄）
'   adjust - 調整後配点（明示的に指定された場合）
' 戻り値: 調整後の配点
'===============================================================================
Public Function 調整後配点計算(alsc As Variant, clip As Variant, conv As String, adjust As Variant) As Double
    On Error GoTo ErrorHandler
    
    Dim result As Double
    
    ' 明示的に調整後配点が指定されている場合はそれを使用
    If Trim(adjust & "") <> "" And adjust <> -1 Then
        調整後配点計算 = CDbl(adjust)
        Exit Function
    End If
    
    ' 基本配点
    result = CDbl(alsc)
    
    ' クリッピング上限が指定されている場合
    If Trim(clip & "") <> "" Then
        result = CDbl(clip)
    End If
    
    ' 得点変換
    Select Case conv
        Case "平方根"
            If result > 0 Then
                result = Sqr(result)
            End If
        Case "対数"
            If result > 0 Then
                result = Log(result) / Log(2)   ' 底2の対数
            End If
    End Select
    
    調整後配点計算 = Round(result, 2)
    Exit Function
    
ErrorHandler:
    調整後配点計算 = 0
End Function

'===============================================================================
' 調整後得点を計算
' 引数:
'   sc - 元の得点
'   alsc - 配点
'   clip_sup - クリッピング上限（オプション）
'   clip_inf - クリッピング下限（オプション）
'   conv - 変換方式（オプション）
'   adj_sup - 調整後範囲上限（オプション）
'   adj_inf - 調整後範囲下限（オプション）
' 戻り値: 調整後の得点
'===============================================================================
Public Function 調整後得点計算(ByVal sc As Double, ByVal alsc As Double, _
                               Optional ByVal clip_sup As Double = -1, _
                               Optional ByVal clip_inf As Double = -1, _
                               Optional ByVal conv As String = "ID", _
                               Optional ByVal adj_sup As Double = -1, _
                               Optional ByVal adj_inf As Double = 0) As Double
    On Error GoTo ErrorHandler
    
    Dim clippedScore As Double
    Dim convertedScore As Double
    Dim convertedAllocate As Double
    Dim finalAllocate As Double
    
    ' クリッピング上限のデフォルト設定
    If clip_sup = -1 Then
        clip_sup = alsc
    End If
    
    ' クリッピング下限のデフォルト設定
    If clip_inf = -1 Then
        clip_inf = 0
    End If
    
    ' クリッピング処理
    clippedScore = Application.WorksheetFunction.Max( _
                   Application.WorksheetFunction.Min(clip_sup, sc), clip_inf)
    
    ' 得点変換
    convertedScore = 得点変換(clippedScore, conv)
    
    ' 調整後配点の計算
    convertedAllocate = 調整後配点計算(alsc, clip_sup, conv, "")
    finalAllocate = 調整後配点計算(alsc, clip_sup, conv, adj_sup)
    
    ' 範囲調整
    If convertedAllocate <> 0 Then
        調整後得点計算 = (convertedScore / convertedAllocate) * (finalAllocate - adj_inf) + adj_inf
    Else
        調整後得点計算 = 0
    End If
    
    Exit Function
    
ErrorHandler:
    調整後得点計算 = 0
End Function

'===============================================================================
' Subjectシート用の調整後得点計算
'===============================================================================
Public Function 調整後得点計算_shsubject(ByVal sc As Double, ByVal col As Long) As Double
    On Error GoTo ErrorHandler
    
    Dim alsc As Double
    Dim clip_sup As Double
    Dim clip_inf As Double
    Dim conv As String
    Dim adj_sup As Double
    Dim adj_inf As Double
    Dim clippedConvertedAllocate As Double
    Dim clippedScore As Double
    Dim convertedScore As Double
    
    With sh_subject
        ' 配点
        alsc = CDbl(.Cells(eRowSubject.rowAllocationScore, col).Value)
        
        ' クリッピング上限
        If Trim(.Cells(eRowSubject.rowClippingSup, col).Value & "") = "" Then
            clip_sup = alsc
        Else
            clip_sup = CDbl(.Cells(eRowSubject.rowClippingSup, col).Value)
        End If
        
        ' クリッピング下限
        If Trim(.Cells(eRowSubject.rowClippingInf, col).Value & "") = "" Then
            clip_inf = 0
        Else
            clip_inf = CDbl(.Cells(eRowSubject.rowClippingInf, col).Value)
        End If
        
        ' 変換方式
        If Trim(.Cells(eRowSubject.rowConvScore, col).Value & "") = "" Then
            conv = "ID"
        Else
            conv = CStr(.Cells(eRowSubject.rowConvScore, col).Value)
        End If
        
        ' 調整後範囲上限
        If Trim(.Cells(eRowSubject.rowAdjScoreSup, col).Value & "") = "" Then
            adj_sup = -1
        Else
            adj_sup = CDbl(.Cells(eRowSubject.rowAdjScoreSup, col).Value)
        End If
        
        ' 調整後範囲下限
        If Trim(.Cells(eRowSubject.rowAdjScoreInf, col).Value & "") = "" Then
            adj_inf = 0
        Else
            adj_inf = CDbl(.Cells(eRowSubject.rowAdjScoreInf, col).Value)
        End If
    End With
    
    ' クリッピング処理
    clippedScore = Application.WorksheetFunction.Max( _
                   Application.WorksheetFunction.Min(clip_sup, sc), clip_inf)
    
    ' 得点変換
    convertedScore = 得点変換(clippedScore, conv)
    
    ' クリップ・変換後の配点
    clippedConvertedAllocate = 得点変換( _
        Application.WorksheetFunction.Max( _
            Application.WorksheetFunction.Min(clip_sup, alsc), clip_inf), conv)
    
    ' 範囲調整
    If clippedConvertedAllocate <> 0 Then
        調整後得点計算_shsubject = (convertedScore / clippedConvertedAllocate) * _
            (調整後配点計算(alsc, clip_sup, conv, adj_sup) - adj_inf) + adj_inf
    Else
        調整後得点計算_shsubject = 0
    End If
    
    Exit Function
    
ErrorHandler:
    調整後得点計算_shsubject = 0
End Function

'===============================================================================
' 得点変換
' 引数:
'   sc - 変換前の得点
'   conv_type - 変換方式
' 戻り値: 変換後の得点
'===============================================================================
Public Function 得点変換(ByVal sc As Double, ByVal conv_type As String) As Double
    On Error GoTo ErrorHandler
    
    Select Case conv_type
        Case "平方根"
            If sc >= 0 Then
                得点変換 = Sqr(sc)
            Else
                得点変換 = 0
            End If
        Case "対数"
            If sc > 0 Then
                得点変換 = Log(sc) / Log(2)
            Else
                得点変換 = 0
            End If
        Case Else   ' "ID" または空欄
            得点変換 = sc
    End Select
    
    Exit Function
    
ErrorHandler:
    得点変換 = sc
End Function
