'===============================================================================
' モジュール名: UIFormatModule
' 説明: 本番で使用されるUI書式関連の処理を提供
'   - FormatMenuDataArea: MENUシートデータエリアの動的書式（HistoryCheckModuleから呼出）
'   - ApplyRetestColumnFormat: 追試中列のオレンジ色表示（PostingModule/RetestModuleから呼出）
'   - ClearRetestColumnFormat: 追試完了時の色クリア（RetestModuleから呼出）
'===============================================================================
Option Explicit

'===============================================================================
' カラー取得関数（ClearRetestColumnFormatでヘッダー行の色復元に使用）
'===============================================================================
Private Function InputBgColor() As Long
    InputBgColor = RGB(218, 232, 247)    ' 淡青（入力欄）
End Function

Private Function SectionBgColor() As Long
    SectionBgColor = RGB(180, 210, 235)  ' やや濃い青（セクション見出し）
End Function

Private Function SubHeaderBgColor() As Long
    SubHeaderBgColor = RGB(141, 180, 226) ' サブヘッダー
End Function

Private Function BorderColor() As Long
    BorderColor = RGB(166, 176, 192)     ' 罫線色（青灰）
End Function

'===============================================================================
' MENUシートのデータエリア書式設定（動的）
' 説明：
'   SearchNotYetInput / SearchNotYetByTest でデータ生成後に呼び出す
'   実際のデータ行数に応じて罫線・入力欄色を設定する
'   ClearMenuData時にも呼ばれ、書式をリセットする
'===============================================================================
Public Sub FormatMenuDataArea()
    Dim ws As Worksheet
    Set ws = sh_MENU

    With ws
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, eColMenu.colCode).End(xlUp).Row

        ' データがない場合は11行目以降の書式をクリアして終了
        If lastRow < eRowMenu.rowStart Then
            Dim clearRange As Range
            Set clearRange = .Range(.Cells(eRowMenu.rowStart, 2), _
                                     .Cells(eRowMenu.rowStart + MAX_CHILDREN, 12))
            clearRange.Interior.ColorIndex = xlColorIndexNone
            clearRange.Borders.LineStyle = xlLineStyleNone
            Exit Sub
        End If

        ' データ範囲
        Dim dataRange As Range
        Set dataRange = .Range(.Cells(eRowMenu.rowStart, eColMenu.colCode), _
                                .Cells(lastRow, eColMenu.colToCol))

        ' 罫線
        Call SetThinBorders(dataRange, BorderColor())

        ' 点数入力列（I列）に入力欄色
        .Range(.Cells(eRowMenu.rowStart, eColMenu.colScore), _
               .Cells(lastRow, eColMenu.colScore)).Interior.Color = InputBgColor()

        ' データ行以降の余白をクリア（前回のデータが多かった場合）
        If lastRow + 1 <= lastRow + MAX_CHILDREN Then
            Dim excessRange As Range
            Set excessRange = .Range(.Cells(lastRow + 1, 2), _
                                      .Cells(lastRow + MAX_CHILDREN, 12))
            excessRange.Interior.ColorIndex = xlColorIndexNone
            excessRange.Borders.LineStyle = xlLineStyleNone
        End If
    End With
End Sub

'===============================================================================
' 罫線設定ヘルパー
' 引数：rng - 対象範囲、clr - 罫線色
'===============================================================================
Private Sub SetThinBorders(ByVal rng As Range, ByVal clr As Long)
    Dim edge As Variant
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
        On Error Resume Next
        With rng.Borders(edge)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = clr
        End With
        On Error GoTo 0
    Next edge
End Sub

'===============================================================================
' 追試中列にオレンジ色のフォーマットを適用
' 説明: 追試中マーカー"N"が入った列のヘッダー行と得点セルを
'       オレンジ系の背景色で強調表示する
' 引数: targetCol - データシートの対象列番号
'===============================================================================
Public Sub ApplyRetestColumnFormat(ByVal targetCol As Long)
    Dim lastRow As Long

    With Sh_data
        lastRow = .Cells(.Rows.Count, eColData.colCode).End(xlUp).Row
        If lastRow < eRowData.rowChildStart Then lastRow = eRowData.rowChildStart

        ' ヘッダー行（4-22行）にオレンジ背景
        .Range(.Cells(eRowData.rowKey, targetCol), _
               .Cells(eRowData.rowCV, targetCol)).Interior.Color = COLOR_RETEST_HEADER

        ' 児童データ行（23行～最終行）に薄オレンジ背景
        .Range(.Cells(eRowData.rowChildStart, targetCol), _
               .Cells(lastRow, targetCol)).Interior.Color = COLOR_RETEST_CELL

        ' "N"セルのフォントを太字・濃いオレンジ色に
        Dim j As Long
        For j = eRowData.rowChildStart To lastRow
            If CStr(.Cells(j, targetCol).Value) = RETEST_MARKER Then
                .Cells(j, targetCol).Font.Bold = True
                .Cells(j, targetCol).Font.Color = RGB(200, 100, 0)  ' 濃いオレンジ
            End If
        Next j
    End With
End Sub

'===============================================================================
' 追試中列のオレンジ色フォーマットをクリアし通常色に戻す
' 説明: 追試完了時に呼び出し、ヘッダー行は行帯ごとの元の色に復元し、
'       得点セルの背景色とフォント装飾をリセットする
' 引数: targetCol - データシートの対象列番号
'===============================================================================
Public Sub ClearRetestColumnFormat(ByVal targetCol As Long)
    Dim lastRow As Long

    With Sh_data
        lastRow = .Cells(.Rows.Count, eColData.colCode).End(xlUp).Row
        If lastRow < eRowData.rowChildStart Then lastRow = eRowData.rowChildStart

        ' ヘッダー行の背景色を行帯ごとに復元
        ' (1) 基本情報（4-10行）→ SubHeaderBgColor
        .Range(.Cells(eRowData.rowKey, targetCol), _
               .Cells(eRowData.rowDetail, targetCol)).Interior.Color = SubHeaderBgColor()

        ' (2) 配点・調整（11-18行）→ SectionBgColor
        .Range(.Cells(eRowData.rowAllocationScore, targetCol), _
               .Cells(eRowData.rowWeight, targetCol)).Interior.Color = SectionBgColor()

        ' (3) 統計値（19-22行）→ InputBgColor
        .Range(.Cells(eRowData.rowAverage, targetCol), _
               .Cells(eRowData.rowCV, targetCol)).Interior.Color = InputBgColor()

        ' 児童データ行の背景色をクリア
        .Range(.Cells(eRowData.rowChildStart, targetCol), _
               .Cells(lastRow, targetCol)).Interior.ColorIndex = xlColorIndexNone

        ' フォント装飾をリセット
        Dim j As Long
        For j = eRowData.rowChildStart To lastRow
            .Cells(j, targetCol).Font.Bold = False
            .Cells(j, targetCol).Font.Color = RGB(0, 0, 0)  ' 黒に戻す
        Next j
    End With
End Sub
