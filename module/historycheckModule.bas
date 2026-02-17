Attribute VB_Name = "historycheckModule"
'===============================================================================
' モジュール名: HistoryCheckModule
' 説明: 未入力データの検索・管理機能を提供
' 修正内容:
'   - 変数宣言をLong型に統一
'   - 未宣言変数（lastcol）を修正
'   - エラーハンドリング追加
'===============================================================================
Option Explicit

'===============================================================================
' 未入力データの検索
' 説明: 全テストデータをスキャンし、未入力のセルをMENUシートに一覧表示
'===============================================================================
Public Sub SearchNotYetInput()
    On Error GoTo ErrorHandler

    Dim i As Long, j As Long
    Dim lastCol As Long
    Dim childCount As Long
    Dim outputRow As Long
    Dim notYetCount As Long

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' 児童数取得
    childCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).value
    If childCount = 0 Then
        Call ErrorHandlerModule.ShowInfo("名簿に児童が登録されていません。")
        GoTo CleanExit
    End If

    ' MENUシートの既存データをクリア（11行目以降を確実にクリア）
    Call ClearMenuData

    notYetCount = 0
    outputRow = eRowMenu.rowStart

    ' データシートをスキャン
    With Sh_data
        lastCol = .Cells(eRowData.rowKey, Columns.count).End(xlToLeft).Column

        ' データがない場合
        If lastCol < eColData.colDataStart Then
            Call ErrorHandlerModule.ShowInfo("登録されているテストデータがありません。")
            GoTo CleanExit
        End If

        For i = eColData.colDataStart To lastCol
            ' テストキーがある列のみ対象
            If Trim(.Cells(eRowData.rowKey, i).value & "") <> "" Then
                For j = eRowData.rowChildStart To eRowData.rowChildStart + childCount - 1
                    ' 空欄チェック（"-"は免除として有効な入力）
                    If Trim(.Cells(j, i).value & "") = "" Then
                        ' MENUシートに出力
                        sh_MENU.Cells(outputRow, eColMenu.colCode) = .Cells(j, eColData.colCode).value
                        sh_MENU.Cells(outputRow, eColMenu.colLastName) = .Cells(j, eColData.colLastName).value
                        sh_MENU.Cells(outputRow, eColMenu.colFirstName) = .Cells(j, eColData.colFirstName).value
                        sh_MENU.Cells(outputRow, eColMenu.colSubject) = .Cells(eRowData.rowSubject, i).value
                        sh_MENU.Cells(outputRow, eColMenu.colPerspective) = .Cells(eRowData.rowPerspective, i).value
                        sh_MENU.Cells(outputRow, eColMenu.colTestName) = .Cells(eRowData.rowTestName, i).value
                        sh_MENU.Cells(outputRow, eColMenu.colDetail) = .Cells(eRowData.rowDetail, i).value
                        sh_MENU.Cells(outputRow, eColMenu.colAllocateScore) = .Cells(eRowData.rowAllocationScore, i).value
                        sh_MENU.Cells(outputRow, eColMenu.colToRow) = j
                        sh_MENU.Cells(outputRow, eColMenu.colToCol) = i

                        outputRow = outputRow + 1
                        notYetCount = notYetCount + 1
                    End If
                Next j
            End If
        Next i
    End With

    ' MENUシートのデータエリア書式を適用
    Call UIFormatModule.FormatMenuDataArea

    ' MENUシートをアクティブにして先頭セルを選択
    sh_MENU.Activate
    sh_MENU.Cells(eRowMenu.rowStart, eColMenu.colCode).Select

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("HistoryCheckModule", "SearchNotYetInput")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' MENUシートから未入力データを一括転記
' 説明: MENUシートで入力された点数をデータシートに反映
'===============================================================================
Public Sub TransferFromMenu()
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim lastRow As Long
    Dim transferCount As Long
    Dim targetRow As Long
    Dim targetCol As Long
    Dim scoreValue As Variant
    Dim allocateScore As Double

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' データシートの保護を一時解除
    On Error Resume Next
    Sh_data.Unprotect Password:=SHEET_PROTECT_PASSWORD
    On Error GoTo ErrorHandler

    transferCount = 0

    With sh_MENU
        lastRow = .Cells(Rows.count, eColMenu.colCode).End(xlUp).Row

        ' データがある場合のみ転記処理
        If lastRow >= eRowMenu.rowStart Then
            For i = eRowMenu.rowStart To lastRow
            scoreValue = .Cells(i, eColMenu.colScore).value

            ' 点数が入力されている場合のみ転記
            If Trim(scoreValue & "") <> "" Then
                targetRow = .Cells(i, eColMenu.colToRow).value
                targetCol = .Cells(i, eColMenu.colToCol).value
                allocateScore = .Cells(i, eColMenu.colAllocateScore).value

                ' 入力値の検証
                If scoreValue <> "-" Then
                    If Not IsNumeric(scoreValue) Then
                        Call ErrorHandlerModule.ShowValidationError( _
                            "行 " & (i - eRowMenu.rowStart + 1) & " の点数が不正です。" & vbCrLf & _
                            "数値または「-」（免除）を入力してください。")
                        .Cells(i, eColMenu.colScore).Select
                        GoTo CleanExit
                    End If

                    If CDbl(scoreValue) < 0 Then
                        Call ErrorHandlerModule.ShowValidationError( _
                            "行 " & (i - eRowMenu.rowStart + 1) & " の点数が不正です。" & vbCrLf & _
                            "0以上の値を入力してください。")
                        .Cells(i, eColMenu.colScore).Select
                        GoTo CleanExit
                    End If

                    If CDbl(scoreValue) > allocateScore Then
                        Call ErrorHandlerModule.ShowValidationError( _
                            "行 " & (i - eRowMenu.rowStart + 1) & " の点数が配点を超えています。" & vbCrLf & _
                            "点数: " & scoreValue & " / 配点: " & allocateScore)
                        .Cells(i, eColMenu.colScore).Select
                        GoTo CleanExit
                    End If
                End If

                ' データシートに転記
                Sh_data.Cells(targetRow, targetCol).value = scoreValue
                transferCount = transferCount + 1
            End If
            Next i
        End If
    End With

    ' 未入力一覧を更新（転記の有無に関わらず実行）
    Call SearchNotYetInput

CleanExit:
    ' データシートの保護を再設定
    Call DataManagementModule.ProtectScoreCells
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    ' データシートの保護を再設定
    On Error Resume Next
    Call DataManagementModule.ProtectScoreCells
    On Error GoTo 0

    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("HistoryCheckModule", "TransferFromMenu")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 特定のテストの未入力者を検索
' 引数:
'   testKey - テストキー（例: J001）
'===============================================================================
Public Sub SearchNotYetByTest(ByVal testKey As String)
    On Error GoTo ErrorHandler

    Dim i As Long, j As Long
    Dim lastCol As Long
    Dim childCount As Long
    Dim outputRow As Long
    Dim targetCol As Long
    Dim notYetCount As Long

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' 児童数取得
    childCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).value

    ' テストキーで列を検索
    targetCol = 0
    With Sh_data
        lastCol = .Cells(eRowData.rowKey, Columns.count).End(xlToLeft).Column
        For i = eColData.colDataStart To lastCol
            If .Cells(eRowData.rowKey, i).value = testKey Then
                targetCol = i
                Exit For
            End If
        Next i
    End With

    If targetCol = 0 Then
        Call ErrorHandlerModule.ShowValidationError("テストキー「" & testKey & "」が見つかりません。")
        GoTo CleanExit
    End If

    ' MENUシートの既存データをクリア
    Call ClearMenuData

    notYetCount = 0
    outputRow = eRowMenu.rowStart

    ' 該当テストの未入力を検索
    With Sh_data
        For j = eRowData.rowChildStart To eRowData.rowChildStart + childCount - 1
            If Trim(.Cells(j, targetCol).value & "") = "" Then
                sh_MENU.Cells(outputRow, eColMenu.colCode) = .Cells(j, eColData.colCode).value
                sh_MENU.Cells(outputRow, eColMenu.colLastName) = .Cells(j, eColData.colLastName).value
                sh_MENU.Cells(outputRow, eColMenu.colFirstName) = .Cells(j, eColData.colFirstName).value
                sh_MENU.Cells(outputRow, eColMenu.colSubject) = .Cells(eRowData.rowSubject, targetCol).value
                sh_MENU.Cells(outputRow, eColMenu.colPerspective) = .Cells(eRowData.rowPerspective, targetCol).value
                sh_MENU.Cells(outputRow, eColMenu.colTestName) = .Cells(eRowData.rowTestName, targetCol).value
                sh_MENU.Cells(outputRow, eColMenu.colDetail) = .Cells(eRowData.rowDetail, targetCol).value
                sh_MENU.Cells(outputRow, eColMenu.colAllocateScore) = .Cells(eRowData.rowAllocationScore, targetCol).value
                sh_MENU.Cells(outputRow, eColMenu.colToRow) = j
                sh_MENU.Cells(outputRow, eColMenu.colToCol) = targetCol

                outputRow = outputRow + 1
                notYetCount = notYetCount + 1
            End If
        Next j
    End With

    ' MENUシートのデータエリア書式を適用
    Call UIFormatModule.FormatMenuDataArea

    ' 結果メッセージ
    If notYetCount = 0 Then
        Call ErrorHandlerModule.ShowSuccess("テスト「" & testKey & "」に未入力はありません。")
    Else
        Call ErrorHandlerModule.ShowInfo("未入力データが " & notYetCount & " 件見つかりました。")
        sh_MENU.Activate
        sh_MENU.Cells(eRowMenu.rowStart, eColMenu.colCode).Select
    End If

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("HistoryCheckModule", "SearchNotYetByTest")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' MENUシートのデータ領域をクリア
' 説明: 11行目以降のデータを確実にクリアする
'===============================================================================
Private Sub ClearMenuData()
    Dim lastRow As Long
    Dim clearEndRow As Long

    With sh_MENU
        ' 各列の最終行を確認して最大値を取得
        lastRow = eRowMenu.rowStart

        ' B列（コード）の最終行
        If .Cells(Rows.count, eColMenu.colCode).End(xlUp).Row > lastRow Then
            lastRow = .Cells(Rows.count, eColMenu.colCode).End(xlUp).Row
        End If

        ' I列（点数）の最終行
        If .Cells(Rows.count, eColMenu.colScore).End(xlUp).Row > lastRow Then
            lastRow = .Cells(Rows.count, eColMenu.colScore).End(xlUp).Row
        End If

        ' L列（転記先列）の最終行
        If .Cells(Rows.count, eColMenu.colToCol).End(xlUp).Row > lastRow Then
            lastRow = .Cells(Rows.count, eColMenu.colToCol).End(xlUp).Row
        End If

        ' クリア対象がある場合のみクリア
        If lastRow >= eRowMenu.rowStart Then
            ' 余裕を持って少し多めにクリア
            clearEndRow = lastRow + 10
            Dim clearRange As Range
            Set clearRange = .Range(.Cells(eRowMenu.rowStart, eColMenu.colCode), _
                                     .Cells(clearEndRow, eColMenu.colToCol))
            clearRange.ClearContents
            ' 書式もクリア（罫線・背景色）
            clearRange.Interior.ColorIndex = xlColorIndexNone
            clearRange.Borders.LineStyle = xlLineStyleNone
        End If
    End With
End Sub


