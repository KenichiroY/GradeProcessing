Attribute VB_Name = "RetestModule"
'===============================================================================
' モジュール名: RetestModule
' 説明: 追試機能の全処理を提供
'       追試ファイルの生成・シート作成・結果反映を一元管理する
'===============================================================================
Option Explicit

'===============================================================================
' 追試ファイルのフルパスを取得
' 戻り値：追試ファイルのパス（例："C:\...\成績処理_追試.xlsm"）
'===============================================================================
Public Function GetRetestFilePath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim baseName As String
    Dim parentFolder As String

    baseName = fso.GetBaseName(ThisWorkbook.FullName)
    parentFolder = fso.GetParentFolderName(ThisWorkbook.FullName)

    GetRetestFilePath = parentFolder & "\" & baseName & RETEST_FILE_SUFFIX & RETEST_FILE_EXT

    Set fso = Nothing
End Function

'===============================================================================
' 追試ファイルを取得（なければ新規作成）
' 戻り値：追試ファイルの Workbook オブジェクト
' 説明：
'   - 既に開いていればそのまま返す
'   - ファイルが存在すれば開いて返す
'   - ファイルが存在しなければ新規作成して返す
'   - 新規作成時は MENU シートを初期化する
'===============================================================================
Public Function GetOrCreateRetestWorkbook() As Workbook
    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim wb As Workbook
    Dim fso As Object

    filePath = GetRetestFilePath()
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 既に開いているか確認
    Dim fileName As String
    fileName = fso.GetFileName(filePath)
    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo ErrorHandler

    If Not wb Is Nothing Then
        Set GetOrCreateRetestWorkbook = wb
        Set fso = Nothing
        Exit Function
    End If

    ' ファイルが存在する場合は開く
    If fso.FileExists(filePath) Then
        Set wb = Workbooks.Open(filePath)
        Set GetOrCreateRetestWorkbook = wb
        Set fso = Nothing
        Exit Function
    End If

    ' 新規作成
    Set wb = Workbooks.Add

    ' テンプレートからMENUシートをコピー（CodeNameで検索）
    Dim menuWs As Worksheet
    Set menuWs = CopyTemplateSheet("sh_rt_menu", wb, "MENU")

    ' デフォルトのSheet1を削除
    Application.DisplayAlerts = False
    Dim defaultSheet As Worksheet
    For Each defaultSheet In wb.sheets
        If defaultSheet.Name <> "MENU" Then
            defaultSheet.Delete
        End If
    Next defaultSheet

    ' ボタンのマクロ参照先を本体ファイルに書き換え
    Call AssignButtonMacros(menuWs)

    wb.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True

    Set GetOrCreateRetestWorkbook = wb
    Set fso = Nothing
    Exit Function

ErrorHandler:
    Set GetOrCreateRetestWorkbook = Nothing
    Set fso = Nothing
End Function

'===============================================================================
' テンプレートシートを追試ファイルにコピー
' 引数：
'   templateCodeName - テンプレートシートのCodeName（sh_rt_menu or sh_rt_template）
'   targetWb - コピー先のWorkbook
'   newName - コピー後のシート名
' 戻り値：コピーされたWorksheet
' 説明：
'   CodeNameでシートを検索するため、シートのタブ名に依存しない
'   VeryHiddenシートは直接コピーできないため、一時的にVisibleにしてコピーする
'===============================================================================
Private Function CopyTemplateSheet(ByVal templateCodeName As String, _
                                    ByVal targetWb As Workbook, _
                                    ByVal newName As String) As Worksheet
    Dim templateWs As Worksheet
    Dim newWs As Worksheet

    ' CodeNameでテンプレートシートを検索
    Set templateWs = GetSheetByCodeName(ThisWorkbook, templateCodeName)
    If templateWs Is Nothing Then
        Err.Raise vbObjectError + 1, "CopyTemplateSheet", _
            "テンプレートシート（CodeName: " & templateCodeName & "）が見つかりません。"
    End If

    ' 元の表示状態を保存
    Dim originalVisible As XlSheetVisibility
    originalVisible = templateWs.Visible

    ' 一時的にVisibleにする（VeryHiddenのままではコピーできない）
    templateWs.Visible = xlSheetVisible

    ' コピー先ワークブックにコピー
    templateWs.Copy After:=targetWb.sheets(targetWb.sheets.count)

    ' テンプレートを元の状態に戻す
    templateWs.Visible = originalVisible

    ' コピーされたシート（最後のシート）を取得してリネーム
    Set newWs = targetWb.sheets(targetWb.sheets.count)
    newWs.Name = newName

    Set CopyTemplateSheet = newWs
End Function

'===============================================================================
' CodeNameでシートを検索
' 引数：wb - 検索対象のWorkbook, codeName - CodeName
' 戻り値：見つかったWorksheet（見つからなければNothing）
'===============================================================================
Private Function GetSheetByCodeName(ByVal wb As Workbook, ByVal codeName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.codeName = codeName Then
            Set GetSheetByCodeName = ws
            Exit Function
        End If
    Next ws
    Set GetSheetByCodeName = Nothing
End Function

'===============================================================================
' コピーされたシートのボタンマクロ参照先を書き換え
' 説明：
'   テンプレートのボタンは「RetestModule.XXX」という名前でマクロ割り当てされている
'   コピー後に「'本体ファイル名'!RetestModule.XXX」に書き換える
'   これにより追試ファイルのボタンから本体ファイルのマクロを実行できる
'===============================================================================
Private Sub AssignButtonMacros(ByVal ws As Worksheet)
    On Error Resume Next

    Dim btn As Button
    Dim macroPrefix As String
    Dim originalMacro As String

    ' マクロ参照先プレフィックス: '本体ファイル名'!
    macroPrefix = "'" & ThisWorkbook.Name & "'!"

    For Each btn In ws.Buttons
        originalMacro = btn.OnAction
        ' 既にプレフィックス付きの場合はスキップ
        If Left(originalMacro, 1) <> "'" Then
            btn.OnAction = macroPrefix & originalMacro
        End If
    Next btn

    On Error GoTo 0
End Sub

'===============================================================================
' 追試シートを作成
' 説明：PostingModule.Posting() から呼び出される
'       本体ファイルにテストデータを登録した直後に呼ばれる
' 引数：
'   numTest - 今回登録したテスト数
'   lastRow - テスト入力シートの最終行
'   lastColData - TransferData 呼び出し前のデータシート最終列
'                 （= 新しいテストの列は lastColData + 1 ～ lastColData + numTest）
'   retestFlags() - 列ごとの追試フラグ（True=追試あり）
'===============================================================================
Public Sub CreateRetestSheet(ByVal numTest As Long, ByVal lastRow As Long, _
                              ByVal lastColData As Long, ByRef retestFlags() As Boolean)
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim childCount As Long
    Dim testKey As String
    Dim sheetName As String
    Dim targetCol As Long

    ' 追試ファイルを取得
    Set wb = GetOrCreateRetestWorkbook()
    If wb Is Nothing Then
        Call ErrorHandlerModule.ShowValidationError("追試ファイルの作成に失敗しました。")
        Exit Sub
    End If

    childCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).value

    ' テスト数分のシートを作成（追試ONの列のみ）
    For i = 1 To numTest
        If Not retestFlags(i) Then GoTo NextTest

        targetCol = lastColData + i
        testKey = Sh_data.Cells(eRowData.rowKey, targetCol).value

        ' シート名の生成（キー_テスト名_観点、31文字制限）
        sheetName = testKey & "_" & _
                    Sh_data.Cells(eRowData.rowTestName, targetCol).value & "_" & _
                    Sh_data.Cells(eRowData.rowPerspective, targetCol).value
        If Len(sheetName) > 31 Then
            sheetName = Left(sheetName, 31)
        End If

        ' 同名シートの重複回避
        sheetName = GetUniqueSheetName(wb, sheetName)

        ' テンプレートからシートをコピー（CodeNameで検索）
        Set ws = CopyTemplateSheet("sh_rt_template", wb, sheetName)

        ' ボタンのマクロ参照先を本体ファイルに書き換え
        Call AssignButtonMacros(ws)

        ' テスト情報の値を書き込み（動的データのみ）
        With ws
            .Range(RNG_RT_PARENT_KEY).value = testKey
            .Range(RNG_RT_SUBJECT).value = Sh_data.Cells(eRowData.rowSubject, targetCol).value
            .Range(RNG_RT_TEST_NAME).value = Sh_data.Cells(eRowData.rowTestName, targetCol).value
            .Range(RNG_RT_PERSPECTIVE).value = Sh_data.Cells(eRowData.rowPerspective, targetCol).value
            .Range(RNG_RT_DETAIL).value = Sh_data.Cells(eRowData.rowDetail, targetCol).value
            .Range(RNG_RT_ALLOCATE).value = Sh_data.Cells(eRowData.rowAllocationScore, targetCol).value

            ' 児童データの転記
            For j = 1 To childCount
                .Cells(RT_DATA_START_ROW + j - 1, RT_COL_CODE).value = _
                    Sh_data.Cells(eRowData.rowChildStart + j - 1, eColData.colCode).value
                .Cells(RT_DATA_START_ROW + j - 1, RT_COL_LASTNAME).value = _
                    Sh_data.Cells(eRowData.rowChildStart + j - 1, eColData.colLastName).value
                .Cells(RT_DATA_START_ROW + j - 1, RT_COL_FIRSTNAME).value = _
                    Sh_data.Cells(eRowData.rowChildStart + j - 1, eColData.colFirstName).value

                ' 本試の得点（sh_input から取得）
                Dim originalScore As Variant
                originalScore = sh_input.Cells(eRowInput.rowChildStart + j - 1, _
                                eColInput.colDataStart + i - 1).value
                .Cells(RT_DATA_START_ROW + j - 1, RT_COL_ORIGINAL).value = originalScore
            Next j
        End With

        ' 合格者数・未合格者数の数式を設定（合格点入力時に自動計算）
        Dim initFinalCol As Long
        initFinalCol = RT_COL_RETEST_START + RT_COL_FINAL_OFFSET
        Call SetPassFailCountFormulas(ws, initFinalCol, childCount)

        ' MENUシートに追加
        Call AddToRetestMenu(wb, testKey, _
            Sh_data.Cells(eRowData.rowSubject, targetCol).value, _
            Sh_data.Cells(eRowData.rowTestName, targetCol).value, _
            Sh_data.Cells(eRowData.rowPerspective, targetCol).value, _
            sheetName)
NextTest:
    Next i

    wb.Save

    Call ErrorHandlerModule.ShowSuccess( _
        "追試シートを作成しました。" & vbCrLf & _
        "追試ファイル: " & wb.Name)

    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "CreateRetestSheet")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 最終得点列に算出方法に応じた数式を設定
' 引数：
'   ws - 追試シート
'   dataRow - 対象行
'   originalCol - 本試列
'   retestStartCol - 追試開始列
'   finalCol - 最終列
'   method - 算出方法
' 説明：
'   追試列が追加されるたびにこの関数を再呼出しして数式を更新する
'   免除 "-" の児童はそのまま "-" を返す
'   追試を受けていない児童は本試の値をそのまま使う
'===============================================================================
Public Sub SetFinalScoreFormula(ByVal ws As Worksheet, ByVal dataRow As Long, _
                                 ByVal originalCol As Long, ByVal retestStartCol As Long, _
                                 ByVal finalCol As Long, ByVal method As String)
    Dim colLetterOrig As String
    Dim colLetterRetestStart As String
    Dim colLetterRetestEnd As String
    Dim paramValue As Double
    Dim lastRetestCol As Long
    Dim formula As String

    colLetterOrig = PostingModule.ColumnIndexToLetter(originalCol)
    colLetterRetestStart = PostingModule.ColumnIndexToLetter(retestStartCol)

    ' 最後の追試列 = finalCol - 1
    lastRetestCol = finalCol - 1
    colLetterRetestEnd = PostingModule.ColumnIndexToLetter(lastRetestCol)

    ' 共通の変数定義
    Dim origCell As String
    Dim retestRange As String
    Dim hasRetest As String
    origCell = colLetterOrig & dataRow
    retestRange = colLetterRetestStart & dataRow & ":" & colLetterRetestEnd & dataRow
    hasRetest = "COUNTA(" & retestRange & ")>0"

    Select Case method
        Case RT_METHOD_PASS_SCORE
            ' 合格点方式：
            '   本試 >= 合格点 → 本試の得点
            '   MAX(追試) >= 合格点 → 合格点
            '   それ以外 → MAX(本試, 追試全て)
            paramValue = ws.Range(RNG_RT_PASS_SCORE).value
            formula = "=IF(" & origCell & "=""-"",""-""," & _
                      "IF(" & origCell & ">=" & paramValue & "," & origCell & "," & _
                      "IF(AND(" & hasRetest & ",MAX(" & retestRange & ")>=" & paramValue & ")," & _
                      paramValue & "," & _
                      "IF(" & hasRetest & "," & _
                      "MAX(" & origCell & "," & retestRange & ")," & _
                      origCell & "))))"

        Case RT_METHOD_MAX
            ' 最大値：MAX(本試, 追試1, 追試2, ...)
            formula = "=IF(" & origCell & "=""-"",""-""," & _
                      "IF(" & hasRetest & "," & _
                      "MAX(" & origCell & "," & retestRange & ")," & _
                      origCell & "))"

        Case RT_METHOD_AVERAGE
            ' 平均値：本試と追試全ての平均
            formula = "=IF(" & origCell & "=""-"",""-""," & _
                      "IF(" & hasRetest & "," & _
                      "ROUND(AVERAGE(" & origCell & "," & retestRange & "),1)," & _
                      origCell & "))"

        Case RT_METHOD_MEDIAN
            ' 中央値：本試と追試全ての中央値
            formula = "=IF(" & origCell & "=""-"",""-""," & _
                      "IF(" & hasRetest & "," & _
                      "MEDIAN(" & origCell & "," & retestRange & ")," & _
                      origCell & "))"

        Case RT_METHOD_INTERPOLATION
            ' 内分点：α × MAX(本試,追試) + (1-α) × 本試
            paramValue = ws.Range(RNG_RT_PARAM).value
            formula = "=IF(" & origCell & "=""-"",""-""," & _
                      "IF(" & hasRetest & "," & _
                      "ROUND(" & paramValue & "*MAX(" & origCell & "," & retestRange & ")+" & _
                      (1 - paramValue) & "*" & origCell & ",1)," & _
                      origCell & "))"

        Case RT_METHOD_ORIGINAL_ONLY
            ' 本試のみ：追試結果を無視
            formula = "=IF(" & origCell & "=""-"",""-""," & _
                      origCell & ")"

        Case Else
            ' デフォルト：本試のみ
            formula = "=IF(" & origCell & "=""-"",""-""," & _
                      origCell & ")"
    End Select

    ws.Cells(dataRow, finalCol).formula = formula
End Sub

'===============================================================================
' 合格者数・未合格者数の数式を設定
' 引数：
'   ws - 追試シート
'   finalCol - 最終列の列番号
'   childCount - 児童数
' 説明：
'   合格点（E4）が入力されている場合のみ、H3/H4に数式を設定する
'   合格点が空の場合はH3/H4をクリアする
'===============================================================================
Private Sub SetPassFailCountFormulas(ByVal ws As Worksheet, ByVal finalCol As Long, _
                                     ByVal childCount As Long)
    Dim passScore As Variant
    passScore = ws.Range(RNG_RT_PASS_SCORE).value

    ' 合格点が未設定の場合はクリア
    If Trim(passScore & "") = "" Or Not IsNumeric(passScore) Then
        ws.Range("H3").value = ""
        ws.Range("H4").value = ""
        Exit Sub
    End If

    Dim colLetterFinal As String
    Dim lastDataRow As Long
    colLetterFinal = PostingModule.ColumnIndexToLetter(finalCol)
    lastDataRow = RT_DATA_START_ROW + childCount - 1

    Dim finalRange As String
    finalRange = colLetterFinal & RT_DATA_START_ROW & ":" & colLetterFinal & lastDataRow

    ' H3: 合格者数 = 最終列で合格点以上の数値セル数
    ' COUNTIF(範囲, ">="&E4) で合格点セルを参照し、合格点変更に自動追従
    ws.Range("H3").formula = "=IF(" & RNG_RT_PASS_SCORE & "="""",""""," & _
        "COUNTIF(" & finalRange & ","">=""&" & RNG_RT_PASS_SCORE & "))"

    ' H4: 未合格者数 = 最終列で合格点未満の数値セル数（"-"は数値でないため自動除外）
    ws.Range("H4").formula = "=IF(" & RNG_RT_PASS_SCORE & "="""",""""," & _
        "COUNTIF(" & finalRange & ",""<""&" & RNG_RT_PASS_SCORE & "))"
End Sub

'===============================================================================
' 追試回を追加
' 引数：ws - 追試シート
' 説明：
'   追試シート上で呼び出す（ボタン or 手動実行）
'   新しい追試列を挿入し、最終列の数式を更新する
'===============================================================================
Public Sub AddRetestRound(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    Dim finalCol As Long
    Dim newRetestCol As Long
    Dim childCount As Long
    Dim i As Long
    Dim method As String
    Dim retestNum As Long

    ' 状態チェック
    If ws.Range(RNG_RT_STATUS).value <> "追試中" Then
        Call ErrorHandlerModule.ShowValidationError( _
            "このテストは既に完了しています。追試回の追加はできません。")
        Exit Sub
    End If

    ' 最終列の位置を検索（ヘッダー行で「最終」を探す）
    finalCol = 0
    Dim col As Long
    For col = RT_COL_RETEST_START To ws.Cells(RT_HEADER_ROW, Columns.count).End(xlToLeft).Column
        If ws.Cells(RT_HEADER_ROW, col).value = "最終" Then
            finalCol = col
            Exit For
        End If
    Next col

    If finalCol = 0 Then
        Call ErrorHandlerModule.ShowValidationError("最終列が見つかりません。シート構造が不正です。")
        Exit Sub
    End If

    ' 新しい追試列を最終列の手前に挿入
    ws.Columns(finalCol).Insert Shift:=xlToRight
    newRetestCol = finalCol
    finalCol = finalCol + 1  ' 最終列が1つ右にずれる

    ' 追試番号を設定
    retestNum = newRetestCol - RT_COL_RETEST_START + 1
    ws.Cells(RT_HEADER_ROW, newRetestCol).value = "追試" & retestNum

    ' 児童数を取得
    childCount = 0
    For i = RT_DATA_START_ROW To ws.Cells(Rows.count, RT_COL_CODE).End(xlUp).Row
        If Trim(ws.Cells(i, RT_COL_CODE).value & "") <> "" Then
            childCount = childCount + 1
        End If
    Next i

    ' 最終列の数式を更新（範囲が広がったため）
    method = ws.Range(RNG_RT_METHOD).value
    For i = 1 To childCount
        Call SetFinalScoreFormula(ws, RT_DATA_START_ROW + i - 1, _
                                  RT_COL_ORIGINAL, RT_COL_RETEST_START, finalCol, method)
    Next i

    ' 合格者数・未合格者数の数式を更新（最終列がずれたため）
    Call SetPassFailCountFormulas(ws, finalCol, childCount)

    Call ErrorHandlerModule.ShowSuccess("追試" & retestNum & " の列を追加しました。")

    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "AddRetestRound")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 追試を完了し、本体ファイルに結果を反映
' 引数：ws - 追試シート
' 説明：
'   1. 全児童の最終列に値があるか確認
'   2. 本体ファイルのデータシートの該当列の "N" を上書き
'   3. 状態を「反映済み」に変更
'   4. MENUシートの状態も更新
'===============================================================================
Public Sub CompleteRetest(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    Dim parentKey As String
    Dim childCount As Long
    Dim i As Long
    Dim finalCol As Long
    Dim targetCol As Long
    Dim finalScore As Variant
    Dim unresolvedCount As Long
    Dim mainWb As Workbook

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' 状態チェック
    If ws.Range(RNG_RT_STATUS).value = "反映済み" Then
        Call ErrorHandlerModule.ShowInfo("このテストは既に本体ファイルに反映済みです。")
        GoTo CleanExit
    End If

    ' 算出方法が設定されているか確認
    Dim method As String
    method = Trim(ws.Range(RNG_RT_METHOD).value & "")
    If method = "" Then
        Call ErrorHandlerModule.ShowValidationError( _
            "算出方法が設定されていません。" & vbCrLf & _
            "セル " & RNG_RT_METHOD & " で算出方法を選択してください。")
        GoTo CleanExit
    End If

    ' 有効な算出方法かチェック
    If Not IsValidRetestMethod(method) Then
        Call ErrorHandlerModule.ShowValidationError( _
            "算出方法「" & method & "」は無効です。" & vbCrLf & _
            "「最終得点計算」ボタンから算出方法を設定してください。")
        GoTo CleanExit
    End If

    ' 内分点の場合、α値チェック
    Dim paramVal As Variant
    If method = RT_METHOD_INTERPOLATION Then
        paramVal = ws.Range(RNG_RT_PARAM).value
        If Trim(paramVal & "") = "" Or Not IsNumeric(paramVal) Then
            Call ErrorHandlerModule.ShowValidationError( _
                "内分点方式ではα値（0～1）が必要です。" & vbCrLf & _
                "「最終得点計算」ボタンから再設定してください。")
            GoTo CleanExit
        End If
        If CDbl(paramVal) < 0 Or CDbl(paramVal) > 1 Then
            Call ErrorHandlerModule.ShowValidationError( _
                "α値は0～1の範囲で入力してください。" & vbCrLf & _
                "「最終得点計算」ボタンから再設定してください。")
            GoTo CleanExit
        End If
    End If

    ' 合格点方式の場合、合格点チェック
    If method = RT_METHOD_PASS_SCORE Then
        paramVal = ws.Range(RNG_RT_PASS_SCORE).value
        If Trim(paramVal & "") = "" Or Not IsNumeric(paramVal) Then
            Call ErrorHandlerModule.ShowValidationError( _
                "合格点方式では合格点（セル " & RNG_RT_PASS_SCORE & "）が必要です。")
            GoTo CleanExit
        End If
        If CDbl(paramVal) <= 0 Then
            Call ErrorHandlerModule.ShowValidationError( _
                "合格点は0より大きい値を入力してください。")
            GoTo CleanExit
        End If
    End If

    parentKey = ws.Range(RNG_RT_PARENT_KEY).value

    ' 児童数を取得
    childCount = 0
    For i = RT_DATA_START_ROW To ws.Cells(Rows.count, RT_COL_CODE).End(xlUp).Row
        If Trim(ws.Cells(i, RT_COL_CODE).value & "") <> "" Then
            childCount = childCount + 1
        End If
    Next i

    ' 最終列を検索
    finalCol = 0
    Dim col As Long
    For col = RT_COL_RETEST_START To ws.Cells(RT_HEADER_ROW, Columns.count).End(xlToLeft).Column
        If ws.Cells(RT_HEADER_ROW, col).value = "最終" Then
            finalCol = col
            Exit For
        End If
    Next col

    If finalCol = 0 Then
        Call ErrorHandlerModule.ShowValidationError("最終列が見つかりません。")
        GoTo CleanExit
    End If

    ' 最終得点の数式を設定・更新（算出方法に基づく）
    For i = 1 To childCount
        Call SetFinalScoreFormula(ws, RT_DATA_START_ROW + i - 1, _
                                  RT_COL_ORIGINAL, RT_COL_RETEST_START, finalCol, method)
    Next i

    ' 未解決チェック（最終列が空欄 or エラーの児童がいないか）
    unresolvedCount = 0
    For i = 1 To childCount
        finalScore = ws.Cells(RT_DATA_START_ROW + i - 1, finalCol).value
        ' 免除 "-" は解決済みとみなす
        If Trim(finalScore & "") = "" Or IsError(finalScore) Then
            unresolvedCount = unresolvedCount + 1
        End If
    Next i

    If unresolvedCount > 0 Then
        If Not ErrorHandlerModule.ShowConfirmation( _
            "最終得点が未確定の児童が " & unresolvedCount & " 名います。" & vbCrLf & _
            "（追試を受けていない児童は本試の得点が使われます）" & vbCrLf & vbCrLf & _
            "このまま完了しますか？", "未確定の児童がいます") Then
            GoTo CleanExit
        End If
    End If

    ' 確認ダイアログ
    If Not ErrorHandlerModule.ShowConfirmation( _
        "追試結果を本体ファイルに反映します。" & vbCrLf & vbCrLf & _
        "テスト: " & parentKey & " (" & ws.Range(RNG_RT_TEST_NAME).value & ")" & vbCrLf & _
        "算出方法: " & ws.Range(RNG_RT_METHOD).value & vbCrLf & vbCrLf & _
        "よろしいですか？", "追試完了確認") Then
        GoTo CleanExit
    End If

    ' 本体ファイルを取得
    Set mainWb = GetMainWorkbook()
    If mainWb Is Nothing Then
        Call ErrorHandlerModule.ShowValidationError( _
            "本体ファイルが見つかりません。" & vbCrLf & _
            "本体ファイルを開いた状態で再度実行してください。")
        GoTo CleanExit
    End If

    ' 本体ファイルのデータシートで該当列を検索
    Dim mainDataSheet As Worksheet
    Set mainDataSheet = mainWb.sheets("データ")

    targetCol = 0
    Dim lastCol As Long
    lastCol = mainDataSheet.Cells(eRowData.rowKey, Columns.count).End(xlToLeft).Column
    For col = eColData.colDataStart To lastCol
        If mainDataSheet.Cells(eRowData.rowKey, col).value = parentKey Then
            targetCol = col
            Exit For
        End If
    Next col

    If targetCol = 0 Then
        Call ErrorHandlerModule.ShowValidationError( _
            "本体ファイルにテストキー「" & parentKey & "」が見つかりません。")
        GoTo CleanExit
    End If

    ' データシートの保護を一時解除
    On Error Resume Next
    mainDataSheet.Unprotect Password:=SHEET_PROTECT_PASSWORD
    On Error GoTo ErrorHandler

    ' 最終得点を本体ファイルに反映
    For i = 1 To childCount
        finalScore = ws.Cells(RT_DATA_START_ROW + i - 1, finalCol).value

        ' 本体データシートで該当児童の行を特定（コードで照合）
        Dim childCode As String
        childCode = CStr(ws.Cells(RT_DATA_START_ROW + i - 1, RT_COL_CODE).value)

        Dim targetRow As Long
        targetRow = 0
        Dim k As Long
        For k = eRowData.rowChildStart To eRowData.rowChildStart + childCount - 1
            If CStr(mainDataSheet.Cells(k, eColData.colCode).value) = childCode Then
                targetRow = k
                Exit For
            End If
        Next k

        If targetRow > 0 Then
            mainDataSheet.Cells(targetRow, targetCol).value = finalScore
        End If
    Next i

    ' 追試中列のオレンジ色フォーマットをクリア（本体ファイルのモジュールを呼び出す）
    On Error Resume Next
    Application.Run "'" & mainWb.Name & "'!UIFormatModule.ClearRetestColumnFormat", CLng(targetCol)
    On Error GoTo ErrorHandler

    ' 保護を再設定（本体ファイルのモジュールを呼び出す）
    On Error Resume Next
    Application.Run "'" & mainWb.Name & "'!DataManagementModule.ProtectScoreCells"
    On Error GoTo ErrorHandler

    ' 状態を更新
    ws.Range(RNG_RT_STATUS).value = "反映済み"

    ' MENUシートの状態も更新
    Call UpdateRetestMenuStatus(ws.Parent, parentKey, "反映済み")

    mainWb.Save

    Call ErrorHandlerModule.ShowSuccess( _
        "追試結果を本体ファイルに反映しました。" & vbCrLf & _
        "テスト: " & parentKey)

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "CompleteRetest")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 本体ファイルの Workbook を取得
' 説明：
'   追試ファイルから本体ファイルを参照するために使用
'   ファイル名から追試サフィックスを除いた名前で検索する
' 戻り値：本体ファイルの Workbook（見つからなければ Nothing）
'===============================================================================
Public Function GetMainWorkbook() As Workbook
    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim retestFileName As String
    Dim mainFileName As String

    retestFileName = fso.GetBaseName(ThisWorkbook.FullName)

    ' サフィックスを除去して本体ファイル名を推定
    If Right(retestFileName, Len(RETEST_FILE_SUFFIX)) = RETEST_FILE_SUFFIX Then
        mainFileName = Left(retestFileName, Len(retestFileName) - Len(RETEST_FILE_SUFFIX))
    Else
        ' このファイル自体が本体ファイルの場合
        Set GetMainWorkbook = ThisWorkbook
        Set fso = Nothing
        Exit Function
    End If

    ' 開いているワークブックから検索
    Dim wb As Workbook
    For Each wb In Workbooks
        If fso.GetBaseName(wb.Name) = mainFileName Then
            Set GetMainWorkbook = wb
            Set fso = Nothing
            Exit Function
        End If
    Next wb

    ' 見つからない場合、同じフォルダから開く
    Dim mainFilePath As String
    mainFilePath = fso.GetParentFolderName(ThisWorkbook.FullName) & "\" & mainFileName & ".xlsm"

    If fso.FileExists(mainFilePath) Then
        Set GetMainWorkbook = Workbooks.Open(mainFilePath)
    Else
        Set GetMainWorkbook = Nothing
    End If

    Set fso = Nothing
    Exit Function

ErrorHandler:
    Set GetMainWorkbook = Nothing
End Function

'===============================================================================
' 追試MENUシートにエントリを追加
'===============================================================================
Private Sub AddToRetestMenu(ByVal wb As Workbook, ByVal testKey As String, _
                             ByVal subjectName As String, ByVal testName As String, _
                             ByVal perspectiveName As String, ByVal sheetName As String)
    Dim ws As Worksheet
    Dim newRow As Long

    Set ws = wb.sheets("MENU")
    newRow = ws.Cells(Rows.count, RT_MENU_COL_KEY).End(xlUp).Row + 1
    If newRow < RT_MENU_DATA_START_ROW Then newRow = RT_MENU_DATA_START_ROW

    With ws
        .Cells(newRow, RT_MENU_COL_KEY).value = testKey
        .Cells(newRow, RT_MENU_COL_SUBJECT).value = subjectName
        .Cells(newRow, RT_MENU_COL_TESTNAME).value = testName
        .Cells(newRow, RT_MENU_COL_PERSPECTIVE).value = perspectiveName
        .Cells(newRow, RT_MENU_COL_STATUS).value = "追試中"
        .Cells(newRow, RT_MENU_COL_SHEETNAME).value = sheetName

        ' 残り人数の数式
        .Cells(newRow, RT_MENU_COL_REMAINING).formula = _
            "=COUNTBLANK(INDIRECT(""'" & sheetName & "'!E" & RT_DATA_START_ROW & _
            ":E" & (RT_DATA_START_ROW + MAX_CHILDREN - 1) & """))"
    End With
End Sub

'===============================================================================
' 追試MENUシートの状態を更新
'===============================================================================
Public Sub UpdateRetestMenuStatus(ByVal wb As Workbook, ByVal testKey As String, _
                                   ByVal newStatus As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long

    Set ws = wb.sheets("MENU")
    lastRow = ws.Cells(Rows.count, RT_MENU_COL_KEY).End(xlUp).Row

    For i = RT_MENU_DATA_START_ROW To lastRow
        If ws.Cells(i, RT_MENU_COL_KEY).value = testKey Then
            ws.Cells(i, RT_MENU_COL_STATUS).value = newStatus
            Exit For
        End If
    Next i
End Sub

'===============================================================================
' MENUからシートへジャンプ
' 引数：rowIndex - MENUシートの行番号
'===============================================================================
Public Sub JumpToRetestSheet(ByVal rowIndex As Long)
    Dim sheetName As String
    Dim ws As Worksheet
    Dim retestWb As Workbook

    Set retestWb = GetOrCreateRetestWorkbook()
    If retestWb Is Nothing Then Exit Sub

    Dim menuWs As Worksheet
    Set menuWs = retestWb.sheets("MENU")

    sheetName = menuWs.Cells(rowIndex, RT_MENU_COL_SHEETNAME).value
    If Trim(sheetName) = "" Then Exit Sub

    On Error Resume Next
    Set ws = retestWb.sheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Activate
        ws.Cells(RT_DATA_START_ROW, RT_COL_RETEST_START).Select
    Else
        Call ErrorHandlerModule.ShowValidationError("シート「" & sheetName & "」が見つかりません。")
    End If
End Sub

'===============================================================================
' 重複しないシート名を取得
'===============================================================================
Private Function GetUniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim testName As String
    Dim counter As Long
    Dim ws As Worksheet
    Dim exists As Boolean

    testName = baseName
    counter = 0

    Do
        exists = False
        On Error Resume Next
        Set ws = wb.sheets(testName)
        On Error GoTo 0
        If Not ws Is Nothing Then
            exists = True
            counter = counter + 1
            testName = Left(baseName, 31 - Len("_" & counter)) & "_" & counter
            Set ws = Nothing
        End If
    Loop While exists

    GetUniqueSheetName = testName
End Function

'===============================================================================
' MENUシートの全更新
' 説明：各追試シートの状態と残り人数を再集計
'===============================================================================
Public Sub RefreshRetestMenu()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim menuWs As Worksheet
    Dim retestWb As Workbook
    Dim i As Long
    Dim lastRow As Long
    Dim sheetName As String

    Set retestWb = GetOrCreateRetestWorkbook()
    If retestWb Is Nothing Then
        Call ErrorHandlerModule.ShowValidationError("追試ファイルを開けませんでした。")
        Exit Sub
    End If

    Set menuWs = retestWb.sheets("MENU")
    lastRow = menuWs.Cells(Rows.count, RT_MENU_COL_KEY).End(xlUp).Row

    For i = RT_MENU_DATA_START_ROW To lastRow
        sheetName = menuWs.Cells(i, RT_MENU_COL_SHEETNAME).value
        If Trim(sheetName) = "" Then GoTo NextRow

        On Error Resume Next
        Set ws = retestWb.sheets(sheetName)
        On Error GoTo ErrorHandler

        If Not ws Is Nothing Then
            menuWs.Cells(i, RT_MENU_COL_STATUS).value = ws.Range(RNG_RT_STATUS).value
        End If

NextRow:
        Set ws = Nothing
    Next i

    Call ErrorHandlerModule.ShowSuccess("MENUを更新しました。")
    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "RefreshRetestMenu")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 追試ファイルを開く
'===============================================================================
Public Sub OpenRetestFile()
    On Error GoTo ErrorHandler

    Dim filePath As String
    filePath = GetRetestFilePath()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then
        Call ErrorHandlerModule.ShowInfo("追試ファイルがまだ作成されていません。" & vbCrLf & _
            "追試ありのテストを登録すると自動的に作成されます。")
        Set fso = Nothing
        Exit Sub
    End If

    Workbooks.Open filePath
    Set fso = Nothing
    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "OpenRetestFile")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 追試回の追加（UI経由 - ActiveSheetに対して実行）
' 説明：追試シート上のボタンから呼び出す。ActiveSheetを対象とする
'===============================================================================
Public Sub AddRetestRoundUI()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' MENUシートでないことを確認
    If ws.Name = "MENU" Then
        Call ErrorHandlerModule.ShowValidationError( _
            "MENUシートでは実行できません。" & vbCrLf & _
            "対象の追試シートを選択してから実行してください。")
        Exit Sub
    End If

    ' 状態チェック
    On Error Resume Next
    Dim statusVal As String
    statusVal = ws.Range(RNG_RT_STATUS).value
    On Error GoTo ErrorHandler

    If statusVal <> "追試中" Then
        Call ErrorHandlerModule.ShowValidationError( _
            "このテストは追試中ではありません。追試回の追加はできません。")
        Exit Sub
    End If

    Call AddRetestRound(ws)

    ws.Parent.Save
    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "AddRetestRoundUI")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 追試完了（UI経由 - ActiveSheetに対して実行）
' 説明：追試シート上のボタンから呼び出す。ActiveSheetを対象とする
'===============================================================================
Public Sub CompleteRetestUI()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' MENUシートでないことを確認
    If ws.Name = "MENU" Then
        Call ErrorHandlerModule.ShowValidationError( _
            "MENUシートでは実行できません。" & vbCrLf & _
            "対象の追試シートを選択してから実行してください。")
        Exit Sub
    End If

    ' 状態チェック
    On Error Resume Next
    Dim statusVal As String
    statusVal = ws.Range(RNG_RT_STATUS).value
    On Error GoTo ErrorHandler

    If statusVal = "反映済み" Then
        Call ErrorHandlerModule.ShowInfo("このテストは既に反映済みです。")
        Exit Sub
    End If

    Call CompleteRetest(ws)

    ws.Parent.Save
    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "CompleteRetestUI")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 最終得点の数式を適用（追試シート上で算出方法設定後に実行）
' 引数：ws - 追試シート
' 説明：
'   算出方法・パラメータの設定後に呼び出すことで、
'   最終列に数式を設定してプレビューできる
'===============================================================================
Public Sub ApplyFinalScoreFormulas(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    Dim method As String
    Dim finalCol As Long
    Dim childCount As Long
    Dim i As Long
    Dim col As Long

    ' 算出方法の確認（フォーム経由で設定済みのはずだが念のためチェック）
    method = Trim(ws.Range(RNG_RT_METHOD).value & "")
    If method = "" Then
        Call ErrorHandlerModule.ShowValidationError( _
            "算出方法が設定されていません。" & vbCrLf & _
            "「最終得点計算」ボタンから算出方法を設定してください。")
        Exit Sub
    End If

    ' 有効な算出方法かチェック
    If Not IsValidRetestMethod(method) Then
        Call ErrorHandlerModule.ShowValidationError( _
            "算出方法「" & method & "」は無効です。" & vbCrLf & _
            "「最終得点計算」ボタンから算出方法を設定してください。")
        Exit Sub
    End If

    ' 最終列を検索
    finalCol = 0
    For col = RT_COL_RETEST_START To ws.Cells(RT_HEADER_ROW, Columns.count).End(xlToLeft).Column
        If ws.Cells(RT_HEADER_ROW, col).value = "最終" Then
            finalCol = col
            Exit For
        End If
    Next col

    If finalCol = 0 Then
        Call ErrorHandlerModule.ShowValidationError("最終列が見つかりません。")
        Exit Sub
    End If

    ' 児童数を取得
    childCount = 0
    For i = RT_DATA_START_ROW To ws.Cells(Rows.count, RT_COL_CODE).End(xlUp).Row
        If Trim(ws.Cells(i, RT_COL_CODE).value & "") <> "" Then
            childCount = childCount + 1
        End If
    Next i

    ' 最終得点の数式を設定
    For i = 1 To childCount
        Call SetFinalScoreFormula(ws, RT_DATA_START_ROW + i - 1, _
                                  RT_COL_ORIGINAL, RT_COL_RETEST_START, finalCol, method)
    Next i

    ' 合格者数・未合格者数の数式を更新
    Call SetPassFailCountFormulas(ws, finalCol, childCount)

    Call ErrorHandlerModule.ShowSuccess("最終得点の数式を設定しました。（算出方法: " & method & "）")

    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "ApplyFinalScoreFormulas")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 最終得点の数式適用（UI経由 - フォームで算出方法を選択してから適用）
' 説明：追試シートの「最終得点計算」ボタンから呼び出す
'       frm_retest_setting フォームを表示し、決定後に数式を適用する
'===============================================================================
Public Sub ApplyFinalScoreFormulasUI()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' MENUシートでないことを確認
    If ws.Name = "MENU" Then
        Call ErrorHandlerModule.ShowValidationError( _
            "MENUシートでは実行できません。" & vbCrLf & _
            "対象の追試シートを選択してから実行してください。")
        Exit Sub
    End If

    ' 状態チェック
    On Error Resume Next
    Dim statusVal As String
    statusVal = ws.Range(RNG_RT_STATUS).value
    On Error GoTo ErrorHandler

    If statusVal = "反映済み" Then
        Call ErrorHandlerModule.ShowInfo("このテストは既に反映済みです。")
        Exit Sub
    End If

    ' フォームを表示して算出方法を選択させる
    Dim frm As frm_retest_setting
    Set frm = New frm_retest_setting
    frm.Show

    ' キャンセルされた場合は何もしない
    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If

    Unload frm

    ' フォームで設定された算出方法に基づいて数式を適用
    Call ApplyFinalScoreFormulas(ws)

    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "ApplyFinalScoreFormulasUI")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 有効な算出方法かどうかを判定
' 引数：method - 算出方法の文字列
' 戻り値：True=有効、False=無効
'===============================================================================
Private Function IsValidRetestMethod(ByVal method As String) As Boolean
    Select Case method
        Case RT_METHOD_PASS_SCORE, RT_METHOD_MAX, RT_METHOD_AVERAGE, _
             RT_METHOD_MEDIAN, RT_METHOD_INTERPOLATION, RT_METHOD_ORIGINAL_ONLY
            IsValidRetestMethod = True
        Case Else
            IsValidRetestMethod = False
    End Select
End Function

'===============================================================================
' 後出し追試: データシートから追試シートを作成
' 説明：テスト登録後に追試を設定する場合に使用。
'       既存のCreateRetestSheetはPosting時（sh_inputから得点取得）に使用するが、
'       この関数はSh_dataから本試得点を取得して追試シートを作成する。
' 引数：targetCol - データシートの対象列番号
'===============================================================================
Public Sub CreateRetestSheetFromData(ByVal targetCol As Long)
    On Error GoTo ErrorHandler

    Call ErrorHandlerModule.BeginProcess

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim j As Long
    Dim childCount As Long
    Dim testKey As String
    Dim sheetName As String

    ' テスト情報の取得
    testKey = Sh_data.Cells(eRowData.rowKey, targetCol).value

    ' 追試ファイルを取得/作成
    Set wb = GetOrCreateRetestWorkbook()
    If wb Is Nothing Then
        Call ErrorHandlerModule.ShowValidationError("追試ファイルの作成に失敗しました。")
        GoTo CleanExit
    End If

    childCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).value

    ' シート名の生成（キー_テスト名_観点、31文字制限）
    sheetName = testKey & "_" & _
                Sh_data.Cells(eRowData.rowTestName, targetCol).value & "_" & _
                Sh_data.Cells(eRowData.rowPerspective, targetCol).value
    If Len(sheetName) > 31 Then
        sheetName = Left(sheetName, 31)
    End If
    sheetName = GetUniqueSheetName(wb, sheetName)

    ' テンプレートからシートをコピー
    Set ws = CopyTemplateSheet("sh_rt_template", wb, sheetName)

    ' ボタンのマクロ参照先を本体ファイルに書き換え
    Call AssignButtonMacros(ws)

    ' テスト情報の書き込み
    With ws
        .Range(RNG_RT_PARENT_KEY).value = testKey
        .Range(RNG_RT_SUBJECT).value = Sh_data.Cells(eRowData.rowSubject, targetCol).value
        .Range(RNG_RT_TEST_NAME).value = Sh_data.Cells(eRowData.rowTestName, targetCol).value
        .Range(RNG_RT_PERSPECTIVE).value = Sh_data.Cells(eRowData.rowPerspective, targetCol).value
        .Range(RNG_RT_DETAIL).value = Sh_data.Cells(eRowData.rowDetail, targetCol).value
        .Range(RNG_RT_ALLOCATE).value = Sh_data.Cells(eRowData.rowAllocationScore, targetCol).value

        ' 児童データの転記（本試の得点はSh_dataから取得）
        For j = 1 To childCount
            .Cells(RT_DATA_START_ROW + j - 1, RT_COL_CODE).value = _
                Sh_data.Cells(eRowData.rowChildStart + j - 1, eColData.colCode).value
            .Cells(RT_DATA_START_ROW + j - 1, RT_COL_LASTNAME).value = _
                Sh_data.Cells(eRowData.rowChildStart + j - 1, eColData.colLastName).value
            .Cells(RT_DATA_START_ROW + j - 1, RT_COL_FIRSTNAME).value = _
                Sh_data.Cells(eRowData.rowChildStart + j - 1, eColData.colFirstName).value

            ' 本試の得点（Sh_dataから取得）
            .Cells(RT_DATA_START_ROW + j - 1, RT_COL_ORIGINAL).value = _
                Sh_data.Cells(eRowData.rowChildStart + j - 1, targetCol).value
        Next j
    End With

    ' 合格者数・未合格者数の数式を設定
    Dim initFinalCol As Long
    initFinalCol = RT_COL_RETEST_START + RT_COL_FINAL_OFFSET
    Call SetPassFailCountFormulas(ws, initFinalCol, childCount)

    ' MENUシートに追加
    Call AddToRetestMenu(wb, testKey, _
        Sh_data.Cells(eRowData.rowSubject, targetCol).value, _
        Sh_data.Cells(eRowData.rowTestName, targetCol).value, _
        Sh_data.Cells(eRowData.rowPerspective, targetCol).value, _
        sheetName)

    ' データシートの得点を"N"マーカーに置換
    Call MarkColumnAsRetestPending(targetCol, childCount)

    wb.Save

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "CreateRetestSheetFromData")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 後出し追試: 単一列の得点を追試中マーカー"N"に置換
' 説明：後出し追試時に、データシートの対象列の児童得点を"N"に置換する。
'       空欄と"-"（免除）はそのまま残す。
' 引数：
'   targetCol  - データシートの対象列番号
'   childCount - 児童数
'===============================================================================
Private Sub MarkColumnAsRetestPending(ByVal targetCol As Long, ByVal childCount As Long)
    Dim j As Long

    ' シート保護を一時解除
    On Error Resume Next
    Sh_data.Unprotect Password:=SHEET_PROTECT_PASSWORD
    On Error GoTo 0

    With Sh_data
        For j = eRowData.rowChildStart To eRowData.rowChildStart + childCount - 1
            Dim cellVal As String
            cellVal = Trim(.Cells(j, targetCol).value & "")
            If cellVal <> "" And cellVal <> "-" Then
                .Cells(j, targetCol).value = RETEST_MARKER
            End If
        Next j
    End With

    ' 追試中列にオレンジ色のフォーマットを適用
    Call UIFormatModule.ApplyRetestColumnFormat(targetCol)

    ' シート再保護
    Call DataManagementModule.ProtectScoreCells
End Sub

'===============================================================================
' 追試ファイル内で指定キーの追試シートが存在するか確認
' 説明：frmTestEditから呼ばれる。追試ファイルを開き、指定テストキーの
'       追試シートが既に存在するかを判定する。
' 引数：testKey - テストキー
' 戻り値：True=存在する、False=存在しない
'===============================================================================
Public Function HasRetestSheetForKey(ByVal testKey As String) As Boolean
    HasRetestSheetForKey = False

    Dim retestWb As Workbook
    Set retestWb = GetOrCreateRetestWorkbook()
    If retestWb Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = FindRetestSheetByKey(retestWb, testKey)
    HasRetestSheetForKey = Not (ws Is Nothing)
End Function

'===============================================================================
' 追試ファイル内で指定キーの追試シートを検索
' 説明：追試ファイル内の全シートをループし、RNG_RT_PARENT_KEY（B3セル）が
'       指定テストキーと一致するシートを返す。
' 引数：wb - 追試ファイルのWorkbook、testKey - テストキー
' 戻り値：見つかったWorksheet、見つからなければNothing
'===============================================================================
Private Function FindRetestSheetByKey(ByVal wb As Workbook, ByVal testKey As String) As Worksheet
    Dim ws As Worksheet
    Set FindRetestSheetByKey = Nothing

    For Each ws In wb.Worksheets
        If ws.Name <> "MENU" Then
            On Error Resume Next
            Dim parentKey As String
            parentKey = ws.Range(RNG_RT_PARENT_KEY).value & ""
            On Error GoTo 0
            If parentKey = testKey Then
                Set FindRetestSheetByKey = ws
                Exit Function
            End If
        End If
    Next ws
End Function


