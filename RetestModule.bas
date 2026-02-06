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
    Set menuWs = CopyTemplateSheet("sh_rt_menu_template", wb, "MENU")

    ' デフォルトのSheet1を削除
    Application.DisplayAlerts = False
    Dim defaultSheet As Worksheet
    For Each defaultSheet In wb.Sheets
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
'   templateCodeName - テンプレートシートのCodeName（sh_rt_menu_template or sh_rt_template）
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
    templateWs.Copy After:=targetWb.Sheets(targetWb.Sheets.Count)

    ' テンプレートを元の状態に戻す
    templateWs.Visible = originalVisible

    ' コピーされたシート（最後のシート）を取得してリネーム
    Set newWs = targetWb.Sheets(targetWb.Sheets.Count)
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

    childCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).Value

    ' テスト数分のシートを作成（追試ONの列のみ）
    For i = 1 To numTest
        If Not retestFlags(i) Then GoTo NextTest

        targetCol = lastColData + i
        testKey = Sh_data.Cells(eRowData.rowKey, targetCol).Value

        ' シート名の生成（キー_テスト名_観点、31文字制限）
        sheetName = testKey & "_" & _
                    Sh_data.Cells(eRowData.rowTestName, targetCol).Value & "_" & _
                    Sh_data.Cells(eRowData.rowPerspective, targetCol).Value
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
            .Range(RNG_RT_PARENT_KEY).Value = testKey
            .Range(RNG_RT_SUBJECT).Value = Sh_data.Cells(eRowData.rowSubject, targetCol).Value
            .Range(RNG_RT_TEST_NAME).Value = Sh_data.Cells(eRowData.rowTestName, targetCol).Value
            .Range(RNG_RT_PERSPECTIVE).Value = Sh_data.Cells(eRowData.rowPerspective, targetCol).Value
            .Range(RNG_RT_DETAIL).Value = Sh_data.Cells(eRowData.rowDetail, targetCol).Value
            .Range(RNG_RT_ALLOCATE).Value = Sh_data.Cells(eRowData.rowAllocationScore, targetCol).Value

            ' 児童データの転記
            For j = 1 To childCount
                .Cells(RT_DATA_START_ROW + j - 1, RT_COL_CODE).Value = _
                    Sh_data.Cells(eRowData.rowChildStart + j - 1, eColData.colCode).Value
                .Cells(RT_DATA_START_ROW + j - 1, RT_COL_LASTNAME).Value = _
                    Sh_data.Cells(eRowData.rowChildStart + j - 1, eColData.colLastName).Value
                .Cells(RT_DATA_START_ROW + j - 1, RT_COL_FIRSTNAME).Value = _
                    Sh_data.Cells(eRowData.rowChildStart + j - 1, eColData.colFirstName).Value

                ' 本試の得点（sh_input から取得）
                Dim originalScore As Variant
                originalScore = sh_input.Cells(eRowInput.rowChildStart + j - 1, _
                                eColInput.colDataStart + i - 1).Value
                .Cells(RT_DATA_START_ROW + j - 1, RT_COL_ORIGINAL).Value = originalScore
            Next j
        End With

        ' MENUシートに追加
        Call AddToRetestMenu(wb, testKey, _
            Sh_data.Cells(eRowData.rowSubject, targetCol).Value, _
            Sh_data.Cells(eRowData.rowTestName, targetCol).Value, _
            Sh_data.Cells(eRowData.rowPerspective, targetCol).Value, _
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

    Select Case method
        Case RT_METHOD_MAX
            ' MAX(本試, 追試1, 追試2, ...)
            formula = "=IF(" & colLetterOrig & dataRow & "=""-"",""-""," & _
                      "IF(COUNTA(" & colLetterRetestStart & dataRow & ":" & colLetterRetestEnd & dataRow & ")=0," & _
                      colLetterOrig & dataRow & "," & _
                      "MAX(" & colLetterOrig & dataRow & "," & _
                      colLetterRetestStart & dataRow & ":" & colLetterRetestEnd & dataRow & ")))"

        Case RT_METHOD_INTERPOLATION
            ' α × MAX(全回) + (1-α) × 本試
            paramValue = ws.Range(RNG_RT_PARAM).Value
            formula = "=IF(" & colLetterOrig & dataRow & "=""-"",""-""," & _
                      "IF(COUNTA(" & colLetterRetestStart & dataRow & ":" & colLetterRetestEnd & dataRow & ")=0," & _
                      colLetterOrig & dataRow & "," & _
                      "ROUND(" & paramValue & "*MAX(" & colLetterOrig & dataRow & "," & _
                      colLetterRetestStart & dataRow & ":" & colLetterRetestEnd & dataRow & ")+" & _
                      (1 - paramValue) & "*" & colLetterOrig & dataRow & ",1)))"

        Case RT_METHOD_CAPPED
            ' MAX(本試, MIN(MAX(追試1,追試2,...), 上限点))
            paramValue = ws.Range(RNG_RT_PARAM).Value
            formula = "=IF(" & colLetterOrig & dataRow & "=""-"",""-""," & _
                      "IF(COUNTA(" & colLetterRetestStart & dataRow & ":" & colLetterRetestEnd & dataRow & ")=0," & _
                      colLetterOrig & dataRow & "," & _
                      "MAX(" & colLetterOrig & dataRow & "," & _
                      "MIN(MAX(" & colLetterRetestStart & dataRow & ":" & _
                      colLetterRetestEnd & dataRow & ")," & paramValue & "))))"

        Case Else
            ' デフォルト：最大値
            formula = "=IF(" & colLetterOrig & dataRow & "=""-"",""-""," & _
                      "IF(COUNTA(" & colLetterRetestStart & dataRow & ":" & colLetterRetestEnd & dataRow & ")=0," & _
                      colLetterOrig & dataRow & "," & _
                      "MAX(" & colLetterOrig & dataRow & "," & _
                      colLetterRetestStart & dataRow & ":" & colLetterRetestEnd & dataRow & ")))"
    End Select

    ws.Cells(dataRow, finalCol).formula = formula
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
    If ws.Range(RNG_RT_STATUS).Value <> "追試中" Then
        Call ErrorHandlerModule.ShowValidationError( _
            "このテストは既に完了しています。追試回の追加はできません。")
        Exit Sub
    End If

    ' 最終列の位置を検索（ヘッダー行で「最終」を探す）
    finalCol = 0
    Dim col As Long
    For col = RT_COL_RETEST_START To ws.Cells(RT_HEADER_ROW, Columns.Count).End(xlToLeft).Column
        If ws.Cells(RT_HEADER_ROW, col).Value = "最終" Then
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
    ws.Cells(RT_HEADER_ROW, newRetestCol).Value = "追試" & retestNum

    ' 児童数を取得
    childCount = 0
    For i = RT_DATA_START_ROW To ws.Cells(Rows.Count, RT_COL_CODE).End(xlUp).Row
        If Trim(ws.Cells(i, RT_COL_CODE).Value & "") <> "" Then
            childCount = childCount + 1
        End If
    Next i

    ' 最終列の数式を更新（範囲が広がったため）
    method = ws.Range(RNG_RT_METHOD).Value
    For i = 1 To childCount
        Call SetFinalScoreFormula(ws, RT_DATA_START_ROW + i - 1, _
                                  RT_COL_ORIGINAL, RT_COL_RETEST_START, finalCol, method)
    Next i

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
    If ws.Range(RNG_RT_STATUS).Value = "反映済み" Then
        Call ErrorHandlerModule.ShowInfo("このテストは既に本体ファイルに反映済みです。")
        GoTo CleanExit
    End If

    ' 算出方法が設定されているか確認
    Dim method As String
    method = Trim(ws.Range(RNG_RT_METHOD).Value & "")
    If method = "" Then
        Call ErrorHandlerModule.ShowValidationError( _
            "算出方法が設定されていません。" & vbCrLf & _
            "セル " & RNG_RT_METHOD & " で算出方法を選択してください。" & vbCrLf & _
            "（「最大値」「内分点」「追試上限付き」のいずれか）")
        GoTo CleanExit
    End If

    ' 有効な算出方法かチェック
    If method <> RT_METHOD_MAX And _
       method <> RT_METHOD_INTERPOLATION And _
       method <> RT_METHOD_CAPPED Then
        Call ErrorHandlerModule.ShowValidationError( _
            "算出方法「" & method & "」は無効です。" & vbCrLf & _
            "「最大値」「内分点」「追試上限付き」から選択してください。")
        GoTo CleanExit
    End If

    ' 内分点・追試上限付きの場合、パラメータチェック
    Dim paramVal As Variant
    paramVal = ws.Range(RNG_RT_PARAM).Value
    If method = RT_METHOD_INTERPOLATION Then
        If Trim(paramVal & "") = "" Or Not IsNumeric(paramVal) Then
            Call ErrorHandlerModule.ShowValidationError( _
                "内分点方式ではパラメータ（α値：0～1）が必要です。" & vbCrLf & _
                "セル " & RNG_RT_PARAM & " に数値を入力してください。")
            GoTo CleanExit
        End If
        If CDbl(paramVal) < 0 Or CDbl(paramVal) > 1 Then
            Call ErrorHandlerModule.ShowValidationError( _
                "内分点のα値は0～1の範囲で入力してください。" & vbCrLf & _
                "（1に近いほど最大値寄り、0に近いほど本試寄り）")
            GoTo CleanExit
        End If
    ElseIf method = RT_METHOD_CAPPED Then
        If Trim(paramVal & "") = "" Or Not IsNumeric(paramVal) Then
            Call ErrorHandlerModule.ShowValidationError( _
                "追試上限付き方式ではパラメータ（上限点）が必要です。" & vbCrLf & _
                "セル " & RNG_RT_PARAM & " に数値を入力してください。")
            GoTo CleanExit
        End If
        If CDbl(paramVal) <= 0 Then
            Call ErrorHandlerModule.ShowValidationError( _
                "上限点は0より大きい値を入力してください。")
            GoTo CleanExit
        End If
    End If

    parentKey = ws.Range(RNG_RT_PARENT_KEY).Value

    ' 児童数を取得
    childCount = 0
    For i = RT_DATA_START_ROW To ws.Cells(Rows.Count, RT_COL_CODE).End(xlUp).Row
        If Trim(ws.Cells(i, RT_COL_CODE).Value & "") <> "" Then
            childCount = childCount + 1
        End If
    Next i

    ' 最終列を検索
    finalCol = 0
    Dim col As Long
    For col = RT_COL_RETEST_START To ws.Cells(RT_HEADER_ROW, Columns.Count).End(xlToLeft).Column
        If ws.Cells(RT_HEADER_ROW, col).Value = "最終" Then
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
        finalScore = ws.Cells(RT_DATA_START_ROW + i - 1, finalCol).Value
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
        "テスト: " & parentKey & " (" & ws.Range(RNG_RT_TEST_NAME).Value & ")" & vbCrLf & _
        "算出方法: " & ws.Range(RNG_RT_METHOD).Value & vbCrLf & vbCrLf & _
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
    Set mainDataSheet = mainWb.Sheets("データ")

    targetCol = 0
    Dim lastCol As Long
    lastCol = mainDataSheet.Cells(eRowData.rowKey, Columns.Count).End(xlToLeft).Column
    For col = eColData.colDataStart To lastCol
        If mainDataSheet.Cells(eRowData.rowKey, col).Value = parentKey Then
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
    mainDataSheet.Unprotect
    On Error GoTo ErrorHandler

    ' 最終得点を本体ファイルに反映
    For i = 1 To childCount
        finalScore = ws.Cells(RT_DATA_START_ROW + i - 1, finalCol).Value

        ' 本体データシートで該当児童の行を特定（コードで照合）
        Dim childCode As String
        childCode = CStr(ws.Cells(RT_DATA_START_ROW + i - 1, RT_COL_CODE).Value)

        Dim targetRow As Long
        targetRow = 0
        Dim k As Long
        For k = eRowData.rowChildStart To eRowData.rowChildStart + childCount - 1
            If CStr(mainDataSheet.Cells(k, eColData.colCode).Value) = childCode Then
                targetRow = k
                Exit For
            End If
        Next k

        If targetRow > 0 Then
            mainDataSheet.Cells(targetRow, targetCol).Value = finalScore
        End If
    Next i

    ' 保護を再設定（本体ファイルのモジュールを呼び出す）
    On Error Resume Next
    Application.Run "'" & mainWb.Name & "'!DataManagementModule.ProtectScoreCells"
    On Error GoTo ErrorHandler

    ' 状態を更新
    ws.Range(RNG_RT_STATUS).Value = "反映済み"

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

    Set ws = wb.Sheets("MENU")
    newRow = ws.Cells(Rows.Count, RT_MENU_COL_KEY).End(xlUp).Row + 1
    If newRow < RT_MENU_DATA_START_ROW Then newRow = RT_MENU_DATA_START_ROW

    With ws
        .Cells(newRow, RT_MENU_COL_KEY).Value = testKey
        .Cells(newRow, RT_MENU_COL_SUBJECT).Value = subjectName
        .Cells(newRow, RT_MENU_COL_TESTNAME).Value = testName
        .Cells(newRow, RT_MENU_COL_PERSPECTIVE).Value = perspectiveName
        .Cells(newRow, RT_MENU_COL_STATUS).Value = "追試中"
        .Cells(newRow, RT_MENU_COL_SHEETNAME).Value = sheetName

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

    Set ws = wb.Sheets("MENU")
    lastRow = ws.Cells(Rows.Count, RT_MENU_COL_KEY).End(xlUp).Row

    For i = RT_MENU_DATA_START_ROW To lastRow
        If ws.Cells(i, RT_MENU_COL_KEY).Value = testKey Then
            ws.Cells(i, RT_MENU_COL_STATUS).Value = newStatus
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
    Set menuWs = retestWb.Sheets("MENU")

    sheetName = menuWs.Cells(rowIndex, RT_MENU_COL_SHEETNAME).Value
    If Trim(sheetName) = "" Then Exit Sub

    On Error Resume Next
    Set ws = retestWb.Sheets(sheetName)
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
        Set ws = wb.Sheets(testName)
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

    Set menuWs = retestWb.Sheets("MENU")
    lastRow = menuWs.Cells(Rows.Count, RT_MENU_COL_KEY).End(xlUp).Row

    For i = RT_MENU_DATA_START_ROW To lastRow
        sheetName = menuWs.Cells(i, RT_MENU_COL_SHEETNAME).Value
        If Trim(sheetName) = "" Then GoTo NextRow

        On Error Resume Next
        Set ws = retestWb.Sheets(sheetName)
        On Error GoTo ErrorHandler

        If Not ws Is Nothing Then
            menuWs.Cells(i, RT_MENU_COL_STATUS).Value = ws.Range(RNG_RT_STATUS).Value
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
    statusVal = ws.Range(RNG_RT_STATUS).Value
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
    statusVal = ws.Range(RNG_RT_STATUS).Value
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

    ' 算出方法の確認
    method = Trim(ws.Range(RNG_RT_METHOD).Value & "")
    If method = "" Then
        Call ErrorHandlerModule.ShowValidationError( _
            "算出方法が設定されていません。" & vbCrLf & _
            "セル " & RNG_RT_METHOD & " で算出方法を選択してください。")
        Exit Sub
    End If

    ' 有効な算出方法かチェック
    If method <> RT_METHOD_MAX And _
       method <> RT_METHOD_INTERPOLATION And _
       method <> RT_METHOD_CAPPED Then
        Call ErrorHandlerModule.ShowValidationError( _
            "算出方法「" & method & "」は無効です。" & vbCrLf & _
            "「最大値」「内分点」「追試上限付き」から選択してください。")
        Exit Sub
    End If

    ' パラメータチェック（内分点・追試上限付きの場合）
    Dim paramVal As Variant
    paramVal = ws.Range(RNG_RT_PARAM).Value
    If method = RT_METHOD_INTERPOLATION Then
        If Trim(paramVal & "") = "" Or Not IsNumeric(paramVal) Then
            Call ErrorHandlerModule.ShowValidationError( _
                "内分点方式ではパラメータ（α値：0～1）が必要です。")
            Exit Sub
        End If
    ElseIf method = RT_METHOD_CAPPED Then
        If Trim(paramVal & "") = "" Or Not IsNumeric(paramVal) Then
            Call ErrorHandlerModule.ShowValidationError( _
                "追試上限付き方式ではパラメータ（上限点）が必要です。")
            Exit Sub
        End If
    End If

    ' 最終列を検索
    finalCol = 0
    For col = RT_COL_RETEST_START To ws.Cells(RT_HEADER_ROW, Columns.Count).End(xlToLeft).Column
        If ws.Cells(RT_HEADER_ROW, col).Value = "最終" Then
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
    For i = RT_DATA_START_ROW To ws.Cells(Rows.Count, RT_COL_CODE).End(xlUp).Row
        If Trim(ws.Cells(i, RT_COL_CODE).Value & "") <> "" Then
            childCount = childCount + 1
        End If
    Next i

    ' 最終得点の数式を設定
    For i = 1 To childCount
        Call SetFinalScoreFormula(ws, RT_DATA_START_ROW + i - 1, _
                                  RT_COL_ORIGINAL, RT_COL_RETEST_START, finalCol, method)
    Next i

    Call ErrorHandlerModule.ShowSuccess("最終得点の数式を設定しました。（算出方法: " & method & "）")

    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "ApplyFinalScoreFormulas")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 最終得点の数式適用（UI経由 - アクティブシートに適用）
' 説明：追試ファイルのシートで算出方法設定後にボタンから実行する
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
    statusVal = ws.Range(RNG_RT_STATUS).Value
    On Error GoTo ErrorHandler

    If statusVal = "反映済み" Then
        Call ErrorHandlerModule.ShowInfo("このテストは既に反映済みです。")
        Exit Sub
    End If

    Call ApplyFinalScoreFormulas(ws)

    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("RetestModule", "ApplyFinalScoreFormulasUI")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub
