Attribute VB_Name = "DataManagementModule"
'===============================================================================
' モジュール名: DataManagementModule
' 機能: テストデータの修正・削除機能等
' 目的: ユーザーがデータシートを直接編集することなく、
'       安全にデータを管理できるようにする
'===============================================================================
Option Explicit

'===============================================================================
' テストデータの削除
' 引数:
'   testKey - 削除するテストキー（例: J001）
'   deleteRetestSheet - True: 追試ファイル内の対応シートも同時に削除する
'===============================================================================
Public Sub DeleteTestData(ByVal testKey As String, Optional ByVal deleteRetestSheet As Boolean = False)
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim lastCol As Long
    Dim targetCol As Long
    Dim testName As String
    Dim subjectName As String

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' テストキーで列を検索
    targetCol = FindTestColumn(testKey)

    If targetCol = 0 Then
        Call ErrorHandlerModule.ShowValidationError("テストキー「" & testKey & "」が見つかりません。")
        GoTo CleanExit
    End If

    ' 削除対象の情報を取得
    testName = Sh_data.Cells(eRowData.rowTestName, targetCol).value
    subjectName = Sh_data.Cells(eRowData.rowSubject, targetCol).value

    ' 確認ダイアログ
    ' 追試テスト: frmTestEditで1回目の確認済みなので「再確認」
    ' 通常テスト: ここが1回目の確認
    Dim confirmMsg As String
    Dim confirmTitle As String
    Dim hasRetestSheet As Boolean
    hasRetestSheet = False

    If deleteRetestSheet Then
        confirmMsg = "【再確認】以下のテストデータを削除します。" & vbCrLf & vbCrLf & _
            "キー: " & testKey & vbCrLf & _
            "教科: " & subjectName & vbCrLf & _
            "テスト名: " & testName & vbCrLf & vbCrLf & _
            "※ 追試中のため、追試ファイル内の対応シートも削除されます。" & vbCrLf & vbCrLf & _
            "この操作は取り消せません。削除してよろしいですか？"
        confirmTitle = "削除の再確認（追試含む）"
    Else
        confirmMsg = "以下のテストデータを削除します。" & vbCrLf & vbCrLf & _
            "キー: " & testKey & vbCrLf & _
            "教科: " & subjectName & vbCrLf & _
            "テスト名: " & testName & vbCrLf & vbCrLf & _
            "この操作は取り消せません。削除してよろしいですか？"
        confirmTitle = "削除確認"
    End If

    If Not ErrorHandlerModule.ShowConfirmation(confirmMsg, confirmTitle) Then
        GoTo CleanExit
    End If

    ' 追試ファイル内の対応シートを削除（再確認後に追試ファイルを開く）
    If deleteRetestSheet Then
        hasRetestSheet = RetestModule.HasRetestSheetForKey(testKey)
        If hasRetestSheet Then
            Call DeleteRetestSheetForKey(testKey)
        End If
    End If

    ' シート保護を一時解除
    On Error Resume Next
    Sh_data.Unprotect Password:=SHEET_PROTECT_PASSWORD
    On Error GoTo ErrorHandler

    ' 列を削除
    Sh_data.Columns(targetCol).Delete

    ' シート再保護
    Call ProtectScoreCells

    ' 本体ファイルを前面に戻す（追試ファイルが前面にある場合の対策）
    ThisWorkbook.Activate
    Sh_data.Activate

    ' 完了メッセージ
    Dim successMsg As String
    successMsg = "テストデータを削除しました。"
    If hasRetestSheet Then
        successMsg = successMsg & vbCrLf & "追試ファイルの対応シートも削除しました。"
    End If
    Call ErrorHandlerModule.ShowSuccess(successMsg)

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("DataManagementModule", "DeleteTestData")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 追試ファイル内の対応する追試シートとMENUエントリを削除
' 動作: 追試ファイルが存在し、対象キーのシートがあれば削除する。
'       MENUシートの該当行も削除する。
'       追試ファイルが存在しない場合やシートが見つからない場合は何もしない。
' 引数:
'   testKey - テストキー
'===============================================================================
Private Sub DeleteRetestSheetForKey(ByVal testKey As String)
    On Error Resume Next

    Dim retestFilePath As String
    retestFilePath = RetestModule.GetRetestFilePath()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 追試ファイルが存在しなければ何もしない
    If Not fso.FileExists(retestFilePath) Then
        Set fso = Nothing
        Exit Sub
    End If
    Set fso = Nothing

    ' 追試ファイルを取得（開く）
    Dim retestWb As Workbook
    Dim fileName As String
    fileName = Dir(retestFilePath)

    On Error Resume Next
    Set retestWb = Workbooks(fileName)
    On Error GoTo 0

    If retestWb Is Nothing Then
        On Error Resume Next
        Set retestWb = Workbooks.Open(retestFilePath)
        On Error GoTo 0
    End If

    If retestWb Is Nothing Then Exit Sub

    ' 対象キーの追試シートを検索・削除
    Dim ws As Worksheet
    Dim found As Boolean
    found = False

    For Each ws In retestWb.Worksheets
        If ws.Name <> "MENU" Then
            Dim parentKey As String
            parentKey = ""
            On Error Resume Next
            parentKey = ws.Range(RNG_RT_PARENT_KEY).value & ""
            On Error GoTo 0
            If parentKey = testKey Then
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
                found = True
                Exit For
            End If
        End If
    Next ws

    ' MENUシートの該当行を削除
    Dim menuWs As Worksheet
    Set menuWs = retestWb.sheets("MENU")
    Dim lastRow As Long
    lastRow = menuWs.Cells(Rows.count, RT_MENU_COL_KEY).End(xlUp).Row

    Dim j As Long
    For j = lastRow To RT_MENU_DATA_START_ROW Step -1
        If menuWs.Cells(j, RT_MENU_COL_KEY).value = testKey Then
            menuWs.Rows(j).Delete
        End If
    Next j

    retestWb.Save
End Sub

'===============================================================================
' テストデータのヘッダー情報を修正
' 引数:
'   testKey - 修正するテストキー
'   fieldType - 修正するフィールド（"テスト名", "観点", "詳細", "配点"）
'   newValue - 新しい値
'===============================================================================
Public Sub UpdateTestHeader(ByVal testKey As String, ByVal fieldType As String, ByVal newValue As Variant)
    On Error GoTo ErrorHandler

    Dim targetCol As Long
    Dim targetRow As Long

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' テストキーで列を検索
    targetCol = FindTestColumn(testKey)

    If targetCol = 0 Then
        Call ErrorHandlerModule.ShowValidationError("テストキー「" & testKey & "」が見つかりません。")
        GoTo CleanExit
    End If

    ' フィールドタイプに応じた行番号を取得
    Select Case fieldType
        Case "テスト名"
            targetRow = eRowData.rowTestName
        Case "観点"
            targetRow = eRowData.rowPerspective
        Case "詳細"
            targetRow = eRowData.rowDetail
        Case "配点"
            targetRow = eRowData.rowAllocationScore
            ' 配点の場合は検証
            If Not IsNumeric(newValue) Then
                Call ErrorHandlerModule.ShowValidationError("配点には数値を入力してください。")
                GoTo CleanExit
            End If
            If CDbl(newValue) <= 0 Then
                Call ErrorHandlerModule.ShowValidationError("配点は0より大きい値を入力してください。")
                GoTo CleanExit
            End If
        Case "実施日"
            targetRow = eRowData.rowTestDate
            If Not IsDate(newValue) Then
                Call ErrorHandlerModule.ShowValidationError("日付の形式が正しくありません。")
                GoTo CleanExit
            End If
        Case Else
            Call ErrorHandlerModule.ShowValidationError("不正なフィールドタイプです: " & fieldType)
            GoTo CleanExit
    End Select

    ' 値を更新
    Sh_data.Cells(targetRow, targetCol).value = newValue

    Call ErrorHandlerModule.ShowSuccess(fieldType & "を更新しました。")

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("DataManagementModule", "UpdateTestHeader")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' 児童の得点を修正
' 引数:
'   testKey - テストキー
'   childCode - 児童コード
'   newScore - 新しい得点（数値または"-"）
'===============================================================================
Public Sub UpdateChildScore(ByVal testKey As String, ByVal childCode As String, ByVal newScore As Variant)
    On Error GoTo ErrorHandler

    Dim targetCol As Long
    Dim targetRow As Long
    Dim allocateScore As Double

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' テストキーで列を検索
    targetCol = FindTestColumn(testKey)
    If targetCol = 0 Then
        Call ErrorHandlerModule.ShowValidationError("テストキー「" & testKey & "」が見つかりません。")
        GoTo CleanExit
    End If

    ' 児童コードで行を検索
    targetRow = FindChildRow(childCode)
    If targetRow = 0 Then
        Call ErrorHandlerModule.ShowValidationError("児童コード「" & childCode & "」が見つかりません。")
        GoTo CleanExit
    End If

    ' 入力値の検証
    If Trim(newScore & "") <> "" And newScore <> "-" Then
        If Not IsNumeric(newScore) Then
            Call ErrorHandlerModule.ShowValidationError("得点には数値または「-」（免除）を入力してください。")
            GoTo CleanExit
        End If

        If CDbl(newScore) < 0 Then
            Call ErrorHandlerModule.ShowValidationError("得点に負の値は入力できません。")
            GoTo CleanExit
        End If

        ' 配点を超えていないかチェック
        allocateScore = Sh_data.Cells(eRowData.rowAllocationScore, targetCol).value
        If CDbl(newScore) > allocateScore Then
            Call ErrorHandlerModule.ShowValidationError( _
                "得点が配点を超えています。" & vbCrLf & _
                "得点: " & newScore & " / 配点: " & allocateScore)
            GoTo CleanExit
        End If
    End If

    ' 値を更新
    Sh_data.Cells(targetRow, targetCol).value = newScore

    Call ErrorHandlerModule.ShowSuccess("得点を更新しました。")

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("DataManagementModule", "UpdateChildScore")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' テストデータ一覧を取得
' 戻り値: テストキー、教科、テスト名、実施日を含む2次元配列
'===============================================================================
Public Function GetTestList() As Variant
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim lastCol As Long
    Dim testCount As Long
    Dim result() As Variant

    With Sh_data
        lastCol = .Cells(eRowData.rowKey, Columns.count).End(xlToLeft).Column

        If lastCol < eColData.colDataStart Then
            GetTestList = Array()
            Exit Function
        End If

        testCount = lastCol - eColData.colDataStart + 1
        ReDim result(1 To testCount, 1 To 5)

        For i = 1 To testCount
            result(i, 1) = .Cells(eRowData.rowKey, eColData.colDataStart + i - 1).value       ' キー
            result(i, 2) = .Cells(eRowData.rowSubject, eColData.colDataStart + i - 1).value   ' 教科
            result(i, 3) = .Cells(eRowData.rowTestName, eColData.colDataStart + i - 1).value  ' テスト名
            result(i, 4) = .Cells(eRowData.rowTestDate, eColData.colDataStart + i - 1).value  ' 実施日
            result(i, 5) = .Cells(eRowData.rowPerspective, eColData.colDataStart + i - 1).value ' 観点
        Next i
    End With

    GetTestList = result
    Exit Function

ErrorHandler:
    GetTestList = Array()
End Function

'===============================================================================
' テストの詳細情報を取得
' 引数:
'   testKey - テストキー
' 戻り値: テスト情報を含むDictionary（Nothing if not found）
'===============================================================================
Public Function GetTestDetails(ByVal testKey As String) As Object
    On Error GoTo ErrorHandler

    Dim targetCol As Long
    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")

    targetCol = FindTestColumn(testKey)
    If targetCol = 0 Then
        Set GetTestDetails = Nothing
        Exit Function
    End If

    With Sh_data
        dict.Add "Key", .Cells(eRowData.rowKey, targetCol).value
        dict.Add "Subject", .Cells(eRowData.rowSubject, targetCol).value
        dict.Add "TestName", .Cells(eRowData.rowTestName, targetCol).value
        dict.Add "TestDate", .Cells(eRowData.rowTestDate, targetCol).value
        dict.Add "Category", .Cells(eRowData.rowCategory, targetCol).value
        dict.Add "Perspective", .Cells(eRowData.rowPerspective, targetCol).value
        dict.Add "Detail", .Cells(eRowData.rowDetail, targetCol).value
        dict.Add "AllocateScore", .Cells(eRowData.rowAllocationScore, targetCol).value
        dict.Add "Average", .Cells(eRowData.rowAverage, targetCol).value
        dict.Add "Median", .Cells(eRowData.rowMedian, targetCol).value
        dict.Add "StdDev", .Cells(eRowData.rowStdDev, targetCol).value
    End With

    Set GetTestDetails = dict
    Exit Function

ErrorHandler:
    Set GetTestDetails = Nothing
End Function

'===============================================================================
' テストキーで列番号を検索する（内部関数）
'===============================================================================
Private Function FindTestColumn(ByVal testKey As String) As Long
    Dim i As Long
    Dim lastCol As Long

    FindTestColumn = 0

    With Sh_data
        lastCol = .Cells(eRowData.rowKey, Columns.count).End(xlToLeft).Column

        For i = eColData.colDataStart To lastCol
            If CStr(.Cells(eRowData.rowKey, i).value) = testKey Then
                FindTestColumn = i
                Exit Function
            End If
        Next i
    End With
End Function

'===============================================================================
' 児童コードで行番号を検索する（内部関数）
'===============================================================================
Private Function FindChildRow(ByVal childCode As String) As Long
    Dim i As Long
    Dim lastRow As Long

    FindChildRow = 0

    With Sh_data
        lastRow = .Cells(Rows.count, eColData.colCode).End(xlUp).Row

        For i = eRowData.rowChildStart To lastRow
            If CStr(.Cells(i, eColData.colCode).value) = childCode Then
                FindChildRow = i
                Exit Function
            End If
        Next i
    End With
End Function

'===============================================================================
' データシートの保護
' 目的: ユーザーが直接データシートを編集できないように保護
'===============================================================================
Public Sub ProtectDataSheet()
    On Error Resume Next

    ' 既存の保護を解除
    Sh_data.Unprotect Password:=SHEET_PROTECT_PASSWORD

    ' 保護設定（VBAからの操作は制限しない）
    Sh_data.Protect Password:=SHEET_PROTECT_PASSWORD, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingCells:=False, AllowFormattingColumns:=False, _
        AllowFormattingRows:=False, AllowInsertingColumns:=False, _
        AllowInsertingRows:=False, AllowInsertingHyperlinks:=False, _
        AllowDeletingColumns:=False, AllowDeletingRows:=False, _
        AllowSorting:=False, AllowFiltering:=False, _
        AllowUsingPivotTables:=False, _
        UserInterfaceOnly:=True

    On Error GoTo 0
End Sub

'===============================================================================
' データシートの保護を解除（内部処理用）
'===============================================================================
Public Sub UnprotectDataSheet()
    On Error Resume Next
    Sh_data.Unprotect Password:=SHEET_PROTECT_PASSWORD
    On Error GoTo 0
End Sub

'===============================================================================
' 得点セルのみを保護
' 目的: 得点セル（23行目以降、D列以降）のみをロックし、シートを保護
'       ダブルクリックでフォームから修正可能
'===============================================================================
Public Sub ProtectScoreCells()
    On Error Resume Next

    Dim lastRow As Long
    Dim lastCol As Long

    With Sh_data
        ' 既存の保護を解除
        .Unprotect Password:=SHEET_PROTECT_PASSWORD

        ' 一旦すべてのセルのロックを解除
        .Cells.Locked = False

        ' データ範囲を取得
        lastRow = .Cells(Rows.count, eColData.colCode).End(xlUp).Row
        lastCol = .Cells(eRowData.rowKey, Columns.count).End(xlToLeft).Column

        If lastCol >= eColData.colDataStart And lastRow >= eRowData.rowChildStart Then
            ' 得点セル（児童データ開始行以降、データ開始列以降）をロック
            .Range(.Cells(eRowData.rowChildStart, eColData.colDataStart), _
                   .Cells(lastRow, lastCol)).Locked = True
        End If

        ' シートを保護（VBAからの操作は制限しない）
        .Protect Password:=SHEET_PROTECT_PASSWORD, _
            DrawingObjects:=True, Contents:=True, Scenarios:=False, _
            AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingColumns:=False, _
            AllowInsertingRows:=False, AllowDeletingColumns:=False, _
            AllowDeletingRows:=False, AllowSorting:=False, AllowFiltering:=False, _
            UserInterfaceOnly:=True
    End With

    On Error GoTo 0
End Sub

'===============================================================================
' 得点セルの保護を解除
'===============================================================================
Public Sub UnprotectScoreCells()
    On Error Resume Next
    Sh_data.Unprotect Password:=SHEET_PROTECT_PASSWORD
    On Error GoTo 0
End Sub

'===============================================================================
' データのエクスポート（CSV形式）
' 引数:
'   filePath - 出力ファイルパス
'===============================================================================
Public Sub ExportToCSV(ByVal filePath As String)
    On Error GoTo ErrorHandler

    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long
    Dim fileNum As Integer
    Dim lineData As String
    Dim cellValue As Variant

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    With Sh_data
        lastRow = .Cells(Rows.count, eColData.colCode).End(xlUp).Row
        lastCol = .Cells(eRowData.rowKey, Columns.count).End(xlToLeft).Column

        If lastCol < eColData.colDataStart Then
            Call ErrorHandlerModule.ShowInfo("エクスポートするデータがありません。")
            GoTo CleanExit
        End If

        fileNum = FreeFile
        Open filePath For Output As #fileNum

        ' ヘッダー行の出力
        lineData = ""
        For j = eColData.colCode To lastCol
            cellValue = .Cells(eRowData.rowKey, j).value
            If j > eColData.colCode Then lineData = lineData & ","
            lineData = lineData & EscapeCSV(CStr(cellValue))
        Next j
        Print #fileNum, lineData

        ' データ行の出力
        For i = eRowData.rowChildStart To lastRow
            lineData = ""
            For j = eColData.colCode To lastCol
                cellValue = .Cells(i, j).value
                If j > eColData.colCode Then lineData = lineData & ","
                lineData = lineData & EscapeCSV(CStr(cellValue))
            Next j
            Print #fileNum, lineData
        Next i

        Close #fileNum
    End With

    Call ErrorHandlerModule.ShowSuccess("データをエクスポートしました。" & vbCrLf & filePath)

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0

    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("DataManagementModule", "ExportToCSV")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' CSV用エスケープ処理
'===============================================================================
Private Function EscapeCSV(ByVal value As String) As String
    If InStr(value, ",") > 0 Or InStr(value, """") > 0 Or InStr(value, vbCrLf) > 0 Then
        EscapeCSV = """" & Replace(value, """", """""") & """"
    Else
        EscapeCSV = value
    End If
End Function

'===============================================================================
' 完全初期化
' 全テストデータ・集計・評価結果を一括クリアし、新しい評価期間を開始できる状態にする。
' 児童名簿・設定シート（教科/観点/カテゴリ/閾値）は保持される。
'===============================================================================
Public Sub CompleteReset()
    On Error GoTo ErrorHandler

    ' --- 二重確認 ---
    If Not ErrorHandlerModule.ShowConfirmation( _
        "すべてのテストデータ・集計・評価結果を初期化します。" & vbCrLf & _
        "（児童名簿と設定は保持されます）" & vbCrLf & vbCrLf & _
        "実行しますか？", "完全初期化") Then
        Exit Sub
    End If

    If Not ErrorHandlerModule.ShowConfirmation( _
        "本当に実行しますか？" & vbCrLf & _
        "この操作は取り消せません。", "最終確認") Then
        Exit Sub
    End If

    Call ErrorHandlerModule.BeginProcess

    ' --- 1. データシート (Sh_data) クリア ---
    Call ClearDataSheet

    ' --- 2. Subjectシート (sh_subject) クリア ---
    Call ClearSubjectSheet

    ' --- 3. Resultシート (sh_result) クリア ---
    Call ClearResultSheet

    ' --- 4. MENUシート (sh_MENU) クリア ---
    Call ClearMenuSheet

    ' --- 5. 入力シート (sh_input) クリア ---
    Call ClearInputSheet

    ' --- 6. IndividualAnalysisシート (sh_individual) クリア ---
    Call ClearIndividualSheet

    ' --- 7. Settingシート キーカウンターリセット ---
    Call ResetKeyCounters

    ' --- 8. 追試ファイルの確認 ---
    Call CheckRetestFile

    Call ErrorHandlerModule.EndProcess
    Call ErrorHandlerModule.ShowSuccess("完全初期化が完了しました。" & vbCrLf & _
        "新しい評価期間のデータ入力を開始できます。")
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("DataManagementModule", "CompleteReset")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'-------------------------------------------------------------------------------
' 完全初期化: データシートクリア
'-------------------------------------------------------------------------------
Private Sub ClearDataSheet()
    Dim lastCol As Long

    On Error Resume Next
    Sh_data.Unprotect Password:=SHEET_PROTECT_PASSWORD
    On Error GoTo 0

    With Sh_data
        lastCol = .Cells(eRowData.rowKey, Columns.count).End(xlToLeft).Column
        If lastCol >= eColData.colDataStart Then
            .Range(.Columns(eColData.colDataStart), .Columns(lastCol)).Delete
        End If
    End With

    Call ProtectScoreCells
End Sub

'-------------------------------------------------------------------------------
' 完全初期化: Subjectシートクリア
'-------------------------------------------------------------------------------
Private Sub ClearSubjectSheet()
    Dim lastRow As Long
    lastRow = eRowSubject.rowChildStart + MAX_CHILDREN - 1

    sh_subject.Range( _
        sh_subject.Cells(eRowSubject.rowKey, eColData.colDataStart), _
        sh_subject.Cells(lastRow, sh_subject.Columns.count) _
    ).ClearContents

    Call SubjectModule.ResetWeightNormalizedStatus
End Sub

'-------------------------------------------------------------------------------
' 完全初期化: Resultシートクリア＋ヘッダー再生成
'-------------------------------------------------------------------------------
Private Sub ClearResultSheet()
    Dim lastRow As Long
    lastRow = RESULT_DATA_START_ROW + MAX_CHILDREN - 1

    With sh_result
        ' ヘッダー行クリア (8-10行目)
        .Range(.Cells(RESULT_SUBJECT_ROW, RESULT_DATA_START_COL), _
               .Cells(RESULT_LABEL_ROW, 200)).ClearContents

        ' データ行クリア (11行目以降)
        .Range(.Cells(RESULT_DATA_START_ROW, RESULT_DATA_START_COL), _
               .Cells(lastRow, 200)).ClearContents
    End With

    ' ヘッダー再生成（HasResultData = False になったので生成される）
    Call ResultModule.GenerateResultHeaders
End Sub

'-------------------------------------------------------------------------------
' 完全初期化: MENUシートクリア
'-------------------------------------------------------------------------------
Private Sub ClearMenuSheet()
    Dim lastRow As Long
    Dim clearEndRow As Long

    With sh_MENU
        ' 最終行を複数列から検索
        lastRow = eRowMenu.rowStart
        If .Cells(Rows.count, eColMenu.colCode).End(xlUp).Row > lastRow Then
            lastRow = .Cells(Rows.count, eColMenu.colCode).End(xlUp).Row
        End If
        If .Cells(Rows.count, eColMenu.colScore).End(xlUp).Row > lastRow Then
            lastRow = .Cells(Rows.count, eColMenu.colScore).End(xlUp).Row
        End If
        If .Cells(Rows.count, eColMenu.colToCol).End(xlUp).Row > lastRow Then
            lastRow = .Cells(Rows.count, eColMenu.colToCol).End(xlUp).Row
        End If

        If lastRow >= eRowMenu.rowStart Then
            clearEndRow = lastRow + 10
            Dim clearRange As Range
            Set clearRange = .Range(.Cells(eRowMenu.rowStart, eColMenu.colCode), _
                                     .Cells(clearEndRow, eColMenu.colToCol))
            clearRange.ClearContents
            clearRange.Interior.ColorIndex = xlColorIndexNone
            clearRange.Borders.LineStyle = xlLineStyleNone
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' 完全初期化: 入力シートクリア
'-------------------------------------------------------------------------------
Private Sub ClearInputSheet()
    ' 入力フォームクリア（日付以外）
    Call PostingModule.ResetInputForm

    ' 日付もクリア（ResetInputForm では保持されるため別途）
    sh_input.Range(RNG_INPUT_DATE).ClearContents

    ' 在籍フィルターボタンのキャプションをリセット
    On Error Resume Next
    sh_input.Buttons("Btn_enrollment").Caption = "在籍"
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' 完全初期化: IndividualAnalysisシートクリア
'-------------------------------------------------------------------------------
Private Sub ClearIndividualSheet()
    ' Analysis.frm の Private Const に対応するローカル定数
    Const IA_HEADER_START_ROW As Long = 9
    Const IA_CHILD_START_ROW As Long = 15
    Const IA_DET1_COL_START As Long = 3     ' C列
    Const IA_DET2_COL_START As Long = 8     ' H列
    Const IA_MAX_SELECTION As Long = 3

    Dim lastRow As Long
    lastRow = IA_CHILD_START_ROW + MAX_CHILDREN - 1

    With sh_individual
        ' グループ1: C-E列
        .Range(.Cells(IA_HEADER_START_ROW, IA_DET1_COL_START), _
               .Cells(lastRow, IA_DET1_COL_START + IA_MAX_SELECTION - 1)).ClearContents

        ' グループ2: H-J列
        .Range(.Cells(IA_HEADER_START_ROW, IA_DET2_COL_START), _
               .Cells(lastRow, IA_DET2_COL_START + IA_MAX_SELECTION - 1)).ClearContents
    End With
End Sub

'-------------------------------------------------------------------------------
' 完全初期化: キーカウンター（Setting C列）リセット
'-------------------------------------------------------------------------------
Private Sub ResetKeyCounters()
    Dim i As Long
    i = SETTING_SUBJECT_START_ROW

    With sh_setting
        Do While .Cells(i, SETTING_SUBJECT_COL).value <> ""
            .Cells(i, SETTING_KEY_COUNT_COL).value = 0
            i = i + 1
        Loop
    End With
End Sub

'-------------------------------------------------------------------------------
' 完全初期化: 追試ファイルの存在確認・警告
'-------------------------------------------------------------------------------
Private Sub CheckRetestFile()
    Dim retestPath As String
    retestPath = RetestModule.GetRetestFilePath()

    If Dir(retestPath) <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Call ErrorHandlerModule.ShowInfo( _
            "追試ファイルが残っています。" & vbCrLf & _
            "必要に応じて手動で削除してください。" & vbCrLf & vbCrLf & _
            fso.GetFileName(retestPath))
    End If
End Sub
