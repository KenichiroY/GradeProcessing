'===============================================================================
' モジュール名: DataManagementModule
' 説明: テストデータの修正・削除機能を提供
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
    testName = Sh_data.Cells(eRowData.rowTestName, targetCol).Value
    subjectName = Sh_data.Cells(eRowData.rowSubject, targetCol).Value

    ' 確認ダイアログ
    ' 追試中テスト: frmTestEditで1回目の確認済みなので「再確認」
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
            "この操作は取り消せません。削除してもよろしいですか？"
        confirmTitle = "削除の再確認（追試含む）"
    Else
        confirmMsg = "以下のテストデータを削除します。" & vbCrLf & vbCrLf & _
            "キー: " & testKey & vbCrLf & _
            "教科: " & subjectName & vbCrLf & _
            "テスト名: " & testName & vbCrLf & vbCrLf & _
            "この操作は取り消せません。削除してもよろしいですか？"
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
    Sh_data.Unprotect
    On Error GoTo ErrorHandler

    ' 列を削除
    Sh_data.Columns(targetCol).Delete

    ' シート再保護
    Call ProtectScoreCells

    ' 本体ファイルを前面に戻す（追試ファイルが前面にある場合の対策）
    ThisWorkbook.Activate
    Sh_data.Activate

    ' 成功メッセージ
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
' 説明: 追試ファイルが存在し、対象キーのシートがあれば削除する。
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
            parentKey = ws.Range(RNG_RT_PARENT_KEY).Value & ""
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
    Set menuWs = retestWb.Sheets("MENU")
    Dim lastRow As Long
    lastRow = menuWs.Cells(Rows.Count, RT_MENU_COL_KEY).End(xlUp).Row

    Dim j As Long
    For j = lastRow To RT_MENU_DATA_START_ROW Step -1
        If menuWs.Cells(j, RT_MENU_COL_KEY).Value = testKey Then
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
            Call ErrorHandlerModule.ShowValidationError("不明なフィールドタイプです: " & fieldType)
            GoTo CleanExit
    End Select
    
    ' 値を更新
    Sh_data.Cells(targetRow, targetCol).Value = newValue
    
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
            Call ErrorHandlerModule.ShowValidationError("得点には数値または「-」（欠席）を入力してください。")
            GoTo CleanExit
        End If
        
        If CDbl(newScore) < 0 Then
            Call ErrorHandlerModule.ShowValidationError("得点に負の値は入力できません。")
            GoTo CleanExit
        End If
        
        ' 配点を超えていないかチェック
        allocateScore = Sh_data.Cells(eRowData.rowAllocationScore, targetCol).Value
        If CDbl(newScore) > allocateScore Then
            Call ErrorHandlerModule.ShowValidationError( _
                "得点が配点を超えています。" & vbCrLf & _
                "得点: " & newScore & " / 配点: " & allocateScore)
            GoTo CleanExit
        End If
    End If
    
    ' 値を更新
    Sh_data.Cells(targetRow, targetCol).Value = newScore
    
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
        lastCol = .Cells(eRowData.rowKey, Columns.Count).End(xlToLeft).Column
        
        If lastCol < eColData.colDataStart Then
            GetTestList = Array()
            Exit Function
        End If
        
        testCount = lastCol - eColData.colDataStart + 1
        ReDim result(1 To testCount, 1 To 5)
        
        For i = 1 To testCount
            result(i, 1) = .Cells(eRowData.rowKey, eColData.colDataStart + i - 1).Value       ' キー
            result(i, 2) = .Cells(eRowData.rowSubject, eColData.colDataStart + i - 1).Value   ' 教科
            result(i, 3) = .Cells(eRowData.rowTestName, eColData.colDataStart + i - 1).Value  ' テスト名
            result(i, 4) = .Cells(eRowData.rowTestDate, eColData.colDataStart + i - 1).Value  ' 実施日
            result(i, 5) = .Cells(eRowData.rowPerspective, eColData.colDataStart + i - 1).Value ' 観点
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
        dict.Add "Key", .Cells(eRowData.rowKey, targetCol).Value
        dict.Add "Subject", .Cells(eRowData.rowSubject, targetCol).Value
        dict.Add "TestName", .Cells(eRowData.rowTestName, targetCol).Value
        dict.Add "TestDate", .Cells(eRowData.rowTestDate, targetCol).Value
        dict.Add "Category", .Cells(eRowData.rowCategory, targetCol).Value
        dict.Add "Perspective", .Cells(eRowData.rowPerspective, targetCol).Value
        dict.Add "Detail", .Cells(eRowData.rowDetail, targetCol).Value
        dict.Add "AllocateScore", .Cells(eRowData.rowAllocationScore, targetCol).Value
        dict.Add "Average", .Cells(eRowData.rowAverage, targetCol).Value
        dict.Add "Median", .Cells(eRowData.rowMedian, targetCol).Value
        dict.Add "StdDev", .Cells(eRowData.rowStdDev, targetCol).Value
    End With
    
    Set GetTestDetails = dict
    Exit Function
    
ErrorHandler:
    Set GetTestDetails = Nothing
End Function

'===============================================================================
' テストキーで列番号を検索（内部関数）
'===============================================================================
Private Function FindTestColumn(ByVal testKey As String) As Long
    Dim i As Long
    Dim lastCol As Long
    
    FindTestColumn = 0
    
    With Sh_data
        lastCol = .Cells(eRowData.rowKey, Columns.Count).End(xlToLeft).Column
        
        For i = eColData.colDataStart To lastCol
            If CStr(.Cells(eRowData.rowKey, i).Value) = testKey Then
                FindTestColumn = i
                Exit Function
            End If
        Next i
    End With
End Function

'===============================================================================
' 児童コードで行番号を検索（内部関数）
'===============================================================================
Private Function FindChildRow(ByVal childCode As String) As Long
    Dim i As Long
    Dim lastRow As Long
    
    FindChildRow = 0
    
    With Sh_data
        lastRow = .Cells(Rows.Count, eColData.colCode).End(xlUp).Row
        
        For i = eRowData.rowChildStart To lastRow
            If CStr(.Cells(i, eColData.colCode).Value) = childCode Then
                FindChildRow = i
                Exit Function
            End If
        Next i
    End With
End Function

'===============================================================================
' データシートの保護
' 説明: ユーザーが直接データシートを編集できないように保護
'===============================================================================
Public Sub ProtectDataSheet()
    On Error Resume Next

    ' 既存の保護を解除
    Sh_data.Unprotect

    ' 保護を設定（ユーザーは閲覧のみ可能）
    Sh_data.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingCells:=False, AllowFormattingColumns:=False, _
        AllowFormattingRows:=False, AllowInsertingColumns:=False, _
        AllowInsertingRows:=False, AllowInsertingHyperlinks:=False, _
        AllowDeletingColumns:=False, AllowDeletingRows:=False, _
        AllowSorting:=False, AllowFiltering:=False, _
        AllowUsingPivotTables:=False

    On Error GoTo 0
End Sub

'===============================================================================
' データシートの保護解除（内部処理用）
'===============================================================================
Public Sub UnprotectDataSheet()
    On Error Resume Next
    Sh_data.Unprotect
    On Error GoTo 0
End Sub

'===============================================================================
' 得点セルのみを保護
' 説明: 得点セル（23行目以降、D列以降）のみをロックし、シートを保護
'       ダブルクリックでフォームから修正可能
'===============================================================================
Public Sub ProtectScoreCells()
    On Error Resume Next

    Dim lastRow As Long
    Dim lastCol As Long

    With Sh_data
        ' 既存の保護を解除
        .Unprotect

        ' 一旦すべてのセルのロックを解除
        .Cells.Locked = False

        ' データ範囲を取得
        lastRow = .Cells(Rows.Count, eColData.colCode).End(xlUp).Row
        lastCol = .Cells(eRowData.rowKey, Columns.Count).End(xlToLeft).Column

        If lastCol >= eColData.colDataStart And lastRow >= eRowData.rowChildStart Then
            ' 得点セル（児童データ開始行以降、データ開始列以降）をロック
            .Range(.Cells(eRowData.rowChildStart, eColData.colDataStart), _
                   .Cells(lastRow, lastCol)).Locked = True
        End If

        ' シートを保護（ロックされたセルのみ編集不可）
        .Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, _
            AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingColumns:=False, _
            AllowInsertingRows:=False, AllowDeletingColumns:=False, _
            AllowDeletingRows:=False, AllowSorting:=False, AllowFiltering:=False
    End With

    On Error GoTo 0
End Sub

'===============================================================================
' 得点セルの保護を解除
'===============================================================================
Public Sub UnprotectScoreCells()
    On Error Resume Next
    Sh_data.Unprotect
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
        lastRow = .Cells(Rows.Count, eColData.colCode).End(xlUp).Row
        lastCol = .Cells(eRowData.rowKey, Columns.Count).End(xlToLeft).Column
        
        If lastCol < eColData.colDataStart Then
            Call ErrorHandlerModule.ShowInfo("エクスポートするデータがありません。")
            GoTo CleanExit
        End If
        
        fileNum = FreeFile
        Open filePath For Output As #fileNum
        
        ' ヘッダー行の出力
        lineData = ""
        For j = eColData.colCode To lastCol
            cellValue = .Cells(eRowData.rowKey, j).Value
            If j > eColData.colCode Then lineData = lineData & ","
            lineData = lineData & EscapeCSV(CStr(cellValue))
        Next j
        Print #fileNum, lineData
        
        ' データ行の出力
        For i = eRowData.rowChildStart To lastRow
            lineData = ""
            For j = eColData.colCode To lastCol
                cellValue = .Cells(i, j).Value
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
