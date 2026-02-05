'===============================================================================
' モジュール名: ResultModule
' 説明: Result転記・スナップショット保存機能を提供
' 機能:
'   - SubjectシートからResultシートへのABC評価転記
'   - Subjectシートのスナップショット保存（別ファイル）
'   - Resultシート列見出しの自動生成
' 注意:
'   - sh_resultはResultシートのCodeName（VBAプロパティで設定が必要）
'   - 定数はPublicConstListModuleで定義
'===============================================================================
Option Explicit

'===============================================================================
' Resultシートの列見出しを自動生成
' 説明: Settingシートの教科・観点から見出しを生成
' 呼び出しタイミング: Workbook_Open（既存データがない場合のみ実行）
'===============================================================================
Public Sub GenerateResultHeaders()
    On Error GoTo ErrorHandler

    Dim i As Long, j As Long
    Dim currentCol As Long
    Dim subjectCount As Long
    Dim perspectiveCount As Long
    Dim subjectName As String
    Dim perspectiveName As String

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' 既存データがあれば生成をスキップ（データ保護のため）
    If HasResultData() Then
        GoTo CleanExit
    End If

    ' 教科数と観点数を取得
    subjectCount = Application.WorksheetFunction.CountA(sh_setting.Range("B3:B20"))
    perspectiveCount = Application.WorksheetFunction.CountA(sh_setting.Range("D3:D10"))

    If subjectCount = 0 Or perspectiveCount = 0 Then
        GoTo CleanExit
    End If

    currentCol = RESULT_DATA_START_COL

    With sh_result
        ' 既存の見出しをクリア
        .Range(.Cells(RESULT_SUBJECT_ROW, RESULT_DATA_START_COL), _
               .Cells(RESULT_LABEL_ROW, 200)).ClearContents

        ' 教科ごとにループ
        For i = 1 To subjectCount
            subjectName = sh_setting.Cells(i + 2, SETTING_SUBJECT_COL).Value

            ' 観点ごとにループ
            For j = 1 To perspectiveCount
                perspectiveName = sh_setting.Cells(j + 2, SETTING_PERSPECTIVE_COL).Value

                ' 教科名（達成率列に配置、2列分結合のため）
                .Cells(RESULT_SUBJECT_ROW, currentCol).Value = subjectName

                ' 観点名
                .Cells(RESULT_PERSPECTIVE_ROW, currentCol).Value = perspectiveName

                ' ラベル（達成率/ABC）
                .Cells(RESULT_LABEL_ROW, currentCol).Value = "達成率"
                .Cells(RESULT_LABEL_ROW, currentCol + 1).Value = "ABC"

                currentCol = currentCol + 2
            Next j
        Next i
    End With

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Debug.Print "GenerateResultHeaders Error: " & Err.Description
End Sub

'===============================================================================
' Resultシートにデータが存在するか確認
' 戻り値: True=データあり、False=データなし
'===============================================================================
Private Function HasResultData() As Boolean
    Dim lastCol As Long
    Dim lastRow As Long

    HasResultData = False

    With sh_result
        ' データ開始列より右に見出しがあるか確認
        lastCol = .Cells(RESULT_SUBJECT_ROW, Columns.Count).End(xlToLeft).Column

        If lastCol < RESULT_DATA_START_COL Then
            Exit Function
        End If

        ' 児童データ行にデータがあるか確認
        lastRow = .Cells(Rows.Count, RESULT_DATA_START_COL).End(xlUp).Row

        If lastRow >= RESULT_DATA_START_ROW Then
            HasResultData = True
        End If
    End With
End Function

'===============================================================================
' SubjectシートからResultシートへABC評価を転記
' 引数:
'   subjectName - 教科名
'   perspectiveName - 観点名
'   ratioCol - Subjectシートの達成率列
'   abcCol - SubjectシートのABC列
'===============================================================================
Public Sub TransferToResult(ByVal subjectName As String, ByVal perspectiveName As String, _
                            ByVal ratioCol As Long, ByVal abcCol As Long)
    On Error GoTo ErrorHandler

    Dim targetCol As Long
    Dim i As Long
    Dim childCount As Long

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' 転記先の列を検索
    targetCol = FindResultColumn(subjectName, perspectiveName)

    If targetCol = 0 Then
        Call ErrorHandlerModule.ShowValidationError( _
            "Resultシートに該当する列が見つかりません。" & vbCrLf & _
            "教科: " & subjectName & vbCrLf & _
            "観点: " & perspectiveName)
        GoTo CleanExit
    End If

    ' 児童数取得
    childCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).Value

    ' 転記
    With sh_result
        For i = 1 To childCount
            ' 達成率
            .Cells(RESULT_DATA_START_ROW + i - 1, targetCol).Value = _
                sh_subject.Cells(eRowSubject.rowChildStart + i - 1, ratioCol).Value
            ' ABC
            .Cells(RESULT_DATA_START_ROW + i - 1, targetCol + 1).Value = _
                sh_subject.Cells(eRowSubject.rowChildStart + i - 1, abcCol).Value
        Next i
    End With

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("ResultModule", "TransferToResult")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' Resultシートで該当する列を検索
' 引数:
'   subjectName - 教科名
'   perspectiveName - 観点名
' 戻り値: 列番号（見つからない場合は0）
'===============================================================================
Private Function FindResultColumn(ByVal subjectName As String, ByVal perspectiveName As String) As Long
    Dim col As Long
    Dim lastCol As Long

    FindResultColumn = 0

    With sh_result
        lastCol = .Cells(RESULT_SUBJECT_ROW, Columns.Count).End(xlToLeft).Column

        For col = RESULT_DATA_START_COL To lastCol
            If .Cells(RESULT_SUBJECT_ROW, col).Value = subjectName And _
               .Cells(RESULT_PERSPECTIVE_ROW, col).Value = perspectiveName Then
                FindResultColumn = col
                Exit Function
            End If
        Next col
    End With
End Function

'===============================================================================
' Subjectシートのスナップショットを保存
' 説明: 現在のSubjectシートを「ファイル名_確定.xlsx」にシートとして追加保存
'       シートには保護をかけ、誤変更を防止
'===============================================================================
Public Sub SaveSubjectSnapshot()
    On Error GoTo ErrorHandler

    Dim fullPath As String
    Dim subjectName As String
    Dim perspectiveName As String
    Dim baseFileName As String
    Dim sheetName As String
    Dim counter As Long
    Dim fso As Object
    Dim targetWb As Workbook
    Dim newSheet As Worksheet
    Dim lastCol As Long
    Dim lastRow As Long
    Dim isNewFile As Boolean
    Dim parentFolder As String

    ' 処理開始
    Call ErrorHandlerModule.BeginProcess

    ' 教科名・観点名を取得
    subjectName = sh_subject.Range(RNG_SUBJECT_SUBJECT).Value
    perspectiveName = GetCurrentPerspectiveName()

    If subjectName = "" Or perspectiveName = "" Then
        Call ErrorHandlerModule.ShowValidationError("教科名または観点名が取得できません。")
        GoTo CleanExit
    End If

    ' FileSystemObject作成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 保存先ファイルパスを生成（ファイル名_確定.xlsx）
    baseFileName = fso.GetBaseName(ThisWorkbook.FullName)
    parentFolder = fso.GetParentFolderName(ThisWorkbook.FullName)
    fullPath = parentFolder & "\" & baseFileName & "_確定.xlsx"

    ' シート名を生成（教科_観点_日付、同名がある場合は連番）
    sheetName = subjectName & "_" & perspectiveName & "_" & Format(Date, "yyyymmdd")
    ' シート名は31文字制限があるため切り詰め
    If Len(sheetName) > 28 Then
        sheetName = Left(sheetName, 28)
    End If

    ' ファイルが存在するか確認
    isNewFile = Not fso.FileExists(fullPath)

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    If isNewFile Then
        ' 新規ファイル作成: Subjectシートをコピーして新規ワークブック作成
        sh_subject.Copy
        Set targetWb = ActiveWorkbook
        Set newSheet = targetWb.Sheets(1)
    Else
        ' 既存ファイルを開く
        Set targetWb = Workbooks.Open(fullPath)

        ' 同名シートがあれば連番を付ける
        counter = 0
        Dim originalSheetName As String
        originalSheetName = sheetName
        Do While SheetExists(targetWb, sheetName)
            counter = counter + 1
            sheetName = originalSheetName & "_" & counter
            ' シート名31文字制限
            If Len(sheetName) > 31 Then
                sheetName = Left(originalSheetName, 31 - Len("_" & counter)) & "_" & counter
            End If
        Loop

        ' Subjectシートをコピーしてこのワークブックに追加
        sh_subject.Copy After:=targetWb.Sheets(targetWb.Sheets.Count)
        Set newSheet = targetWb.Sheets(targetWb.Sheets.Count)
    End If

    ' シート名を設定
    newSheet.Name = sheetName

    With newSheet
        ' フォームボタン・図形・OLEオブジェクトを削除（マクロ実行防止）
        Call DeleteAllControls(newSheet)

        ' シート全体を値に変換（すべての数式を解消）
        .UsedRange.Copy
        .Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

        ' シート保護（パスワードなし）
        .Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True
    End With

    ' 保存
    If isNewFile Then
        targetWb.SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
    Else
        targetWb.Save
    End If

    targetWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Call ErrorHandlerModule.ShowSuccess( _
        "スナップショットを保存しました。" & vbCrLf & vbCrLf & _
        "保存先: " & fullPath & vbCrLf & _
        "シート名: " & sheetName)

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Set fso = Nothing
    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Call ErrorHandlerModule.CleanupOnError

    On Error Resume Next
    If Not targetWb Is Nothing Then
        targetWb.Close SaveChanges:=False
    End If
    On Error GoTo 0

    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("ResultModule", "SaveSubjectSnapshot")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub

'===============================================================================
' ワークブック内に指定した名前のシートが存在するか確認
' 引数:
'   wb - 対象ワークブック
'   sheetName - シート名
' 戻り値: True=存在する、False=存在しない
'===============================================================================
Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    SheetExists = False
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        SheetExists = True
    End If
End Function

'===============================================================================
' シート上のすべてのコントロール・図形を削除
' 説明: フォームボタン、OLEオブジェクト（チェックボックス等）、図形を削除
'       スナップショットからマクロが実行されるのを防止
' 引数:
'   ws - 対象ワークシート
'===============================================================================
Private Sub DeleteAllControls(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim oleObj As OLEObject

    On Error Resume Next

    ' 図形（フォームボタン含む）を削除
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ' OLEオブジェクト（ActiveXコントロール）を削除
    For Each oleObj In ws.OLEObjects
        oleObj.Delete
    Next oleObj

    On Error GoTo 0
End Sub

'===============================================================================
' 現在選択されている観点名を取得
'===============================================================================
Private Function GetCurrentPerspectiveName() As String
    Dim i As Long
    Dim perspectiveName As String

    GetCurrentPerspectiveName = ""

    ' Subjectシートのチェックボックスから取得
    On Error Resume Next
    For i = 1 To MAX_PERSPECTIVES
        If sh_subject.OLEObjects("perspective" & i).Object.Value = True Then
            perspectiveName = sh_subject.OLEObjects("perspective" & i).Object.Caption
            If GetCurrentPerspectiveName = "" Then
                GetCurrentPerspectiveName = perspectiveName
            Else
                ' 複数選択の場合は最初の1つだけ使用
                Exit Function
            End If
        End If
    Next i
    On Error GoTo 0

    ' チェックボックスから取得できない場合、Subjectシートのデータから取得
    If GetCurrentPerspectiveName = "" Then
        If sh_subject.Cells(eRowSubject.rowKey, eColData.colDataStart).Value <> "" Then
            GetCurrentPerspectiveName = sh_subject.Cells(eRowSubject.rowPerspective, eColData.colDataStart).Value
        End If
    End If
End Function

'===============================================================================
' 最終決定時の処理（Result転記 + スナップショット保存）
' 説明: 最終決定列をダブルクリックした時に呼び出される
' 引数:
'   ratioCol - 達成率列
'   abcCol - 最終決定ABC列
'===============================================================================
Public Sub FinalizeEvaluation(ByVal ratioCol As Long, ByVal abcCol As Long)
    On Error GoTo ErrorHandler

    Dim subjectName As String
    Dim perspectiveName As String

    ' 教科名・観点名を取得
    subjectName = sh_subject.Range(RNG_SUBJECT_SUBJECT).Value
    perspectiveName = GetCurrentPerspectiveName()

    ' 確認ダイアログ
    If Not ErrorHandlerModule.ShowConfirmation( _
        "評価を確定してResultシートに転記しますか？" & vbCrLf & vbCrLf & _
        "教科: " & subjectName & vbCrLf & _
        "観点: " & perspectiveName & vbCrLf & vbCrLf & _
        "※スナップショットも自動保存されます。", "最終決定確認") Then
        Exit Sub
    End If

    ' Result転記
    Call TransferToResult(subjectName, perspectiveName, ratioCol, abcCol)

    ' スナップショット保存
    Call SaveSubjectSnapshot

    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("ResultModule", "FinalizeEvaluation")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub
